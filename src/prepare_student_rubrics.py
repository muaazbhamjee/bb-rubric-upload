#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prepare_student_rubrics.py - combine MIA 320 GA7 rubrics, one per student.

The GA7 individual-project rubrics are scattered across per-TA subfolders (and a
Calibration folder), sometimes with duplicates. This script gathers them into a
single folder, keeps the most complete copy per student, and renames each to the
Blackboard/email convention  feedback_u<studentno>.pdf  so it drops straight into
the existing flatten -> email pipeline.

Runs offline (POPIA-safe). Dependency: pypdf.
    pip install pypdf

PIPELINE POSITION (this is the MIA 320 equivalent of rename_for_blackboard.py):
    prepare_student_rubrics.py  ->  flatten_pdfs.py  ->  email.js

Usage:
    python prepare_student_rubrics.py <input_dir> [output_dir]
                                      [--include-ungraded]

    <input_dir>   folder holding the graded rubrics (per-TA subfolders, etc.)
    [output_dir]  where 'combined/' is created (default: <input_dir>)
    --include-ungraded   also keep rubrics with no marks at all (default: skip)

Output:  <output_dir>/combined/feedback_u<sn>.pdf   (editable originals, deduped)
Then:    python flatten_pdfs.py <output_dir>/combined/   -> combined/flattened/

A rubric is any PDF whose form contains the GA7 fields; other files are ignored.
The student number is read from the rubric's form field (falling back to the
filename), which is more reliable than a filename-only match.
"""

import argparse
import csv
import os
import re
import shutil
import sys

try:
    from pypdf import PdfReader
except ImportError:
    sys.exit("Missing dependency. Run:  pip install pypdf")

CRIT_FIELDS = ["f_m1", "f_m2", "f_m3", "f_m4", "f_m5", "f_m6"]
C_PAIRS = [("f_c1y", "f_c1n"), ("f_c2y", "f_c2n"),
           ("f_c3y", "f_c3n"), ("f_c4y", "f_c4n")]
OVERALL_PAIR = ("f_overall_comp", "f_overall_notyet")
RUBRIC_MARKERS = set(CRIT_FIELDS + ["f_mtotal", "f_core_count", "f_student_no"]
                     + [c for pair in C_PAIRS for c in pair])

COMBINED_DIR = "combined"
STATUS_RANK = {"full": 0, "partial": 1, "ungraded": 2}


def norm_sn(raw):
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"^[uU]", "", s)
    return re.sub(r"\D", "", s)


def is_checked(v):
    return v not in (None, "", "/Off", "Off", False)


def inspect(path):
    """Return (is_rubric, student_no, status, filled) for a PDF."""
    try:
        fields = PdfReader(path).get_fields() or {}
    except Exception:  # noqa: BLE001
        return False, "", "unreadable", 0
    if not (set(fields.keys()) & RUBRIC_MARKERS):
        return False, "", "not-a-rubric", 0

    def tx(n):
        f = fields.get(n)
        return "" if f is None or f.get("/V") is None else str(f.get("/V")).strip()

    def btn(n):
        f = fields.get(n)
        return is_checked(f.get("/V")) if f is not None else False

    sn = (norm_sn(tx("f_student_no")) or norm_sn(tx("f_student_no_b"))
          or norm_sn(os.path.splitext(os.path.basename(path))[0]))

    present = 0
    any_content = False
    for f in CRIT_FIELDS:
        if tx(f) != "":
            present += 1
            any_content = True
    for f in ("f_mtotal", "f_core_count"):
        if tx(f) != "":
            present += 1
            any_content = True
    for yf, nf in C_PAIRS:
        y, n = btn(yf), btn(nf)
        if y or n:
            any_content = True
        if y != n:
            present += 1
    comp, nyc = btn(OVERALL_PAIR[0]), btn(OVERALL_PAIR[1])
    if comp or nyc:
        any_content = True
    if comp != nyc:
        present += 1

    required_total = 6 + 2 + 4 + 1
    if present >= required_total:
        status = "full"
    elif not any_content:
        status = "ungraded"
    else:
        status = "partial"
    return True, sn, status, present


def scan(input_dir, exclude_names):
    out = []
    for root, dirs, files in os.walk(input_dir):
        dirs[:] = [d for d in dirs if d not in exclude_names]
        for f in sorted(files):
            if f.lower().endswith(".pdf") and not f.startswith("~"):
                out.append(os.path.join(root, f))
    return out


def main():
    ap = argparse.ArgumentParser(description="Combine GA7 rubrics, one per student.",
                                 usage="%(prog)s <input_dir> [output_dir] [--include-ungraded]")
    ap.add_argument("input_dir")
    ap.add_argument("output_dir", nargs="?", default=None)
    ap.add_argument("--include-ungraded", action="store_true",
                    help="Also include rubrics with no marks at all (default: skip).")
    args = ap.parse_args()

    input_dir = os.path.abspath(args.input_dir)
    if not os.path.isdir(input_dir):
        sys.exit(f"Input folder not found: {input_dir}")
    output_dir = os.path.abspath(args.output_dir or input_dir)
    combined_dir = os.path.join(output_dir, COMBINED_DIR)

    pdfs = scan(input_dir, {COMBINED_DIR})
    if not pdfs:
        sys.exit("No PDF files found to scan.")

    groups = {}
    non_rubric = 0
    for p in pdfs:
        ok, sn, status, filled = inspect(p)
        if not ok:
            non_rubric += 1
            continue
        key = sn if sn else f"__nosn__{os.path.basename(p)}"
        groups.setdefault(key, []).append((p, status, filled, os.path.getmtime(p)))

    chosen = {}
    skipped_ungraded = 0
    for key, items in groups.items():
        best = sorted(items, key=lambda t: (STATUS_RANK.get(t[1], 9), -t[2], -t[3]))[0]
        if best[1] == "ungraded" and not args.include_ungraded:
            skipped_ungraded += 1
            continue
        chosen[key] = best

    if not chosen:
        sys.exit("No graded rubrics found. Use --include-ungraded to include blanks.")

    if os.path.isdir(combined_dir):
        shutil.rmtree(combined_dir)
    os.makedirs(combined_dir, exist_ok=True)

    manifest = []
    combined = 0
    for key, (src, status, filled, _m) in sorted(chosen.items()):
        if key.startswith("__nosn__"):
            name = "feedback_" + os.path.splitext(os.path.basename(src))[0] + ".pdf"
        else:
            name = f"feedback_u{key}.pdf"
        dest = os.path.join(combined_dir, name)
        n = 1
        while os.path.exists(dest):
            dest = os.path.join(combined_dir, name.replace(".pdf", f"__{n}.pdf"))
            n += 1
        shutil.copy2(src, dest)
        combined += 1
        manifest.append({
            "student_no": ("u" + key) if not key.startswith("__nosn__") else "",
            "status": status,
            "source": os.path.relpath(src, input_dir),
            "combined": os.path.join(COMBINED_DIR, os.path.basename(dest)),
        })

    with open(os.path.join(output_dir, "combined_manifest.csv"), "w",
              newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=list(manifest[0].keys()))
        w.writeheader()
        w.writerows(manifest)

    print(f"Scanned folder : {input_dir}")
    print(f"  PDFs scanned     : {len(pdfs)}")
    print(f"  Non-rubric files : {non_rubric} (ignored)")
    print(f"  Students combined: {combined}  -> {COMBINED_DIR}/feedback_u<sn>.pdf")
    if skipped_ungraded:
        print(f"  Skipped ungraded : {skipped_ungraded} (use --include-ungraded to keep)")
    print(f"Manifest          : {os.path.join(output_dir, 'combined_manifest.csv')}")
    print(f"\nNext:  python flatten_pdfs.py {combined_dir}")


if __name__ == "__main__":
    main()
