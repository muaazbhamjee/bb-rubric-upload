#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sort_rubrics.py - classify MIA 320 GA7 rubric PDFs and sort them into folders.

Runs entirely offline (POPIA-safe); only dependency is pypdf.

What it does
------------
1. Recursively scans a folder (and all its subfolders) for rubric PDFs.
2. Classifies each rubric as:
       Fully graded     - every required field is filled
       Partially graded - some but not all required fields filled
       Ungraded         - no grading content at all
       Unreadable       - PDF could not be parsed
3. Detects duplicates (several files for the same student number): keeps the
   most complete copy and moves the rest to a Duplicates folder.
4. Reports counts and the total number of files scanned.
5. Moves each file into a categorised subfolder and writes manifest.csv.

Required fields that define "fully graded":
    f_m1..f_m6, f_mtotal, f_core_count,
    each of C1..C4 decided (exactly one of yes/no), and the overall outcome.

USAGE
-----
    pip install pypdf                      # one-off
    python sort_rubrics.py --input "C:/path/to/Marking"          # move files
    python sort_rubrics.py --input . --dry-run                   # report only
    #   --output DIR   where the category folders are created (default: <input>)

Destination layout (created in --output):
    Fully_Graded/        Partially_Graded/        Ungraded/
    Duplicates/          Unreadable/              manifest.csv

Moved files are renamed "<origin-subfolder>__<original-name>" so files from
different subfolders never collide. Use --dry-run first if unsure; nothing is
moved in that mode.
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

# ---- field contract (2026 GA7 fillable rubric) ----
CRIT_FIELDS = ["f_m1", "f_m2", "f_m3", "f_m4", "f_m5", "f_m6"]
C_PAIRS = [("f_c1y", "f_c1n"), ("f_c2y", "f_c2n"),
           ("f_c3y", "f_c3n"), ("f_c4y", "f_c4n")]
OVERALL_PAIR = ("f_overall_comp", "f_overall_notyet")

CATEGORY_DIRS = {
    "full": "Fully_Graded",
    "partial": "Partially_Graded",
    "ungraded": "Ungraded",
    "unreadable": "Unreadable",
    "duplicate": "Duplicates",
}


def norm_sn(raw):
    """Canonical student number: strip spaces, trailing .0, leading u/U; digits only."""
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"^[uU]", "", s)
    return re.sub(r"\D", "", s)


def is_checked(v):
    return v not in (None, "", "/Off", "Off", False)


def read_rubric(path):
    """Parse a rubric PDF and return a classification dict."""
    info = {"path": path, "ok": True, "sn": "", "status": "ungraded",
            "filled": 0, "error": ""}
    try:
        fields = PdfReader(path).get_fields() or {}
    except Exception as e:  # noqa: BLE001
        info["ok"] = False
        info["status"] = "unreadable"
        info["error"] = str(e)
        return info

    def tx(name):
        f = fields.get(name)
        return "" if f is None or f.get("/V") is None else str(f.get("/V")).strip()

    def btn(name):
        f = fields.get(name)
        return is_checked(f.get("/V")) if f is not None else False

    # student number
    info["sn"] = (norm_sn(tx("f_student_no")) or norm_sn(tx("f_student_no_b"))
                  or norm_sn(os.path.splitext(os.path.basename(path))[0]))

    # count filled "required" items and track how many are present
    present = 0
    required_total = 6 + 1 + 1 + 4 + 1  # marks + total + core + C1-4 + overall
    any_content = False

    for f in CRIT_FIELDS:
        if tx(f) != "":
            present += 1
            any_content = True
    if tx("f_mtotal") != "":
        present += 1
        any_content = True
    if tx("f_core_count") != "":
        present += 1
        any_content = True
    for yf, nf in C_PAIRS:
        y, n = btn(yf), btn(nf)
        if y or n:
            any_content = True
        if y != n:  # exactly one ticked = a decided indicator
            present += 1
    comp, nyc = btn(OVERALL_PAIR[0]), btn(OVERALL_PAIR[1])
    if comp or nyc:
        any_content = True
    if comp != nyc:
        present += 1

    info["filled"] = present
    if present >= required_total:
        info["status"] = "full"
    elif not any_content:
        info["status"] = "ungraded"
    else:
        info["status"] = "partial"
    return info


def scan(input_dir, exclude_names):
    """Yield every PDF path under input_dir, skipping excluded destination dirs."""
    out = []
    for root, dirs, files in os.walk(input_dir):
        dirs[:] = [d for d in dirs if d not in exclude_names]
        for f in sorted(files):
            if f.lower().endswith(".pdf") and not f.startswith("~"):
                out.append(os.path.join(root, f))
    return out


def main():
    ap = argparse.ArgumentParser(description="Classify and sort GA7 rubric PDFs.")
    ap.add_argument("--input", required=True, help="Folder to scan (recursively).")
    ap.add_argument("--output", default=None,
                    help="Where category folders are created (default: <input>).")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report only; do not move any files.")
    args = ap.parse_args()

    input_dir = os.path.abspath(args.input)
    if not os.path.isdir(input_dir):
        sys.exit(f"Input folder not found: {input_dir}")
    output_dir = os.path.abspath(args.output or input_dir)
    exclude = set(CATEGORY_DIRS.values())

    pdfs = scan(input_dir, exclude)
    if not pdfs:
        sys.exit("No PDF files found to scan.")

    records = [read_rubric(p) for p in pdfs]

    # group by student number (blank sn -> unique key per file, never a duplicate)
    groups = {}
    for i, r in enumerate(records):
        key = r["sn"] if r["sn"] else f"__nosn__{i}"
        groups.setdefault(key, []).append(r)

    # Choose which copy to preserve in a duplicate set:
    #   1. a fully graded copy, if any;
    #   2. otherwise a partially graded copy;
    #   3. otherwise (all ungraded/unreadable) preserve none - move them all.
    # Within a tier, prefer the most-filled, then the newest.
    def _best(cands):
        return sorted(cands, key=lambda r: (-r["filled"], -os.path.getmtime(r["path"])))[0]

    for key, rs in groups.items():
        if len(rs) == 1:
            rs[0]["is_dup"] = False
            continue
        fulls = [r for r in rs if r["status"] == "full"]
        partials = [r for r in rs if r["status"] == "partial"]
        if fulls:
            primary = _best(fulls)
        elif partials:
            primary = _best(partials)
        else:
            primary = None  # nothing worth preserving: all copies go to Duplicates
        for r in rs:
            if r is primary:
                r["is_dup"] = False
            else:
                r["is_dup"] = True
                r["primary"] = primary["path"] if primary else ""

    # tally
    counts = {"full": 0, "partial": 0, "ungraded": 0, "unreadable": 0}
    dup_versions = 0
    dup_students = 0
    for key, rs in groups.items():
        if len(rs) > 1 and not key.startswith("__nosn__"):
            dup_students += 1
    for r in records:
        if r.get("is_dup"):
            dup_versions += 1
        else:
            counts[r["status"]] += 1

    # plan destinations
    def dest_for(r):
        cat = "duplicate" if r.get("is_dup") else r["status"]
        folder = os.path.join(output_dir, CATEGORY_DIRS[cat])
        rel = os.path.relpath(os.path.dirname(r["path"]), input_dir)
        prefix = "" if rel in (".", "") else (rel.replace(os.sep, "_") + "__")
        name = prefix + os.path.basename(r["path"])
        return folder, name

    manifest = []
    for r in records:
        folder, name = dest_for(r)
        manifest.append({
            "student_no": ("u" + r["sn"]) if r["sn"] else "",
            "status": "duplicate" if r.get("is_dup") else r["status"],
            "source": os.path.relpath(r["path"], input_dir),
            "destination": os.path.join(os.path.basename(folder), name),
            "filled_of_13": r["filled"],
            "kept_primary": (os.path.relpath(r["primary"], input_dir)
                             if r.get("is_dup") and r.get("primary") else ""),
            "note": r["error"] or ("all copies ungraded; none preserved"
                                   if r.get("is_dup") and not r.get("primary") else ""),
        })

    # ---- report ----
    print(f"Scanned folder : {input_dir}")
    print(f"Total files    : {len(records)}")
    print(f"  Fully graded     : {counts['full']}")
    print(f"  Partially graded : {counts['partial']}")
    print(f"  Ungraded         : {counts['ungraded']}")
    if counts["unreadable"]:
        print(f"  Unreadable       : {counts['unreadable']}")
    print(f"  Duplicate copies : {dup_versions}  (across {dup_students} student(s))")
    check = counts['full'] + counts['partial'] + counts['ungraded'] \
        + counts['unreadable'] + dup_versions
    print(f"  (check: {check} == {len(records)} scanned)")

    if dup_students:
        print("\nDuplicates:")
        for key, rs in groups.items():
            if len(rs) > 1 and not key.startswith("__nosn__"):
                prim = next((r for r in rs if not r.get("is_dup")), None)
                extras = [r for r in rs if r.get("is_dup")]
                if prim:
                    print(f"  u{key}: kept {os.path.relpath(prim['path'], input_dir)} "
                          f"({prim['status']}); moved {len(extras)} copy(ies) to Duplicates")
                else:
                    print(f"  u{key}: all {len(extras)} copies ungraded - none preserved, "
                          "all moved to Duplicates")

    if args.dry_run:
        print("\nDRY RUN - no files moved. Re-run without --dry-run to apply.")
        _write_manifest(output_dir, manifest, dry=True)
        return

    # ---- move ----
    moved = 0
    for r in records:
        folder, name = dest_for(r)
        os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, name)
        base, ext = os.path.splitext(target)
        n = 1
        while os.path.exists(target):   # collision guard
            target = f"{base}__{n}{ext}"
            n += 1
        try:
            shutil.move(r["path"], target)
            moved += 1
        except Exception as e:  # noqa: BLE001
            print(f"  WARNING: could not move {r['path']}: {e}")

    _write_manifest(output_dir, manifest, dry=False)
    print(f"\nMoved {moved} file(s) into category folders under: {output_dir}")


def _write_manifest(output_dir, manifest, dry):
    name = "manifest_dryrun.csv" if dry else "manifest.csv"
    path = os.path.join(output_dir, name)
    try:
        os.makedirs(output_dir, exist_ok=True)
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=list(manifest[0].keys()))
            w.writeheader()
            w.writerows(manifest)
        print(f"Manifest       : {path}")
    except Exception as e:  # noqa: BLE001
        print(f"  WARNING: could not write manifest: {e}")


if __name__ == "__main__":
    main()
