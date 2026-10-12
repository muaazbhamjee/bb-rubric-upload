#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MIA 320 GA7 - Individual Project rubric reporting tool (offline).

Runs entirely on your own machine. No network access, nothing is uploaded.
Reads the completed fillable-PDF rubrics and the allocation workbook, then
writes a single .xlsx report with completion, per-TA grading spread,
empty-field flags, and duplicate / missing / misfiled detection.

------------------------------------------------------------------------------
REQUIREMENTS (one-off, offline once installed):
    pip install pypdf openpyxl

USAGE:
    python ga7_rubric_report.py --input "C:/path/to/Marking"
    # optional:
    #   --workbook "C:/path/to/allocations.xlsx"   (default: the single .xlsx in --input)
    #   --output   "C:/path/to/GA7_report.xlsx"     (default: <input>/GA7_report.xlsx)

EXPECTED INPUT LAYOUT:
    Marking/
        Abhinav/   <-- subfolder named per TA (matches a workbook tab)
            *.pdf  <-- that TA's completed rubrics
        Aaron/
            *.pdf
        ...
        allocations.xlsx   <-- one tab per TA; columns include Student Id, First Name, Last Name

The rubric PDFs must be the 2026 GA7 fillable form (fields f_student_no,
f_m1..f_m6, f_mtotal, f_c1y/f_c1n .. f_c4y/f_c4n, f_core_count,
f_overall_comp/f_overall_notyet, etc.).
------------------------------------------------------------------------------
"""

import argparse
import os
import re
import shutil
import csv
import statistics
import sys
from collections import defaultdict, OrderedDict

try:
    from pypdf import PdfReader
except ImportError:
    sys.exit("Missing dependency. Run:  pip install pypdf openpyxl")

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource, StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.legend import LegendEntry

# ----------------------------------------------------------------------------
# Field contract (from the 2026 GA7 fillable rubric)
# ----------------------------------------------------------------------------
CRIT_FIELDS = ["f_m1", "f_m2", "f_m3", "f_m4", "f_m5", "f_m6"]
CRIT_LABELS = ["Tech/SDG (5)", "Society (12)", "Economy (12)",
               "Environment (12)", "Risk (6)", "Modifications (3)"]
CRIT_MAX = [5, 12, 12, 12, 6, 3]
TOTAL_MAX = 50
PASS_PCT = 50.0

# Required = must be filled for a rubric to count as complete.
REQUIRED_TEXT = CRIT_FIELDS + ["f_mtotal", "f_core_count"]
# Optional text fields (reported but never counted as "missing")
OPTIONAL_TEXT = ["f_grader", "f_date", "f_comment", "f_student_name"]

C_PAIRS = [("f_c1y", "f_c1n"), ("f_c2y", "f_c2n"),
           ("f_c3y", "f_c3n"), ("f_c4y", "f_c4n")]
OVERALL_PAIR = ("f_overall_comp", "f_overall_notyet")

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def norm_sn(raw):
    """Canonical student-number key: strip whitespace, trailing .0, leading u/U."""
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"^[uU]", "", s)
    s = re.sub(r"\D", "", s)  # digits only
    return s


def display_sn(canonical):
    return ("u" + canonical) if canonical else ""


def is_checked(v):
    return v not in (None, "", "/Off", "Off", False)


def to_float(v):
    if v in (None, ""):
        return None
    try:
        return float(str(v).strip())
    except ValueError:
        return None


def btn_state(field):
    """Return the current value of a button/checkbox field."""
    v = field.get("/V")
    if v is None:
        return None
    return str(v)


# ----------------------------------------------------------------------------
# Read one rubric PDF
# ----------------------------------------------------------------------------
def read_rubric(path):
    r = {
        "path": path, "file": os.path.basename(path),
        "ok": True, "error": "",
        "f_student_no": "", "f_student_no_b": "", "f_student_name": "",
        "f_grader": "", "f_date": "", "f_comment": "",
        "marks": [None] * 6, "mtotal": None, "core_count": None,
        "c_met": [None] * 4, "overall": "", "missing_required": [],
    }
    try:
        fields = PdfReader(path).get_fields() or {}
    except Exception as e:  # noqa: BLE001
        r["ok"] = False
        r["error"] = f"unreadable PDF: {e}"
        return r

    def tx(name):
        f = fields.get(name)
        return "" if f is None or f.get("/V") is None else str(f.get("/V")).strip()

    r["f_student_no"] = tx("f_student_no")
    r["f_student_no_b"] = tx("f_student_no_b")
    r["f_student_name"] = tx("f_student_name")
    r["f_grader"] = tx("f_grader")
    r["f_date"] = tx("f_date")
    r["f_comment"] = tx("f_comment")

    r["marks"] = [to_float(tx(f)) for f in CRIT_FIELDS]
    r["mtotal"] = to_float(tx("f_mtotal"))
    r["core_count"] = to_float(tx("f_core_count"))

    # C1..C4 competency indicators
    for i, (yf, nf) in enumerate(C_PAIRS):
        yy = is_checked(btn_state(fields.get(yf, {})))
        nn = is_checked(btn_state(fields.get(nf, {})))
        if yy and not nn:
            r["c_met"][i] = True
        elif nn and not yy:
            r["c_met"][i] = False
        else:
            r["c_met"][i] = None  # neither or both -> inconsistent

    comp = is_checked(btn_state(fields.get(OVERALL_PAIR[0], {})))
    nyc = is_checked(btn_state(fields.get(OVERALL_PAIR[1], {})))
    if comp and not nyc:
        r["overall"] = "Competent"
    elif nyc and not comp:
        r["overall"] = "Not Yet Competent"
    else:
        r["overall"] = ""  # blank or both

    # Required-field completeness
    miss = []
    for f, lab in zip(CRIT_FIELDS, CRIT_LABELS):
        if to_float(tx(f)) is None:
            miss.append(lab)
    if r["mtotal"] is None:
        miss.append("Total mark")
    if r["core_count"] is None:
        miss.append("Core count")
    for i, cm in enumerate(r["c_met"]):
        if cm is None:
            miss.append(f"C{i+1}")
    if r["overall"] == "":
        miss.append("Overall outcome")
    r["missing_required"] = miss

    return r


# ----------------------------------------------------------------------------
# Load allocation workbook
# ----------------------------------------------------------------------------
def load_allocations(path):
    """Return OrderedDict {tab_name: OrderedDict {canonical_sn: {display, first, last}}}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    allocs = OrderedDict()
    dup_rows = []  # (tab, display_sn, count) for student numbers repeated within a tab
    for ws in wb.worksheets:
        header = [str(c.value).strip().lower() if c.value is not None else ""
                  for c in ws[1]]

        def col(*names, default=None):
            for n in names:
                if n in header:
                    return header.index(n)
            return default

        c_id = col("student id", "user name", default=1)
        c_first = col("first name", default=2)
        c_last = col("last name", default=3)

        students = OrderedDict()
        seen_counts = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None:
                continue
            raw = row[c_id] if c_id < len(row) else None
            key = norm_sn(raw)
            if not key:
                continue
            seen_counts[key] = seen_counts.get(key, 0) + 1
            first = row[c_first] if (c_first is not None and c_first < len(row)) else ""
            last = row[c_last] if (c_last is not None and c_last < len(row)) else ""
            students[key] = {
                "display": display_sn(key),
                "first": "" if first is None else str(first).strip(),
                "last": "" if last is None else str(last).strip(),
            }
        for key, cnt in seen_counts.items():
            if cnt > 1:
                dup_rows.append((ws.title.strip(), display_sn(key), cnt))
        allocs[ws.title.strip()] = students
    return allocs, dup_rows


# ----------------------------------------------------------------------------
# Walk input tree
# ----------------------------------------------------------------------------
def walk_rubrics(input_dir, workbook_path, exclude_dirs=None):
    """Return OrderedDict {ta_folder_name: [pdf paths]}."""
    exclude = {os.path.abspath(d) for d in (exclude_dirs or [])}
    out = OrderedDict()
    for name in sorted(os.listdir(input_dir)):
        sub = os.path.join(input_dir, name)
        if not os.path.isdir(sub):
            continue
        if os.path.abspath(sub) in exclude:
            continue  # never scan the moderation-sample output folder
        pdfs = []
        for f in sorted(os.listdir(sub)):
            if f.lower().endswith(".pdf") and not f.startswith("~"):
                pdfs.append(os.path.join(sub, f))
        out[name] = pdfs
    return out


def match_tab(folder_name, alloc_tabs):
    """Case-insensitive match of a subfolder to a workbook tab."""
    fn = folder_name.strip().lower()
    for t in alloc_tabs:
        if t.strip().lower() == fn:
            return t
    return None


# ----------------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------------
HDR = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="005BAA")      # UP Blue
BAD_FILL = PatternFill("solid", fgColor="F8D0D5")      # light UP Red
WARN_FILL = PatternFill("solid", fgColor="FCE8C8")     # amber
OK_FILL = PatternFill("solid", fgColor="D9EDDA")       # light green
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR
        cell.fill = HDR_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _set_fill(series, hexcolor):
    series.graphicalProperties = GraphicalProperties(solidFill=hexcolor)


def _fix_axes(chart):
    """Force Excel to show axis lines and tick labels.

    openpyxl leaves the axis 'delete' flag unset, which Excel treats as
    'hide the axis' (tick values disappear); LibreOffice ignores it, which
    is why the labels can look fine there but vanish in Excel.
    """
    for ax in (chart.x_axis, chart.y_axis):
        ax.delete = False
        ax.tickLblPos = "nextTo"
        ax.majorTickMark = "out"


def add_overall_charts(ws, per_ta, all_totals):
    """Add spread/variation charts to the Overall Stats sheet.

    Backing data is written to helper columns on the right (hidden) so the
    charts reference clean ranges independent of the tables above.
    """
    tas = list(per_ta.keys())
    if not tas:
        return

    anchor = ws.max_row + 2          # charts start just below the tables on the host sheet
    D = 1                            # first data column on the hidden data sheet

    # Chart backing data goes on its own hidden sheet. Charts do NOT render data
    # from hidden columns, but they DO render data from a hidden sheet, so this
    # keeps Overall Stats clean while the charts still display.
    dws = ws.parent.create_sheet("Chart data")
    dws.sheet_state = "hidden"
    dws.cell(row=1, column=1, value="Backing data for the Overall Stats charts - do not edit.")
    sheet = dws.title

    def numsrc(col, r0, r1):
        ref = "'%s'!$%s$%d:$%s$%d" % (sheet, get_column_letter(col), r0,
                                      get_column_letter(col), r1)
        return NumDataSource(NumRef(f=ref))

    def textcats(chart, r0, r1):
        """Set the category axis to a string reference on the TA/label column."""
        ref = "'%s'!$%s$%d:$%s$%d" % (sheet, get_column_letter(D), r0,
                                      get_column_letter(D), r1)
        src = AxDataSource(strRef=StrRef(f=ref))
        for s in chart.series:
            s.cat = src

    # ---- Block 1: TA | Min | Mean | Max | Stdev | Competent%  (for c2, c3) ----
    b1 = 2
    for j, h in enumerate(["TA", "Min", "Mean", "Max", "Stdev", "Competent%"]):
        dws.cell(row=b1, column=D + j, value=h)
    for i, ta in enumerate(tas):
        d = per_ta[ta]
        tv = d["totals"]
        row = b1 + 1 + i
        dws.cell(row=row, column=D, value=ta)
        if tv:
            dws.cell(row=row, column=D + 1, value=round(min(tv), 2))
            dws.cell(row=row, column=D + 2, value=round(statistics.mean(tv), 2))
            dws.cell(row=row, column=D + 3, value=round(max(tv), 2))
            dws.cell(row=row, column=D + 4,
                    value=round(statistics.pstdev(tv), 2) if len(tv) > 1 else 0)
        dws.cell(row=row, column=D + 5,
                value=round(100.0 * d["competent"] / d["judged"], 1) if d["judged"] else 0)
    b1_last = b1 + len(tas)

    # ---- Box block: TA | base(Q1) | Q1->Med | Med->Q3 | lowerWhisker | upperWhisker ----
    def five_num(v):
        v = sorted(v)
        lo, hi = v[0], v[-1]
        med = statistics.median(v)
        if len(v) >= 2:
            q = statistics.quantiles(v, n=4, method="inclusive")
            q1, q3 = q[0], q[2]
        else:
            q1 = q3 = v[0]
        return lo, q1, med, q3, hi

    bb = b1_last + 3
    for j, h in enumerate(["TA", "", "Q1-Median", "Median-Q3", "lowerW", "upperW"]):
        dws.cell(row=bb, column=D + j, value=h)
    for i, ta in enumerate(tas):
        tv = per_ta[ta]["totals"]
        row = bb + 1 + i
        dws.cell(row=row, column=D, value=ta)
        if tv:
            lo, q1, med, q3, hi = five_num(tv)
            dws.cell(row=row, column=D + 1, value=round(q1, 2))          # base (hidden)
            dws.cell(row=row, column=D + 2, value=round(med - q1, 2))    # Q1 -> median
            dws.cell(row=row, column=D + 3, value=round(q3 - med, 2))    # median -> Q3
            dws.cell(row=row, column=D + 4, value=round(q1 - lo, 2))     # lower whisker
            dws.cell(row=row, column=D + 5, value=round(hi - q3, 2))     # upper whisker
    bb_last = bb + len(tas)

    # Chart 1: BOX & WHISKER of total mark per TA
    box = BarChart()
    box.type = "col"
    box.grouping = "stacked"
    box.overlap = 100
    box.gapWidth = 60
    box.title = "Mark spread per TA (box & whisker, /50)"
    box.y_axis.title = "Mark /50"
    box.height, box.width = 8.5, 13
    box.add_data(Reference(dws, min_col=D + 1, max_col=D + 3, min_row=bb, max_row=bb_last),
                 titles_from_data=True)
    textcats(box, bb + 1, bb_last)
    # base segment invisible; box halves in UP palette
    gp0 = GraphicalProperties()
    gp0.noFill = True
    box.series[0].graphicalProperties = gp0
    _set_fill(box.series[1], "9EC4E6")   # Q1->median (light blue)
    _set_fill(box.series[2], "005BAA")   # median->Q3 (UP Blue)
    # whiskers via custom error bars
    box.series[0].errBars = ErrorBars(errDir="y", errBarType="minus",
                                      errValType="cust", noEndCap=False,
                                      minus=numsrc(D + 4, bb + 1, bb_last))
    box.series[2].errBars = ErrorBars(errDir="y", errBarType="plus",
                                      errValType="cust", noEndCap=False,
                                      plus=numsrc(D + 5, bb + 1, bb_last))
    # drop the invisible base series from the legend
    box.legend.legendEntry = [LegendEntry(idx=0, delete=True)]
    _fix_axes(box)
    ws.add_chart(box, f"A{anchor}")

    # Chart 2: Stdev per TA
    c2 = BarChart()
    c2.type = "col"
    c2.title = "Grading variation per TA (std dev)"
    c2.y_axis.title = "Std dev of total mark"
    c2.height, c2.width = 8.5, 13
    c2.legend = None
    c2.add_data(Reference(dws, min_col=D + 4, min_row=b1, max_row=b1_last),
                titles_from_data=True)
    textcats(c2, b1 + 1, b1_last)
    _set_fill(c2.series[0], "D71C33")
    _fix_axes(c2)
    ws.add_chart(c2, f"J{anchor}")

    # Chart 3: Competent rate per TA
    c3 = BarChart()
    c3.type = "col"
    c3.title = "Competent rate per TA (%)"
    c3.y_axis.title = "% Competent (of judged)"
    c3.y_axis.scaling.min = 0
    c3.y_axis.scaling.max = 100
    c3.height, c3.width = 8.5, 13
    c3.legend = None
    c3.add_data(Reference(dws, min_col=D + 5, min_row=b1, max_row=b1_last),
                titles_from_data=True)
    textcats(c3, b1 + 1, b1_last)
    _set_fill(c3.series[0], "007B7C")
    _fix_axes(c3)
    ws.add_chart(c3, f"A{anchor + 18}")

    # ---- Block 2: criterion mean as % of max, by TA ----
    b2 = bb_last + 3
    dws.cell(row=b2, column=D, value="Criterion")
    for j, ta in enumerate(tas):
        dws.cell(row=b2, column=D + 1 + j, value=ta)
    for i, lab in enumerate(CRIT_LABELS):
        row = b2 + 1 + i
        dws.cell(row=row, column=D, value=lab)
        for j, ta in enumerate(tas):
            cvals = per_ta[ta]["crit"][i]
            pct = round(100.0 * statistics.mean(cvals) / CRIT_MAX[i], 1) if cvals else None
            dws.cell(row=row, column=D + 1 + j, value=pct)
    b2_last = b2 + len(CRIT_LABELS)

    c4 = BarChart()
    c4.type = "col"
    c4.title = "Mean mark by criterion (% of max), by TA"
    c4.y_axis.title = "% of criterion max"
    c4.y_axis.scaling.min = 0
    c4.y_axis.scaling.max = 100
    c4.height, c4.width = 8.5, 13
    c4.add_data(Reference(dws, min_col=D + 1, max_col=D + len(tas),
                          min_row=b2, max_row=b2_last), titles_from_data=True)
    textcats(c4, b2 + 1, b2_last)
    _fix_axes(c4)
    ws.add_chart(c4, f"J{anchor + 18}")

    # ---- Block 3: cohort total-mark distribution ----
    bands = [("<50 (fail)", 0, 50), ("50-59", 50, 60), ("60-69", 60, 70),
             ("70-79", 70, 80), ("80-100", 80, 100.01)]
    pcts = [100.0 * t / TOTAL_MAX for t in all_totals]
    b3 = b2_last + 3
    dws.cell(row=b3, column=D, value="Band")
    dws.cell(row=b3, column=D + 1, value="Count")
    for i, (lab, lo, hi) in enumerate(bands):
        row = b3 + 1 + i
        cnt = sum(1 for p in pcts if lo <= p < hi)
        dws.cell(row=row, column=D, value=lab)
        dws.cell(row=row, column=D + 1, value=cnt)
    b3_last = b3 + len(bands)

    c5 = BarChart()
    c5.type = "col"
    c5.title = "Cohort total-mark distribution"
    c5.y_axis.title = "Students"
    c5.x_axis.title = "Mark band (%)"
    c5.height, c5.width = 8.5, 13
    c5.legend = None
    c5.add_data(Reference(dws, min_col=D + 1, min_row=b3, max_row=b3_last),
                titles_from_data=True)
    textcats(c5, b3 + 1, b3_last)
    _set_fill(c5.series[0], "005BAA")
    c5.dataLabels = DataLabelList()
    c5.dataLabels.showVal = True
    _fix_axes(c5)
    ws.add_chart(c5, f"A{anchor + 36}")

    # hide helper columns
    for c in range(D, D + max(6, 1 + len(tas)) + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = True


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="MIA 320 GA7 rubric report (offline).")
    ap.add_argument("--input", required=True, help="Folder containing per-TA subfolders.")
    ap.add_argument("--workbook", default=None, help="Allocation .xlsx (default: single .xlsx in --input).")
    ap.add_argument("--output", default=None, help="Output .xlsx path.")
    ap.add_argument("--moderation-dir", default=None,
                    help="Folder to copy the moderation-sample rubrics into "
                         "(default: <input>/Moderation_Sample).")
    ap.add_argument("--no-copy", action="store_true",
                    help="Do not copy moderation-sample rubric files (build the tab only).")
    ap.add_argument("--calibration-dir", default=None,
                    help="Folder of final agreed calibration rubrics "
                         "(default: <input>/Calibration). These students are counted "
                         "as graded-by-consensus, not missing.")
    args = ap.parse_args()

    input_dir = os.path.abspath(args.input)
    if not os.path.isdir(input_dir):
        sys.exit(f"Input folder not found: {input_dir}")

    # Output goes to the root input folder by default.
    out_path = os.path.abspath(args.output or os.path.join(input_dir, "GA7_report.xlsx"))
    out_name = os.path.basename(out_path).lower()
    moderation_dir = os.path.abspath(args.moderation_dir
                                     or os.path.join(input_dir, "Moderation_Sample"))
    if args.calibration_dir:
        calibration_dir = os.path.abspath(args.calibration_dir)
    else:
        # default <input>/Calibration, but match any top-level folder called
        # "calibration" regardless of case (e.g. Calibration, calibration, CALIBRATION)
        calibration_dir = os.path.join(input_dir, "Calibration")
        for d in sorted(os.listdir(input_dir)):
            full = os.path.join(input_dir, d)
            if os.path.isdir(full) and d.lower() == "calibration":
                calibration_dir = full
                break
        calibration_dir = os.path.abspath(calibration_dir)

    # locate workbook (in the root input folder), ignoring the report we write
    wb_path = args.workbook
    if wb_path is None:
        candidates = []
        for f in os.listdir(input_dir):
            fl = f.lower()
            if not fl.endswith(".xlsx") or f.startswith("~"):
                continue
            if fl == out_name or fl.startswith("ga7_report"):
                continue  # skip our own output / prior reports
            candidates.append(f)
        # if still ambiguous, keep only files that look like an allocation workbook
        if len(candidates) > 1:
            looks_alloc = []
            for f in candidates:
                try:
                    twb = openpyxl.load_workbook(os.path.join(input_dir, f),
                                                 read_only=True, data_only=True)
                    hdr = [str(c.value).strip().lower() if c.value is not None else ""
                           for c in next(twb.worksheets[0].iter_rows(max_row=1))]
                    twb.close()
                    if "student id" in hdr or "user name" in hdr:
                        looks_alloc.append(f)
                except Exception:  # noqa: BLE001
                    pass
            if looks_alloc:
                candidates = looks_alloc
        if len(candidates) != 1:
            sys.exit("Could not auto-find the allocation workbook in the root folder. "
                     "Pass it with --workbook. Candidates: " + str(candidates))
        wb_path = os.path.join(input_dir, candidates[0])
    wb_path = os.path.abspath(wb_path)

    print(f"Input     : {input_dir}")
    print(f"Workbook  : {wb_path}")
    print(f"Output    : {out_path}\n")

    allocs, alloc_dup_rows = load_allocations(wb_path)
    alloc_tabs = list(allocs.keys())

    # student -> allocated TA(s)  (detect a student listed under two tabs)
    sn_to_tabs = defaultdict(list)
    for tab, studs in allocs.items():
        for sn in studs:
            sn_to_tabs[sn].append(tab)

    folders = walk_rubrics(input_dir, wb_path,
                           exclude_dirs=[moderation_dir, calibration_dir])

    # Parse every rubric; index by canonical SN
    rubrics = []                      # list of parsed rubric dicts (+ folder, tab, sn)
    sn_to_rubrics = defaultdict(list)
    anomalies = []                    # (type, TA, student, detail)

    # allocation-sheet data quality: same student listed more than once in a tab
    for tab, disp, cnt in alloc_dup_rows:
        anomalies.append(("Duplicate row in allocation tab", tab, disp,
                          f"listed {cnt} times in this tab"))

    for folder, pdfs in folders.items():
        tab = match_tab(folder, alloc_tabs)
        if tab is None:
            anomalies.append(("Unmatched folder", folder, "",
                              "Subfolder name does not match any workbook tab."))
        for path in pdfs:
            r = read_rubric(path)
            r["folder"] = folder
            r["tab"] = tab
            if not r["ok"]:
                anomalies.append(("Unreadable PDF", folder, "", f"{r['file']}: {r['error']}"))
                continue

            sn_a = norm_sn(r["f_student_no"])
            sn_b = norm_sn(r["f_student_no_b"])
            sn_file = norm_sn(os.path.splitext(r["file"])[0])
            sn = sn_a or sn_b or sn_file
            r["sn"] = sn
            r["display_sn"] = display_sn(sn)

            # internal consistency checks
            if sn_a and sn_b and sn_a != sn_b:
                anomalies.append(("Page1/Page2 SN mismatch", folder, display_sn(sn_a),
                                  f"{r['file']}: page 2 shows {display_sn(sn_b)}"))
            if sn_file and sn and sn_file != sn:
                anomalies.append(("Filename vs field SN mismatch", folder, r["display_sn"],
                                  f"{r['file']}: filename says {display_sn(sn_file)}"))
            # core_count vs ticked C-boxes
            met = sum(1 for c in r["c_met"] if c is True)
            if r["core_count"] is not None and int(r["core_count"]) != met:
                anomalies.append(("Core-count mismatch", folder, r["display_sn"],
                                  f"stated {int(r['core_count'])}, ticked {met}"))
            # total vs sum of criteria
            if r["mtotal"] is not None and all(m is not None for m in r["marks"]):
                s = round(sum(r["marks"]), 2)
                if abs(s - r["mtotal"]) > 0.01:
                    anomalies.append(("Total mismatch", folder, r["display_sn"],
                                      f"sum={s}, stated total={r['mtotal']}"))
            # over-max criterion marks
            for m, mx, lab in zip(r["marks"], CRIT_MAX, CRIT_LABELS):
                if m is not None and (m < 0 or m > mx):
                    anomalies.append(("Out-of-range mark", folder, r["display_sn"],
                                      f"{lab}={m} (max {mx})"))

            rubrics.append(r)
            if sn:
                sn_to_rubrics[sn].append(r)

    # Duplicates
    for sn, rs in sn_to_rubrics.items():
        if len(rs) > 1:
            folders_used = {x["folder"] for x in rs}
            kind = "Cross-TA duplicate" if len(folders_used) > 1 else "Same-folder duplicate"
            detail = " | ".join(f"{x['folder']}/{x['file']}" for x in rs)
            ta_label = ", ".join(sorted(folders_used))
            anomalies.append((kind, ta_label, display_sn(sn), detail))

    # Students listed under two tabs
    for sn, tabs in sn_to_tabs.items():
        if len(tabs) > 1:
            anomalies.append(("Allocated to >1 TA", ", ".join(tabs), display_sn(sn),
                              "Student appears in multiple allocation tabs."))

    # Orphans: rubric SN not allocated anywhere
    for sn, rs in sn_to_rubrics.items():
        if sn and sn not in sn_to_tabs:
            for x in rs:
                anomalies.append(("Orphan (unallocated student)", x["folder"], display_sn(sn),
                                  x["file"]))

    # ------------------------------------------------------------------
    # Calibration: final agreed rubrics (flat PDFs). One consensus rubric per
    # calibration student. These students are graded-by-consensus, not missing.
    # ------------------------------------------------------------------
    cal_records = {}   # canonical_sn -> parsed agreed rubric
    cal_dir_exists = os.path.isdir(calibration_dir)
    cal_pdfs = []
    if cal_dir_exists:
        for root, _dirs, files in os.walk(calibration_dir):
            for f in sorted(files):
                if f.lower().endswith(".pdf") and not f.startswith("~"):
                    cal_pdfs.append(os.path.join(root, f))
        for path in sorted(cal_pdfs):
            r = read_rubric(path)
            if not r["ok"]:
                anomalies.append(("Unreadable PDF", "Calibration", "",
                                  f"{r['file']}: {r['error']}"))
                continue
            sn = (norm_sn(r["f_student_no"]) or norm_sn(r["f_student_no_b"])
                  or norm_sn(os.path.splitext(r["file"])[0]))
            r["sn"] = sn
            r["display_sn"] = display_sn(sn)
            r["relpath"] = os.path.relpath(path, input_dir)
            if sn in cal_records:
                anomalies.append(("Duplicate calibration rubric", "Calibration",
                                  display_sn(sn),
                                  f"{r['file']} and {cal_records[sn]['file']}"))
                continue
            cal_records[sn] = r
            if sn and sn not in sn_to_tabs:
                anomalies.append(("Calibration student not allocated", "Calibration",
                                  display_sn(sn), r["file"]))

    # ------------------------------------------------------------------
    # Build report
    # ------------------------------------------------------------------
    out = openpyxl.Workbook()
    out.remove(out.active)

    # Per-TA reconciliation results, reused by Completion/Overall tabs
    per_ta = OrderedDict()
    graded_records = []   # one entry per graded student, for moderation sampling
    outcomes = []         # (total, overall) per student for cohort stats; agreed mark wins
    combined_rows = []    # all students across every TA, for the "All Students" tab

    for tab in alloc_tabs:
        studs = allocs[tab]
        # rubrics physically filed under this TA folder (matched)
        filed = [r for r in rubrics if r["tab"] == tab]
        filed_by_sn = defaultdict(list)
        for r in filed:
            filed_by_sn[r["sn"]].append(r)

        rows = []
        n_graded = n_missing = n_misfiled = n_dup = n_calib = 0
        totals_for_stats = []
        crit_for_stats = [[] for _ in range(6)]
        competent_count = 0
        judged_count = 0
        pass_count = 0

        for sn, info in studs.items():
            here = filed_by_sn.get(sn, [])
            elsewhere = [r for r in sn_to_rubrics.get(sn, []) if r["tab"] != tab]

            if here:
                status = "Graded"
                if len(here) > 1:
                    status = "Graded (DUPLICATE)"
                    n_dup += 1
                r = here[0]
                n_graded += 1
                empties = "; ".join(r["missing_required"]) if r["missing_required"] else ""
                marks = r["marks"]
                total = r["mtotal"]
                cc = "" if r["core_count"] is None else int(r["core_count"])
                cmet = "".join("Y" if c is True else ("N" if c is False else "?")
                               for c in r["c_met"])
                overall = r["overall"]
                grader = r["f_grader"]
                # stats
                if total is not None:
                    totals_for_stats.append(total)
                for i, m in enumerate(marks):
                    if m is not None:
                        crit_for_stats[i].append(m)
                if overall in ("Competent", "Not Yet Competent"):
                    judged_count += 1
                    if overall == "Competent":
                        competent_count += 1
                if (total is not None and overall == "Competent"
                        and (100.0 * total / TOTAL_MAX) >= PASS_PCT):
                    pass_count += 1
                note = ""
                if len(here) > 1:
                    note = f"{len(here)} rubrics in this folder"
                if total is not None:
                    graded_records.append({
                        "display_sn": info["display"],
                        "name": (info["first"] + " " + info["last"]).strip(),
                        "ta": tab, "total": total, "overall": overall,
                        "core": cc, "empties": empties,
                        "relpath": os.path.relpath(r["path"], input_dir),
                        "srcpath": r["path"],
                    })
            elif elsewhere:
                status = "Misfiled (graded under " + ", ".join(sorted({r['folder'] for r in elsewhere})) + ")"
                n_misfiled += 1
                marks = [None] * 6
                total = cc = overall = grader = ""
                cmet = ""
                empties = ""
                note = "; ".join(f"{r['folder']}/{r['file']}" for r in elsewhere)
            elif sn in cal_records:
                cr = cal_records[sn]
                status = "In calibration"
                n_calib += 1
                empties = "; ".join(cr["missing_required"]) if cr["missing_required"] else ""
                marks = cr["marks"]
                total = cr["mtotal"]
                cc = "" if cr["core_count"] is None else int(cr["core_count"])
                cmet = "".join("Y" if c is True else ("N" if c is False else "?")
                               for c in cr["c_met"])
                overall = cr["overall"]
                grader = cr["f_grader"]
                note = "final agreed calibration rubric"
            else:
                status = "MISSING"
                n_missing += 1
                marks = [None] * 6
                total = cc = overall = grader = ""
                cmet = ""
                empties = ""
                note = ""

            # cohort outcome: agreed calibration mark wins over the TA's own mark
            if sn in cal_records and cal_records[sn]["mtotal"] is not None:
                cr = cal_records[sn]
                outcomes.append((cr["mtotal"], cr["overall"]))
            elif here and here[0]["mtotal"] is not None:
                outcomes.append((here[0]["mtotal"], here[0]["overall"]))

            name = (info["first"] + " " + info["last"]).strip()
            row_cells = ([info["display"], name, status, grader]
                         + [("" if m is None else m) for m in marks]
                         + [("" if total in ("", None) else total), cc, cmet,
                            overall, empties, note])
            rows.append(row_cells)
            # combined view: insert the TA after Name
            combined_rows.append([row_cells[0], row_cells[1], tab] + row_cells[2:])

        per_ta[tab] = {
            "allocated": len(studs), "graded": n_graded, "missing": n_missing,
            "misfiled": n_misfiled, "dup": n_dup, "calib": n_calib,
            "totals": totals_for_stats, "crit": crit_for_stats,
            "competent": competent_count, "judged": judged_count,
            "pass_ga7": pass_count,
        }

        # write per-TA sheet
        ws = out.create_sheet(tab[:31])
        head = (["Student no", "Name", "Status", "Grader"]
                + CRIT_LABELS + ["Total/50", "Core", "C1-4", "GA7 outcome",
                                 "Empty required fields", "Notes"])
        ws.append(head)
        for row in rows:
            ws.append(row)
        style_header(ws, len(head))
        for ridx in range(2, ws.max_row + 1):
            st = ws.cell(row=ridx, column=3).value or ""
            fill = None
            if st.startswith("MISSING"):
                fill = BAD_FILL
            elif "DUPLICATE" in st or st.startswith("Misfiled"):
                fill = WARN_FILL
            elif st.startswith("In calibration"):
                fill = PatternFill("solid", fgColor="D6ECEC")  # light teal
            elif ws.cell(row=ridx, column=len(head) - 1).value:  # empty-fields col
                fill = WARN_FILL
            if fill:
                for c in range(1, len(head) + 1):
                    ws.cell(row=ridx, column=c).fill = fill
        autosize(ws, [12, 22, 26, 8] + [11] * 6 + [9, 6, 8, 18, 34, 34])

    # ---------------- All Students tab (per-TA layout, whole class) ----------------
    ws = out.create_sheet("All Students")
    head = (["Student no", "Name", "TA", "Status", "Grader"]
            + CRIT_LABELS + ["Total/50", "Core", "C1-4", "GA7 outcome",
                             "Empty required fields", "Notes"])
    ws.append(head)
    for row in combined_rows:
        ws.append(row)
    style_header(ws, len(head))
    for ridx in range(2, ws.max_row + 1):
        st = ws.cell(row=ridx, column=4).value or ""   # Status is column 4 here
        fill = None
        if st.startswith("MISSING"):
            fill = BAD_FILL
        elif "DUPLICATE" in st or st.startswith("Misfiled"):
            fill = WARN_FILL
        elif st.startswith("In calibration"):
            fill = PatternFill("solid", fgColor="D6ECEC")
        elif ws.cell(row=ridx, column=len(head) - 1).value:  # empty-fields col
            fill = WARN_FILL
        if fill:
            for c in range(1, len(head) + 1):
                ws.cell(row=ridx, column=c).fill = fill
    autosize(ws, [12, 22, 14, 26, 8] + [11] * 6 + [9, 6, 8, 18, 30, 30])

    # ---------------- Completion tab ----------------
    ws = out.create_sheet("Completion")
    head = ["TA", "Allocated", "Graded", "In calibration", "Missing", "Misfiled",
            "Duplicates", "% complete", "Missing students"]
    ws.append(head)
    tot = defaultdict(int)
    cal_sns = set(cal_records.keys())
    for tab, d in per_ta.items():
        # accounted for = own graded + calibration-agreed + misfiled elsewhere
        accounted = d["graded"] + d["calib"] + d["misfiled"]
        pct = round(100.0 * accounted / d["allocated"], 1) if d["allocated"] else 0.0
        # names of genuinely missing students (not graded anywhere, not in calibration)
        miss_names = []
        filed_sns = {r["sn"] for r in rubrics if r["tab"] == tab}
        any_sns = set(sn_to_rubrics.keys())
        for sn, info in allocs[tab].items():
            if sn not in filed_sns and sn not in any_sns and sn not in cal_sns:
                miss_names.append(info["display"])
        ws.append([tab, d["allocated"], d["graded"], d["calib"], d["missing"],
                   d["misfiled"], d["dup"], pct, ", ".join(miss_names)])
        for k in ("allocated", "graded", "calib", "missing", "misfiled", "dup"):
            tot[k] += d[k]
    grand_acc = tot["graded"] + tot["calib"] + tot["misfiled"]
    grand_pct = round(100.0 * grand_acc / tot["allocated"], 1) if tot["allocated"] else 0.0
    ws.append(["TOTAL", tot["allocated"], tot["graded"], tot["calib"], tot["missing"],
               tot["misfiled"], tot["dup"], grand_pct, ""])
    style_header(ws, len(head))
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    for ridx in range(2, ws.max_row):
        if ws.cell(row=ridx, column=5).value:      # missing > 0
            ws.cell(row=ridx, column=5).fill = BAD_FILL
        if ws.cell(row=ridx, column=7).value:      # dup > 0
            ws.cell(row=ridx, column=7).fill = WARN_FILL
    autosize(ws, [14, 11, 9, 13, 9, 9, 11, 11, 55])

    # ---------------- Overall Stats tab ----------------
    ws = out.create_sheet("Overall Stats")

    def stat_block(title, values):
        ws.append([title])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        if values:
            ws.append(["n", len(values)])
            ws.append(["mean", round(statistics.mean(values), 2)])
            ws.append(["median", round(statistics.median(values), 2)])
            ws.append(["stdev", round(statistics.pstdev(values), 2) if len(values) > 1 else 0])
            ws.append(["min", round(min(values), 2)])
            ws.append(["max", round(max(values), 2)])
        else:
            ws.append(["(no data)"])
        ws.append([])

    # Cohort figures use resolved outcomes (own marking + calibration-agreed,
    # agreed mark winning). Per-TA spread below stays on own marking only.
    all_totals = [t for (t, o) in outcomes]
    all_competent = sum(1 for (t, o) in outcomes if o == "Competent")
    all_judged = sum(1 for (t, o) in outcomes if o in ("Competent", "Not Yet Competent"))

    ws.append(["COHORT SUMMARY"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
    ws.append(["Allocated", tot["allocated"]])
    ws.append(["Graded (TA marking)", tot["graded"]])
    ws.append(["In calibration (agreed)", tot["calib"]])
    ws.append(["Overall completion %", grand_pct])
    ws.append(["Declared Competent", all_competent])
    ws.append(["Declared Not Yet Competent", all_judged - all_competent])
    ws.append(["Competent rate % (of judged)",
               round(100.0 * all_competent / all_judged, 1) if all_judged else 0.0])
    # pass GA7 = >=50% in Part A AND Competent (resolved outcomes, incl. calibration)
    pass_ga7 = sum(1 for (t, o) in outcomes
                   if o == "Competent" and (100.0 * t / TOTAL_MAX) >= PASS_PCT)
    ws.append(["Passing GA7 (>=50% AND Competent)", pass_ga7])
    ws.append([])

    stat_block("TOTAL MARK /50 (all outcomes, incl. calibration)", all_totals)

    # per-TA grading spread
    ws.append(["PER-TA GRADING SPREAD"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
    head = ["TA", "n", "mean/50", "median", "stdev", "min", "max",
            "Competent %"] + [f"mean {l}" for l in CRIT_LABELS]
    ws.append(head)
    hdr_row = ws.max_row
    for tab, d in per_ta.items():
        tv = d["totals"]
        if tv:
            comp_rate = round(100.0 * d["competent"] / d["judged"], 1) if d["judged"] else 0.0
            row = [tab, len(tv), round(statistics.mean(tv), 2),
                   round(statistics.median(tv), 2),
                   round(statistics.pstdev(tv), 2) if len(tv) > 1 else 0,
                   round(min(tv), 2), round(max(tv), 2), comp_rate]
            for cvals in d["crit"]:
                row.append(round(statistics.mean(cvals), 2) if cvals else "")
            ws.append(row)
        else:
            ws.append([tab, 0] + [""] * (len(head) - 2))
    for c in range(1, len(head) + 1):
        cell = ws.cell(row=hdr_row, column=c)
        cell.font = HDR
        cell.fill = HDR_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    autosize(ws, [14, 6, 9, 9, 8, 7, 7, 12] + [13] * 6)

    add_overall_charts(ws, per_ta, all_totals)

    # ---------------- Calibration tab (shown whenever a Calibration folder exists) ----------------
    if cal_dir_exists:
        cws = out.create_sheet("Calibration")
        if not cal_records:
            cws.append(["Final agreed calibration rubrics"])
            cws.cell(row=1, column=1).font = Font(bold=True, size=13, color="005BAA")
            cws.append([f"No readable calibration rubrics were found in "
                        f"{os.path.relpath(calibration_dir, input_dir)} "
                        f"({len(cal_pdfs)} PDF(s) scanned). Check the folder contents."])
            cws.cell(row=2, column=1).font = Font(italic=True, size=9, color="555555")
            autosize(cws, [60])
    if cal_records:
        # name / allocated-TA lookup
        name_of = {}
        for tb, studs in allocs.items():
            for sn, info in studs.items():
                name_of[sn] = (info["first"] + " " + info["last"]).strip()
        cws.append(["Final agreed calibration rubrics"])
        cws.cell(row=1, column=1).font = Font(bold=True, size=13, color="005BAA")
        cal_totals = [r["mtotal"] for r in cal_records.values() if r["mtotal"] is not None]
        summ = f"{len(cal_records)} agreed rubric(s)."
        if cal_totals:
            summ += (f" Mean {round(statistics.mean(cal_totals),1)}/50, "
                     f"median {round(statistics.median(cal_totals),1)}/50. ")
        summ += ("These are consensus marks and are used as the official outcome for "
                 "these students; they are excluded from per-TA grading-spread stats "
                 "and from the moderation sample.")
        cws.append([summ])
        cws.cell(row=2, column=1).font = Font(italic=True, size=9, color="555555")
        cws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
        cws.cell(row=2, column=1).alignment = Alignment(wrap_text=True)
        # same layout as the per-TA tabs, with the agreed per-criterion marks
        chead = (["Student no", "Name", "Allocated TA"] + CRIT_LABELS
                 + ["Total/50", "Core", "C1-4", "GA7 outcome",
                    "Empty required fields", "Rubric file"])
        cws.append(chead)
        for sn in sorted(cal_records, key=lambda s: (cal_records[s]["mtotal"] is None,
                                                     -(cal_records[s]["mtotal"] or 0))):
            r = cal_records[sn]
            ta = ", ".join(sn_to_tabs.get(sn, [])) or "unallocated"
            tot_v = r["mtotal"]
            cc = "" if r["core_count"] is None else int(r["core_count"])
            cmet = "".join("Y" if c is True else ("N" if c is False else "?")
                           for c in r["c_met"])
            empties = "; ".join(r["missing_required"]) if r["missing_required"] else ""
            cws.append([r["display_sn"], name_of.get(sn, ""), ta]
                       + [("" if m is None else m) for m in r["marks"]]
                       + [("" if tot_v is None else tot_v), cc, cmet, r["overall"],
                          empties, r["relpath"]])
            if empties:
                for c in range(1, len(chead) + 1):
                    cws.cell(row=cws.max_row, column=c).fill = PatternFill("solid", fgColor="FCE8C8")
        style_header(cws, len(chead), row=3)
        autosize(cws, [12, 22, 16] + [11] * 6 + [9, 6, 8, 18, 30, 40])

    # ---------------- Moderation Sample tab (ECSA external moderation) ----------------
    ws = out.create_sheet("Moderation Sample")
    NCOL = 10
    pool = [g for g in graded_records if g["total"] is not None]
    pool_sorted = sorted(pool, key=lambda g: g["total"])
    med = statistics.median([g["total"] for g in pool]) if pool else 0.0
    mean = statistics.mean([g["total"] for g in pool]) if pool else 0.0
    pass_mark = TOTAL_MAX * PASS_PCT / 100.0  # 25/50

    top5 = list(reversed(pool_sorted[-5:]))
    # lowest 5 excludes zero totals (a 0 is a non-submission, not a gradeable low mark)
    low5 = [g for g in pool_sorted if g["total"] and g["total"] > 0][:5]
    middle5 = sorted(sorted(pool, key=lambda g: abs(g["total"] - med))[:5],
                     key=lambda g: -g["total"])
    boundary5 = sorted(sorted(pool, key=lambda g: abs(g["total"] - pass_mark))[:5],
                       key=lambda g: -g["total"])

    # intro
    ws.append(["External moderation sample - MIA 320 GA7 Individual Project"])
    ws.cell(row=1, column=1).font = Font(bold=True, size=13, color="005BAA")
    ws.append([f"Graded reports available: {len(pool)} | mean {round(mean,1)}/50 | "
               f"median {round(med,1)}/50 | pass mark {pass_mark:g}/50 (50%). "
               "Provide the moderator with each listed student's report and completed rubric."])
    ws.cell(row=2, column=1).font = Font(italic=True, size=9, color="555555")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOL)
    ws.cell(row=2, column=1).alignment = Alignment(wrap_text=True)

    def write_block(title, rationale, records):
        ws.append([])
        br = ws.max_row + 1
        ws.cell(row=br, column=1, value=title)
        ws.merge_cells(start_row=br, start_column=1, end_row=br, end_column=NCOL)
        b = ws.cell(row=br, column=1)
        b.font = Font(bold=True, color="FFFFFF")
        b.fill = HDR_FILL
        b.alignment = Alignment(vertical="center")
        rr = br + 1
        ws.cell(row=rr, column=1, value=rationale)
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=NCOL)
        ws.cell(row=rr, column=1).font = Font(italic=True, size=9, color="555555")
        ws.cell(row=rr, column=1).alignment = Alignment(wrap_text=True)
        hdr = ["#", "Student no", "Name", "TA", "Total/50", "%", "GA7 outcome",
               "Core (/4)", "Empty fields?", "Rubric file"]
        ws.append(hdr)
        hrow = ws.max_row
        for c in range(1, NCOL + 1):
            x = ws.cell(row=hrow, column=c)
            x.font = Font(bold=True)
            x.fill = PatternFill("solid", fgColor="E7EEF6")
        if not records:
            ws.append(["(no graded reports available)"])
            return
        for i, g in enumerate(records, 1):
            pct = round(100.0 * g["total"] / TOTAL_MAX, 1)
            outcome = g["overall"]
            # for the boundary block, tag pass/fail on the mark
            ws.append([i, g["display_sn"], g["name"], g["ta"], g["total"], pct,
                       outcome, g["core"],
                       ("Yes (" + g["empties"] + ")") if g["empties"] else "",
                       g["relpath"]])
            if pct < PASS_PCT:
                ws.cell(row=ws.max_row, column=6).fill = BAD_FILL

    write_block("1. Top 5 (highest marks)",
                "Verifies the ceiling of marking: are the strongest reports awarded appropriately.",
                top5)
    write_block("2. Lowest 5 (lowest non-zero marks)",
                "Verifies the floor of marking and the treatment of weak or failing work. "
                "Zero totals are excluded as non-submissions.",
                low5)
    write_block(f"3. Middle 5 (nearest median {round(med,1)}/50)",
                "Verifies typical, mid-range marking around the central tendency of the cohort.",
                middle5)
    write_block(f"4. Boundary 5 (nearest the 50% pass line, {pass_mark:g}/50)",
                "Verifies the pass/fail decision. Rows shaded in the % column are below 50% "
                "(just failing); the remainder are just passing.",
                boundary5)

    autosize(ws, [4, 12, 22, 12, 9, 7, 18, 9, 30, 46])
    ws.freeze_panes = "A3"

    # ---------------- Empty Fields tab ----------------
    ws = out.create_sheet("Empty Fields")
    head = ["TA", "Student no", "File", "Empty required fields"]
    ws.append(head)
    for r in rubrics:
        if r["missing_required"]:
            ws.append([r["folder"], r["display_sn"], r["file"],
                       "; ".join(r["missing_required"])])
    if ws.max_row == 1:
        ws.append(["(none)", "", "", ""])
    style_header(ws, len(head))
    autosize(ws, [14, 12, 34, 60])

    # ---------------- Anomalies tab ----------------
    ws = out.create_sheet("Anomalies")
    head = ["Type", "TA / folder", "Student no", "Detail"]
    ws.append(head)
    for a in sorted(anomalies, key=lambda x: x[0]):
        ws.append(list(a))
    if ws.max_row == 1:
        ws.append(["(none)", "", "", ""])
    style_header(ws, len(head))
    for ridx in range(2, ws.max_row + 1):
        ws.cell(row=ridx, column=1).fill = BAD_FILL
    autosize(ws, [30, 22, 12, 70])

    # order tabs: Completion, Overall, then TAs, then Empty, Anomalies
    order = ["Completion", "Overall Stats", "Moderation Sample", "All Students"] \
            + [t[:31] for t in alloc_tabs] + ["Calibration"] \
            + ["Empty Fields", "Anomalies", "Chart data"]
    out._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

    out.save(out_path)

    # ---- copy sampled rubrics for the moderator ----
    copied = 0
    if not args.no_copy and pool:
        sample_sets = [("1_Top5", top5), ("2_Lowest5", low5),
                       ("3_Middle5", middle5), ("4_Boundary5", boundary5)]
        if os.path.isdir(moderation_dir):
            shutil.rmtree(moderation_dir)  # clear stale copies from a previous run
        os.makedirs(moderation_dir, exist_ok=True)
        manifest_rows = []
        for folder_name, records in sample_sets:
            dest = os.path.join(moderation_dir, folder_name)
            os.makedirs(dest, exist_ok=True)
            for i, g in enumerate(records, 1):
                base = os.path.basename(g["srcpath"])
                copied_as = f"{i:02d}_{base}"
                try:
                    shutil.copy2(g["srcpath"], os.path.join(dest, copied_as))
                    copied += 1
                except Exception as e:  # noqa: BLE001
                    copied_as = f"COPY FAILED: {e}"
                manifest_rows.append([
                    folder_name, i, g["display_sn"], g["name"], g["ta"],
                    g["total"], round(100.0 * g["total"] / TOTAL_MAX, 1),
                    g["overall"], g["relpath"], os.path.join(folder_name, copied_as)])
        with open(os.path.join(moderation_dir, "manifest.csv"), "w",
                  newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["Category", "Rank", "Student no", "Name", "TA", "Total/50",
                        "%", "GA7 outcome", "Rubric source", "Copied as"])
            w.writerows(manifest_rows)
        with open(os.path.join(moderation_dir, "README.txt"), "w", encoding="utf-8") as fh:
            fh.write("MIA 320 GA7 - external moderation sample\n"
                     "Rubric copies grouped by sampling category. Provide each "
                     "student's report alongside the rubric. See manifest.csv and "
                     "the 'Moderation Sample' tab of the report for details.\n")

    print("Done.")
    if cal_dir_exists:
        print(f"  Calibration    : {len(cal_records)} rubric(s) from "
              f"{os.path.relpath(calibration_dir, input_dir)} "
              f"({len(cal_pdfs)} PDF(s) scanned)")
    else:
        print(f"  Calibration    : no folder found "
              f"(looked for {os.path.relpath(calibration_dir, input_dir)})")
    print(f"  TAs            : {len(alloc_tabs)}")
    print(f"  Allocated      : {tot['allocated']}")
    print(f"  Graded         : {tot['graded']}  ({grand_pct}%)")
    print(f"  In calibration : {tot['calib']}")
    print(f"  Missing        : {tot['missing']}")
    print(f"  Misfiled       : {tot['misfiled']}")
    print(f"  Duplicates     : {tot['dup']}")
    print(f"  Anomalies      : {len(anomalies)}")
    print(f"\nReport written to: {out_path}")
    if not args.no_copy and pool:
        print(f"Moderation copies: {copied} rubric(s) -> {moderation_dir}")


if __name__ == "__main__":
    main()
