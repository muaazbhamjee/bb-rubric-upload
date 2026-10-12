# UP Rubric Tooling - Grade Upload, Reporting & Student Feedback

Offline tooling for marking-rubric workflows at the University of Pretoria. It covers
two courses that share the same flatten and email steps:

- **MKM411** - bulk semester-test grade upload to Blackboard Ultra, plus emailing per-student rubric PDFs.
- **MIA 320** - GA7 individual-project rubric reporting, external moderation, triage, and emailing per-student rubric PDFs.

Everything runs locally for POPIA compliance; no student data leaves your machine except
the rubric PDFs you choose to email, which are sent from your own UP Google Workspace account.

---

## Prerequisites

### Python (3.10+)

```bash
pip install pymupdf pypdf openpyxl
```

| Package   | Used by |
|-----------|---------|
| `pymupdf` | `flatten_pdfs.py` (preferred flatten backend) |
| `pypdf`   | `prepare_student_rubrics.py`, `sort_rubrics.py`, `ga7_rubric_report.py`, flatten fallback |
| `openpyxl`| `ga7_rubric_report.py`, `package_for_upload.py` |

### Optional system package - pdftk

Only needed if you choose `--backend pdftk` for flattening (PyMuPDF is the default and
needs no system install). Linux/WSL: `sudo apt install pdftk-java`; macOS: `brew install pdftk-java`.

### Google Apps Script

No install. `email.js` runs in the browser at [script.google.com](https://script.google.com)
using your UP Google Workspace account (Drive + Gmail).

---

## Scripts

| Script | Purpose | Course | Runs on |
|---|---|---|---|
| `ga7_rubric_report.py` | GA7 report: completion, per-TA grading spread (box plots), calibration, moderation sample | MIA 320 | Python |
| `sort_rubrics.py` | Classify rubrics (fully/partially/ungraded) and sort into folders; handle duplicates | MIA 320 | Python |
| `prepare_student_rubrics.py` | Combine graded rubrics, one per student, as `feedback_u<sn>.pdf` | MIA 320 | Python |
| `rename_for_blackboard.py` | Rename test rubrics to `feedback_u<sn>.pdf` from the filename | MKM411 | Python |
| `flatten_pdfs.py` | Flatten fillable PDFs to static, non-editable content (parallel) | shared | Python |
| `package_for_upload.py` | Build the Blackboard Marks & Feedback upload ZIP with an Excel template | shared | Python |
| `email.js` | Email rubric PDFs to students via Gmail | shared | Google Apps Script |

**Naming convention (both courses):** `feedback_u<studentnumber>.pdf`.
Blackboard username at UP = `u` + student number; student email = `u<number>@tuks.co.za`.

---

## MIA 320 - GA7 individual project

### 1. (Optional) Triage the marking folder

Classify rubrics and separate graded from ungraded, handling duplicates.

```bash
python sort_rubrics.py --input "/path/to/Graded" --dry-run   # preview
python sort_rubrics.py --input "/path/to/Graded"             # then move
```

Sorts into `Fully_Graded/`, `Partially_Graded/`, `Ungraded/`, `Unreadable/`, `Duplicates/`
with a manifest. Duplicate rule: keep a fully graded copy, else a partially graded copy, else
move all copies to `Duplicates/`.

### 2. Reporting, moderation & calibration

Run against the marking folder (per-TA subfolders + the allocation workbook; optional
`Calibration/` folder of final agreed rubrics).

```bash
python ga7_rubric_report.py --input "/path/to/Graded"
```

Produces `GA7_report.xlsx`: Completion, Overall Stats (box-and-whisker spread per TA, variation,
competency, distribution), All Students, per-TA tabs, Calibration, Moderation Sample, Empty
Fields, Anomalies. Also copies the external-moderation sample into `Moderation_Sample/`.

### 3. Combine rubrics for students

Gather the graded rubrics (deduped, most complete copy per student) into one folder, named for
email/Blackboard.

```bash
python prepare_student_rubrics.py "/path/to/Graded"
# -> /path/to/Graded/combined/feedback_u<sn>.pdf  (+ combined_manifest.csv)
```

### 4. Flatten

```bash
python flatten_pdfs.py "/path/to/Graded/combined/"
# -> combined/flattened/   (non-editable; open one to verify)
```

### 5. Email

Upload `combined/flattened/` to a Drive folder, then follow **Emailing rubrics** below with the
MIA 320 config (the default in `email.js`).

---

## MKM411 - semester-test grade upload

### 1. Rename

```bash
python rename_for_blackboard.py /path/to/filled-rubrics/
# -> renamed/feedback_u<sn>.pdf
```

### 2. Flatten

```bash
python flatten_pdfs.py /path/to/filled-rubrics/renamed/
# -> renamed/flattened/
```

### 3. Upload grades to Blackboard

```bash
python package_for_upload.py /path/to/renamed/flattened/ "MKM411 Semester Test 1" "MKM411_2026" MKM411_ST1_upload.zip
```

Extract the ZIP, fill in `FIRSTNAME`, `LASTNAME`, `Grade` (and optional `Feedback`) in the Excel,
remove rows for students who did not write, do **not** rename the Excel, then re-zip preserving
structure:

```bash
cd extracted_folder/ && zip -r ../MKM411_ST1_upload.zip *
```

Upload via **Content Market → Marks and Feedback → Upload**, then **Release Grades**.
ZIP constraints: no spaces in the filename, under 250 MB (batch if larger).

### 4. Email

Upload the flattened PDFs to a Drive folder, then follow **Emailing rubrics** with the MKM411
preset from the bottom of `email.js`.

---

## Emailing rubrics (shared, Google Apps Script)

1. Upload the flattened `feedback_u*.pdf` files to a Google Drive folder, and copy the folder ID
   from the URL (`.../folders/<FOLDER_ID>`).
2. At [script.google.com](https://script.google.com) → **New Project**, paste `email.js`.
3. In the CONFIG block set `FOLDER_ID`, and `EMAIL_SUBJECT` / `EMAIL_BODY` / `VERIFICATION_SHEET_NAME`
   for the course (MIA 320 is the default; an MKM411 preset is at the bottom of the file).
4. Run `dryRun` → check the Execution Log pairings.
5. Run `createVerificationSheet` → review the sheet of username, email, filename, size.
6. With `TEST_MODE = true` and `TEST_EMAIL` set to yourself, run `sendRubrics` and check your inbox.
7. Set `TEST_MODE = false` and run `sendRubrics` to send to all students.

Google Apps Script has a 6-minute execution limit; ~200 emails complete in about 2 minutes. For
larger classes, split the PDFs across multiple Drive folders and run separately.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Flatten: a backend is unavailable | Default is PyMuPDF (`pip install pymupdf`). Force one with `--backend pymupdf\|pdftk\|pypdf`. |
| Student number field blank after rename | Viewer caching; content is identical. The flatten step resolves it. |
| Calibration tab missing from GA7 report | The console prints what the calibration scan found; detection is case-insensitive. Check the folder is named `Calibration` and holds the agreed rubric PDFs. |
| Upload fails: "students do not match submissions" | Remove rows for students who did not write from the Excel. |
| Upload shows "0 Submission feedback" | Blackboard limitation for offline assignments without submissions. Email the rubrics instead. |
| ZIP too large (>250 MB) | Split into batches and upload sequentially. |
| Apps Script: permission denied | Approve Drive + Gmail access on first run. |
| Apps Script: 6-minute timeout | Split PDFs across multiple Drive folders and run separately. |
| Emails not arriving | Check Gmail Sent; check student spam; verify the `@tuks.co.za` domain. |
