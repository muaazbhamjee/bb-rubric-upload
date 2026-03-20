# Blackboard Ultra — Bulk Grade & Rubric Upload

Workflow for bulk uploading semester test grades and emailing per-student rubric PDFs for courses at UP.

**Grades** are uploaded via the Blackboard Marks and Feedback Tool.
**Rubric PDFs** are emailed to students via Google Apps Script (Blackboard does not attach feedback files to offline assignments without existing submissions).

---

## Prerequisites

### Python Packages

```bash
pip install openpyxl
```

Only third-party Python package required (used by `package_for_upload.py`). All other scripts use only the Python 3.10+ standard library.

If `pdftk` is not available (see below), the flatten script falls back to `pypdf`:

```bash
pip install pypdf
```

### System Packages — pdftk

Used by `flatten_pdfs.py` to burn fillable form fields into static PDF content.

**Linux (Ubuntu/Debian/WSL):**

```bash
sudo apt update
sudo apt install pdftk-java
```

**macOS (Homebrew):**

```bash
brew install pdftk-java
```

If Homebrew is not installed, get it from [brew.sh](https://brew.sh/) or download pdftk from [pdflabs.com](https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/).

**Windows (native):**

1. Download the installer from [pdflabs.com](https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/)
2. Run the installer — adds `pdftk` to PATH automatically
3. Verify: `pdftk --version`

**Windows (WSL — recommended if already using WSL):**

```bash
sudo apt update
sudo apt install pdftk-java
```

### Google Apps Script

No installation needed — runs in browser at [script.google.com](https://script.google.com) using your UP Google Workspace account. Requires access to Google Drive and Gmail.

### Verify Installation

```bash
python --version        # Should be 3.10+
pdftk --version         # Should print version info
python -c "import openpyxl; print(openpyxl.__version__)"
```

---

## Scripts

| Script | Purpose | Runs on |
|---|---|---|
| `rename_for_blackboard.py` | Renames rubric PDFs to `feedback_uXXXXXXXX.pdf` format | Python (WSL/Linux/Mac/Win) |
| `flatten_pdfs.py` | Flattens fillable PDFs so form fields become static (parallel) | Python (WSL/Linux/Mac/Win) |
| `package_for_upload.py` | Creates Blackboard-compatible upload ZIP with Excel template | Python (WSL/Linux/Mac/Win) |
| `email.js` | Emails rubric PDFs to students via Gmail | Google Apps Script (browser) |

## Rubric Naming Convention

```
MKM411_ST1_2026_Question123_marking_rubric_XXXXXXXX.pdf
                                           ^^^^^^^^
                                           8-digit student number
```

Blackboard username at UP: `u` + student number (e.g. `u12345678`).
Student email at UP: `uXXXXXXXX@tuks.co.za`.

---

## Blackboard Setup

1. Create an **offline assignment** in your course — it appears on the Content page
2. Access the Marks and Feedback Tool via **Content Market → Marks and Feedback**
3. Note the **assignment name** and **course ID** — needed for the packaging step

> **Note:** The Marks and Feedback Tool is an LTI integration that must be enabled at institutional level. If you don't see it, contact UP's e-learning support.

---

## Pipeline

### Step 1 — Rename

Converts rubric filenames to Blackboard/email feedback format.

```bash
python rename_for_blackboard.py /path/to/filled-rubrics/
```

Output: `renamed/` subfolder with files like `feedback_u12345678.pdf`.

### Step 2 — Flatten

Burns form field values into the page as static content so students cannot edit the rubric.

```bash
python flatten_pdfs.py /path/to/filled-rubrics/renamed/
```

Output: `flattened/` subfolder. Runs in parallel — adjust the worker cap in the script if needed (default is `min(cpu_count(), total, 10)`; 20 works well on WSL with 10+ cores).

Open a flattened PDF to verify all fields are visible and non-editable.

### Step 3 — Upload Grades to Blackboard

#### 3a — Package

Creates the upload ZIP with an Excel template for grades.

```bash
python package_for_upload.py /path/to/renamed/flattened/ "Assignment Name" "CourseID" output.zip
```

Example:

```bash
python package_for_upload.py rubrics/renamed/flattened/ "MKM411 Semester Test 1" "MKM411_2026" MKM411_ST1_upload.zip
```

#### 3b — Fill in the Excel

1. Extract the ZIP
2. Open the Excel file and fill in `FIRSTNAME`, `LASTNAME`, `Grade`, and optionally `Feedback` columns
3. Remove rows for students who did not write
4. **Do NOT rename the Excel file**
5. Re-zip everything preserving the folder structure:

**Linux / macOS / WSL:**

```bash
cd extracted_folder/
zip -r ../MKM411_ST1_upload.zip *
```

**Windows (PowerShell):**

```powershell
Compress-Archive -Path .\extracted_folder\* -DestinationPath .\MKM411_ST1_upload.zip
```

**ZIP constraints:**
- Filename must not contain spaces
- Must be under 250 MB (split into batches if exceeded)

#### 3c — Upload

1. Go to **Content Market → Marks and Feedback**
2. Click **Upload** next to your offline assignment
3. Browse to the ZIP and click **Submit**
4. Check the upload report — grades should show as found for all students

#### 3d — Release Grades

1. Click **Release Grades** in the Marks and Feedback Tool
2. Verify via **Student Preview** mode

### Step 4 — Email Rubric PDFs to Students

#### 4a — Upload PDFs to Google Drive

Upload all flattened `feedback_u*.pdf` files to a folder in Google Drive (e.g. `MKM411_ST1_Rubrics`).

Copy the **folder ID** from the URL:
```
https://drive.google.com/drive/folders/1aBcDeFgHiJkLmNoPqRsTuVwXyZ
                                       ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                                       this is the folder ID
```

#### 4b — Set up Google Apps Script

1. Go to [script.google.com](https://script.google.com) → **New Project**
2. Delete everything and paste the contents of `email.js`
3. Replace `PASTE_YOUR_FOLDER_ID_HERE` with your folder ID
4. Update `EMAIL_SUBJECT`, `EMAIL_BODY`, and `SENDER_NAME` as needed
5. Approve permissions when prompted (Drive + Gmail access)

#### 4c — Verify

1. Select `dryRun` from the function dropdown → **Run** → check Execution Log for correct pairings
2. Select `createVerificationSheet` → **Run** → opens a Google Sheet listing every username, email, filename, and file size — review for correctness

#### 4d — Test

1. Ensure `TEST_MODE = true` and `TEST_EMAIL` is set to your own email
2. Select `sendRubrics` → **Run**
3. Check your inbox — verify subject line includes student number, rubric PDF is attached, and email body is correct

#### 4e — Send to all students

1. Set `TEST_MODE = false`
2. Select `sendRubrics` → **Run**
3. Monitor the Execution Log — should complete in ~2 minutes for 200 students

> **Note:** Google Apps Script has a 6-minute execution limit. At 200 emails with 0.5s delay this completes in ~2 minutes. For larger classes, split the PDFs across multiple Drive folders and run separately.

---

## Quick Reference

```bash
# Step 1-2: Rename and flatten
python rename_for_blackboard.py  rubrics/filled/
python flatten_pdfs.py           rubrics/filled/renamed/

# Step 3: Package and upload grades to Blackboard
python package_for_upload.py     rubrics/filled/renamed/flattened/  "MKM411 Semester Test 1"  "MKM411_2026"  MKM411_ST1_upload.zip
unzip MKM411_ST1_upload.zip -d upload_staging/
# ... edit Excel: add names, grades, remove absent students ...
cd upload_staging/ && zip -r ../MKM411_ST1_upload.zip *
# Upload via Content Market -> Marks and Feedback -> Upload
# Release Grades

# Step 4: Email rubrics via Google Apps Script
# Upload flattened PDFs to Google Drive folder
# Run email.js functions in order: dryRun -> createVerificationSheet -> sendRubrics (test) -> sendRubrics (live)
```

## Troubleshooting

| Problem | Fix |
|---|---|
| Marks and Feedback Tool not visible | LTI must be enabled at institutional level — contact UP e-learning support |
| Upload fails: "students do not match submissions" | Remove rows for students who did not write from the Excel |
| Upload shows "0 Submission feedback" | Known limitation — Blackboard does not attach feedback files to offline assignments without existing submissions. Use email instead (Step 4) |
| ZIP too large (>250 MB) | Split into batches — upload multiple ZIPs sequentially |
| `pdftk` not found | Install: `sudo apt install pdftk-java` (Linux/WSL), `brew install pdftk-java` (macOS), or [pdflabs.com](https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/) (Windows) |
| Student number field blank after rename | Viewer caching issue — file content is identical (verify with `md5sum`). Flatten step resolves it |
| Flatten slow | Increase worker count in `flatten_pdfs.py` (default caps at 10; 20 works well on WSL with 10+ cores) |
| `openpyxl` not found | `pip install openpyxl` — only needed for `package_for_upload.py` |
| `pypdf` not found (fallback) | `pip install pypdf` — only needed if `pdftk` is not installed |
| Google Apps Script: permission denied | Approve Drive + Gmail access when prompted on first run |
| Google Apps Script: 6-minute timeout | Split PDFs across multiple Drive folders and run separately |
| Emails not arriving | Check Gmail Sent folder. Check spam on student side. Verify `@tuks.co.za` is correct domain |