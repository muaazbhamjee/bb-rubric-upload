#!/usr/bin/env python3
"""
Rename rubric PDFs to Blackboard's feedback file naming convention.

Expected input naming:
    MKM411_ST1_2026_Question123_marking_rubric_XXXXXXXXX.pdf
    (where XXXXXXXXX is the 9-digit student number)

Output naming:
    feedback_uXXXXXXXXX.pdf
    (Blackboard username at UP = "u" + student number)

Usage:
    python rename_for_blackboard.py <input_dir> [output_dir]

If output_dir is not specified, renamed files are saved to <input_dir>/renamed/
"""

import sys
import re
import shutil
from pathlib import Path


# Pattern: 9-digit student number just before .pdf
STUDENT_NUM_PATTERN = re.compile(r'(\d{8})\.pdf$', re.IGNORECASE)


def extract_student_number(filename: str) -> str | None:
    """Extract the 9-digit student number from the rubric filename."""
    match = STUDENT_NUM_PATTERN.search(filename)
    if match:
        return match.group(1)
    return None


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_dir = Path(sys.argv[1])
    if not input_dir.is_dir():
        print(f"Error: '{input_dir}' is not a directory.")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_dir = Path(sys.argv[2])
    else:
        output_dir = input_dir / "renamed"

    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in '{input_dir}'.")
        sys.exit(1)

    print(f"Found {len(pdf_files)} PDF(s) in '{input_dir}'")
    print(f"Renamed output -> '{output_dir}'\n")

    success = 0
    skipped = 0

    for pdf_path in pdf_files:
        student_num = extract_student_number(pdf_path.name)

        if not student_num:
            print(f"  SKIP: {pdf_path.name} (no 9-digit student number found)")
            skipped += 1
            continue

        bb_username = f"u{student_num}"
        new_name = f"feedback_{bb_username}.pdf"
        out_path = output_dir / new_name

        shutil.copy2(pdf_path, out_path)
        print(f"  {pdf_path.name}  ->  {new_name}")
        success += 1

    print(f"\nDone: {success} renamed, {skipped} skipped.")


if __name__ == "__main__":
    main()
