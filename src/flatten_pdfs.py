#!/usr/bin/env python3
"""
Flatten fillable PDF rubrics so form fields become static (non-editable) content.

Uses pdftk (preferred, most reliable) with a pypdf fallback.
Runs in parallel using multiple CPU cores with real-time progress reporting.

Usage:
    python flatten_pdfs.py <input_dir> [output_dir]

If output_dir is not specified, flattened files are saved to <input_dir>/flattened/
"""

import sys
import subprocess
import shutil
from pathlib import Path
from multiprocessing import Pool, cpu_count


def flatten_with_pdftk(args: tuple) -> tuple:
    """Flatten a single PDF using pdftk. Designed for use with Pool."""
    input_path, output_path = args
    try:
        result = subprocess.run(
            ["pdftk", str(input_path), "output", str(output_path), "flatten"],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            return (input_path.name, False, result.stderr.strip())
        return (input_path.name, True, "")
    except FileNotFoundError:
        return (input_path.name, False, "pdftk not found")
    except subprocess.TimeoutExpired:
        return (input_path.name, False, "timed out")


def flatten_with_pypdf(args: tuple) -> tuple:
    """Fallback: flatten using pypdf. Designed for use with Pool."""
    input_path, output_path = args
    try:
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import NameObject, NumberObject

        reader = PdfReader(str(input_path))
        writer = PdfWriter()
        writer.clone_document_from_reader(reader)

        if "/AcroForm" in writer._root_object:
            del writer._root_object["/AcroForm"]

        for page in writer.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    annot = annot_ref.get_object()
                    if annot.get("/Subtype") == "/Widget":
                        flags = int(annot.get("/Ff", 0))
                        annot[NameObject("/Ff")] = NumberObject(flags | 1)

        with open(str(output_path), "wb") as f:
            writer.write(f)
        return (input_path.name, True, "")

    except Exception as e:
        return (input_path.name, False, str(e))


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
        output_dir = input_dir / "flattened"

    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in '{input_dir}'.")
        sys.exit(1)

    total = len(pdf_files)
    has_pdftk = shutil.which("pdftk") is not None
    backend = "pdftk" if has_pdftk else "pypdf"
    workers = min(cpu_count(), total, 20)
    flatten_fn = flatten_with_pdftk if has_pdftk else flatten_with_pypdf

    print(f"Backend:  {backend}")
    print(f"Workers:  {workers}")
    print(f"Files:    {total}")
    print(f"Output -> '{output_dir}'\n")

    task_args = [(pdf_path, output_dir / pdf_path.name) for pdf_path in pdf_files]

    success = 0
    failed = 0

    with Pool(workers) as pool:
        for name, ok, err in pool.imap_unordered(flatten_fn, task_args):
            if ok:
                success += 1
                print(f"  [{success + failed}/{total}] {name} ... OK")
            else:
                failed += 1
                print(f"  [{success + failed}/{total}] {name} ... FAILED ({err})")

    print(f"\nDone: {success} flattened, {failed} failed.")


if __name__ == "__main__":
    main()