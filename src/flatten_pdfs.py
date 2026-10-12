#!/usr/bin/env python3
"""
Flatten fillable PDF rubrics so form fields become static (non-editable) content.

Backends, tried in order of quality:
    1. PyMuPDF  (pymupdf)  - bakes field appearances into page content. Best
                            quality, no system install, cross-platform. Preferred.
    2. pdftk    (pdftk-java) - reliable system flattener if PyMuPDF is absent.
    3. pypdf               - last-resort fallback (strips the AcroForm and marks
                            widgets read-only). Weaker; use only if nothing else.

Runs in parallel across CPU cores with real-time progress reporting.

Usage:
    python flatten_pdfs.py <input_dir> [output_dir] [--backend auto|pymupdf|pdftk|pypdf]

If output_dir is not specified, flattened files are saved to <input_dir>/flattened/

Install the preferred backend with:  pip install pymupdf
"""

import sys
import argparse
import subprocess
import shutil
from pathlib import Path
from multiprocessing import Pool, cpu_count


def flatten_with_pymupdf(args: tuple) -> tuple:
    """Flatten a single PDF using PyMuPDF bake(). Designed for use with Pool."""
    input_path, output_path = args
    try:
        import pymupdf
        doc = pymupdf.open(str(input_path))
        doc.bake()  # render widgets + annotations into static page content
        doc.save(str(output_path), garbage=4, deflate=True)
        doc.close()
        return (input_path.name, True, "")
    except Exception as e:  # noqa: BLE001
        return (input_path.name, False, str(e))


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

    except Exception as e:  # noqa: BLE001
        return (input_path.name, False, str(e))


def pymupdf_available() -> bool:
    try:
        import pymupdf  # noqa: F401
        return True
    except ImportError:
        return False


def choose_backend(requested: str):
    """Return (name, worker_fn) for the backend to use."""
    backends = {
        "pymupdf": ("pymupdf", flatten_with_pymupdf),
        "pdftk": ("pdftk", flatten_with_pdftk),
        "pypdf": ("pypdf", flatten_with_pypdf),
    }
    if requested != "auto":
        return backends[requested]
    if pymupdf_available():
        return backends["pymupdf"]
    if shutil.which("pdftk") is not None:
        return backends["pdftk"]
    return backends["pypdf"]


def main():
    ap = argparse.ArgumentParser(description="Flatten fillable rubric PDFs.",
                                 usage="%(prog)s <input_dir> [output_dir] "
                                       "[--backend auto|pymupdf|pdftk|pypdf]")
    ap.add_argument("input_dir")
    ap.add_argument("output_dir", nargs="?", default=None)
    ap.add_argument("--backend", choices=["auto", "pymupdf", "pdftk", "pypdf"],
                    default="auto")
    args = ap.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"Error: '{input_dir}' is not a directory.")
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else input_dir / "flattened"
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in '{input_dir}'.")
        sys.exit(1)

    total = len(pdf_files)
    backend, flatten_fn = choose_backend(args.backend)
    workers = min(cpu_count(), total, 20)

    print(f"Backend:  {backend}")
    print(f"Workers:  {workers}")
    print(f"Files:    {total}")
    print(f"Output -> '{output_dir}'\n")

    task_args = [(p, output_dir / p.name) for p in pdf_files]
    success = failed = 0
    with Pool(workers) as pool:
        for name, ok, err in pool.imap_unordered(flatten_fn, task_args):
            if ok:
                success += 1
                print(f"  [{success + failed}/{total}] {name} ... OK")
            else:
                failed += 1
                print(f"  [{success + failed}/{total}] {name} ... FAILED ({err})")

    print(f"\nDone: {success} flattened, {failed} failed.")
    if failed and backend != "pypdf":
        print("Tip: re-run with --backend pypdf if a backend is unavailable.")


if __name__ == "__main__":
    main()
