"""PDF utilities for merging and OCR."""

from pathlib import Path
from typing import List


def merge_pdfs(paths: List[str], output_path: str) -> None:
    """Merge PDFs by concatenating bytes (placeholder)."""
    with open(output_path, "wb") as out_f:
        for p in paths:
            with open(p, "rb") as in_f:
                out_f.write(in_f.read())


def ocr_pdf(path: str) -> None:
    """Pretend to OCR by appending a marker."""
    p = Path(path)
    p.write_bytes(p.read_bytes() + b"\nOCR")
