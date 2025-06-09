"""PDF utilities for merging and OCR."""

from pathlib import Path
from typing import List
from concurrent.futures import ProcessPoolExecutor
import shutil


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


def fast_merge(paths: List[str], output_path: str) -> None:
    """Merge PDFs using PyMuPDF if available."""
    try:
        import fitz  # type: ignore

        doc = fitz.open()
        for p in paths:
            with fitz.open(p) as src:
                doc.insert_pdf(src)
        doc.save(output_path, deflate=True)
    except Exception:
        merge_pdfs(paths, output_path)


def _ocrmypdf(src: str, dst: str, jobs: int) -> None:
    """Placeholder OCR implementation."""
    shutil.copy(src, dst)
    ocr_pdf(dst)


def smart_ocr(src_pdf: str, dst_pdf: str, pages_per_chunk: int = 10, jobs: int = 2) -> None:
    """Simplified OCR that skips files already containing text marker."""
    data = Path(src_pdf).read_bytes()
    if b"TEXT" in data:
        shutil.copy(src_pdf, dst_pdf)
        return
    with ProcessPoolExecutor(max_workers=1) as exc:
        exc.submit(_ocrmypdf, src_pdf, dst_pdf, jobs).result()
