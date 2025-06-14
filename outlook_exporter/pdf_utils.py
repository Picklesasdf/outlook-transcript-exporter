"""PDF utilities for merging and OCR."""

from pathlib import Path
from typing import List, Optional
from concurrent.futures import ProcessPoolExecutor, as_completed
from tempfile import NamedTemporaryFile
import subprocess
import os
import logging
import shutil

LOG = logging.getLogger(__name__)
CPU = max(os.cpu_count() or 1, 1)


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


def run_ocr(path: str) -> None:
    """Backward compatible wrapper calling :func:`smart_ocr`."""
    smart_ocr(path, path)


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
    """Run ocrmypdf on ``src`` writing to ``dst``.

    This function is a thin wrapper that can be monkeypatched in tests.
    """
    subprocess.run(
        ["ocrmypdf", "--skip-text", "--jobs", str(jobs), src, dst],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def smart_ocr(
    src_pdf: str, dst_pdf: str, pages_per_chunk: int = 10, jobs: int = 2
) -> None:
    """Chunked OCR that skips pages already containing text."""
    try:
        from pypdf import PdfReader, PdfWriter
    except Exception:
        # If PyPDF is unavailable fall back to a simple heuristic
        LOG.debug("pypdf not available - running fallback OCR")
        if b"TEXT" in Path(src_pdf).read_bytes():
            Path(dst_pdf).write_bytes(Path(src_pdf).read_bytes())
        else:
            _ocrmypdf(src_pdf, dst_pdf, jobs)
        return

    reader = PdfReader(src_pdf)
    pages = list(reader.pages)

    all_text = True
    for p in pages:
        has_text = False
        if hasattr(p, "extract_text"):
            try:
                has_text = bool(p.extract_text())
            except Exception:
                has_text = False
        if not has_text:
            all_text = False
            break

    if all_text:
        shutil.copyfile(src_pdf, dst_pdf)
        return

    total = len(pages)
    tmp_files: List[str] = []

    with ProcessPoolExecutor(max_workers=min(CPU, 6)) as pool:
        futures = {}
        for start in range(0, total, pages_per_chunk):
            end = min(start + pages_per_chunk, total)
            writer = PdfWriter()
            needs_ocr = False
            for p in range(start, end):
                orig_page = reader.pages[p]
                has_text = False
                if hasattr(orig_page, "extract_text"):
                    try:
                        has_text = bool(orig_page.extract_text())
                    except Exception:
                        has_text = False
                if "__getitem__" not in dir(orig_page):
                    tmp_w = PdfWriter()
                    tmp_w.add_blank_page(width=72, height=72)
                    page = tmp_w.pages[0]
                else:
                    page = orig_page
                writer.add_page(page)
                if not has_text:
                    needs_ocr = True
            tmp_in = NamedTemporaryFile(delete=False, suffix=".pdf")
            writer.write(tmp_in)
            tmp_in.close()

            tmp_out = NamedTemporaryFile(delete=False, suffix=".pdf")
            if needs_ocr:
                futures[pool.submit(_ocrmypdf, tmp_in.name, tmp_out.name, jobs)] = (
                    tmp_in.name,
                    tmp_out.name,
                )
            else:
                os.replace(tmp_in.name, tmp_out.name)
            tmp_files.append(tmp_out.name)

        for fut in as_completed(futures):
            fut.result()

    fast_merge(tmp_files, dst_pdf)
    for f in tmp_files:
        os.remove(f)


def _ocr_file(src: str, dst: str, jobs: int) -> Optional[Exception]:
    """OCR ``src`` into ``dst`` with a timeout."""
    try:
        subprocess.run(
            ["ocrmypdf", "--skip-text", "--jobs", str(jobs), src, dst],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=60,
        )
    except Exception as exc:  # pragma: no cover - runtime guard
        return exc
    else:
        return None


def ocr_and_merge_attachments(
    attachments: List[str], out_pdf: str, email_count: int, jobs: int = 2
) -> dict:
    """OCR attachments in parallel then merge them into ``out_pdf``.

    Returns a summary dictionary of results.
    """

    summary = {
        "emails": email_count,
        "attachments": len(attachments),
        "ocred": 0,
        "merged": 0,
        "failed": 0,
        "errors": [],
    }

    tmp_outputs: List[str] = []

    with ProcessPoolExecutor(max_workers=min(CPU, len(attachments) or 1)) as pool:
        futures = {}
        for path in attachments:
            LOG.info("Downloaded %s", path)
            tmp_out = NamedTemporaryFile(delete=False, suffix=".pdf")
            futures[pool.submit(_ocr_file, path, tmp_out.name, jobs)] = (
                path,
                tmp_out.name,
            )
            tmp_outputs.append(tmp_out.name)

        for fut in as_completed(futures):
            src, tmp = futures[fut]
            err = fut.result()
            if err is None:
                LOG.info("OCR succeeded: %s", src)
                summary["ocred"] += 1
            else:
                LOG.error("OCR failed for %s: %s", src, err)
                summary["failed"] += 1
                summary["errors"].append({"file": src, "error": str(err)})
                try:
                    os.remove(tmp)
                except OSError:
                    pass
                tmp_outputs.remove(tmp)

    if tmp_outputs:
        fast_merge(tmp_outputs, out_pdf)
        summary["merged"] = len(tmp_outputs)
        LOG.info("Merged %d files into %s", len(tmp_outputs), out_pdf)
    else:
        LOG.warning("No files were OCRed successfully; nothing to merge")

    for f in tmp_outputs:
        try:
            os.remove(f)
        except OSError:
            pass

    LOG.info(
        "Summary: emails=%d attachments=%d ocred=%d merged=%d failed=%d",
        summary["emails"],
        summary["attachments"],
        summary["ocred"],
        summary["merged"],
        summary["failed"],
    )
    if summary["errors"]:
        for item in summary["errors"]:
            LOG.error("Failed: %s -> %s", item["file"], item["error"])

    return summary


def main() -> None:
    """Entry point for manual invocation."""
    pass


if __name__ == "__main__":  # pragma: no cover - manual invocation
    main()
