from pathlib import Path


def gpu_ocr_to_pdf(
    src: str, dst: str, pages_per_chunk: int, workers: int, timeout: int
) -> None:
    """Placeholder GPU-based OCR. Simply copies the PDF."""

    Path(dst).write_bytes(Path(src).read_bytes())
