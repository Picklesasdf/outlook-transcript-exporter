from pathlib import Path
from typing import List
from .config import Config
from .logger import logger
from . import outlook, pdf_utils, gdrive


def export(keywords: str, output_dir: str, config: Config) -> None:
    """Export emails matching keywords to PDF."""
    logger.info("Starting export")
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    try:
        emails = outlook.search_mail(keywords)
        pdf_paths: List[str] = []
        for idx, mail in enumerate(emails, start=1):
            pdf_path = Path(output_dir) / f"email_{idx}.pdf"
            # Placeholder: just write the subject to PDF
            pdf_path.write_text(mail)
            pdf_paths.append(str(pdf_path))
        merged = Path(output_dir) / "merged.pdf"
        merge_fn = (
            pdf_utils.fast_merge
            if config.merge_backend == "pymupdf"
            else pdf_utils.merge_pdfs
        )
        merge_fn(pdf_paths, str(merged))
        if config.use_ocr:
            if config.ocr_backend == "gpu":
                from .gpu_ocr import gpu_ocr_to_pdf as ocr_func
            else:
                ocr_func = pdf_utils.smart_ocr

            ocr_func(str(merged), str(merged), config.pages_per_chunk, config.workers)
        if config.drive_folder:
            gdrive.download_transcripts(config.drive_folder, config.credential_path)
        logger.info("Export finished")
    except Exception as exc:  # pragma: no cover - basic error output
        logger.error("Export failed: %s", exc)
        raise
