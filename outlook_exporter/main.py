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
        pdf_utils.merge_pdfs(pdf_paths, str(merged))
        if config.use_ocr:
            pdf_utils.ocr_pdf(str(merged))
        if config.drive_folder:
            gdrive.download_transcripts(config.drive_folder, config.credential_path)
        logger.info("Export finished")
    except Exception as exc:  # pragma: no cover - basic error output
        logger.error("Export failed: %s", exc)
        raise
