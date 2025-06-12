from pathlib import Path
from typing import List
from concurrent.futures import ThreadPoolExecutor, as_completed
from .config import Config
from .logger import logger
from . import outlook, pdf_utils, gdrive


def export(keywords: str, output_dir: Path, config: Config) -> None:
    """Export emails matching keywords to PDF."""

    logger.info("Starting export")
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        emails = outlook.search_mail(keywords)
        email_pdfs: List[Path] = []
        for i, mail_item in enumerate(emails, start=1):
            email_pdf = output_dir / f"email_{i:03}.pdf"
            pdf_utils.write_email_pdf(mail_item, email_pdf)
            email_pdfs.append(email_pdf)

        transcripts = (
            gdrive.download_transcripts(config.drive_folder, config.credential_path)
            if config.drive_folder
            else []
        )

        success_paths: List[Path] = []
        failures: List[str] = []

        with ThreadPoolExecutor(max_workers=config.workers) as pool:
            futures = {
                pool.submit(pdf_utils.ocr_attachment, Path(p)): Path(p)
                for p in transcripts
            }
            for fut in as_completed(futures):
                dst, ok, err = fut.result()
                src = futures[fut]
                if ok and dst:
                    success_paths.append(dst)
                else:
                    failures.append(src.name)
                    logger.error("OCR failed for %s: %s", src, err)

        merge_fn = (
            pdf_utils.fast_merge
            if config.merge_backend == "pymupdf"
            else pdf_utils.merge_pdfs
        )
        all_paths = [str(p) for p in email_pdfs + success_paths]
        merged = output_dir / "final_output.pdf"
        if all_paths:
            merge_fn(all_paths, str(merged))

        if config.use_ocr and not success_paths and not failures:
            logger.info("No attachments to OCR")

        total_emails = len(email_pdfs)
        total_attachments = len(transcripts)
        ocr_ok = len(success_paths)
        ocr_fail = len(failures)
        total_transcripts = len(all_paths)

        logger.info(
            "=== RUN SUMMARY ===\n"
            "Emails processed:      %d\n"
            "Attachments found:     %d\n"
            "  \u2022 OCR succeeded:      %d\n"
            "  \u2022 OCR failed:         %d - %s\n"
            "Transcripts merged:    %d\n"
            "Output PDF:            %s",
            total_emails,
            total_attachments,
            ocr_ok,
            ocr_fail,
            ", ".join(failures),
            total_transcripts,
            merged.name,
        )
        logger.info("Export finished")
    except Exception as exc:  # pragma: no cover - basic error output
        logger.error("Export failed: %s", exc)
        raise
