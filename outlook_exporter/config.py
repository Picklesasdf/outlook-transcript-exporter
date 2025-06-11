from dataclasses import dataclass
import os


@dataclass
class Config:
    drive_folder: str = os.getenv("OUTLOOK_EXPORTER_GDRIVE_FOLDER", "")
    credential_path: str = os.getenv("OUTLOOK_EXPORTER_CRED_PATH", "")
    check_interval: int = int(os.getenv("OUTLOOK_EXPORTER_CHECK_INTERVAL", "10"))
    use_ocr: bool = os.getenv("OUTLOOK_EXPORTER_USE_OCR", "false").lower() == "true"
    ocr_backend: str = os.getenv("OUTLOOK_EXPORTER_OCR_BACKEND", "ocrmypdf")
    pages_per_chunk: int = int(os.getenv("OUTLOOK_EXPORTER_PAGES_PER_CHUNK", "10"))
    max_mb: int = int(os.getenv("OUTLOOK_EXPORTER_MAX_MB", "25"))
    merge_backend: str = os.getenv("OUTLOOK_EXPORTER_MERGE_BACKEND", "pymupdf")
    workers: int = int(os.getenv("OUTLOOK_EXPORTER_WORKERS", str(os.cpu_count() or 4)))
