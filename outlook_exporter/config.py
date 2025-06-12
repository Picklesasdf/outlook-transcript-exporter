from dataclasses import dataclass
from pathlib import Path
import configparser
import os
from typing import Optional


@dataclass
class Config:
    keywords: Optional[str] = os.getenv("OUTLOOK_EXPORTER_KEYWORDS")
    output_dir: str = os.getenv("OUTLOOK_EXPORTER_OUTPUT_DIR", ".")
    drive_folder: str = os.getenv("OUTLOOK_EXPORTER_GDRIVE_FOLDER", "")
    credential_path: str = os.getenv("OUTLOOK_EXPORTER_CRED_PATH", "")
    check_interval: int = int(os.getenv("OUTLOOK_EXPORTER_CHECK_INTERVAL", "10"))
    use_ocr: bool = os.getenv("OUTLOOK_EXPORTER_USE_OCR", "false").lower() == "true"
    ocr_backend: str = os.getenv("OUTLOOK_EXPORTER_OCR_BACKEND", "ocrmypdf")
    pages_per_chunk: int = int(os.getenv("OUTLOOK_EXPORTER_PAGES_PER_CHUNK", "10"))
    max_mb: int = int(os.getenv("OUTLOOK_EXPORTER_MAX_MB", "25"))
    merge_backend: str = os.getenv("OUTLOOK_EXPORTER_MERGE_BACKEND", "pymupdf")
    workers: int = int(os.getenv("OUTLOOK_EXPORTER_WORKERS", str(os.cpu_count() or 4)))


def load_config(path: str) -> "Config":
    """Load configuration from ``path`` if it exists."""

    cfg = Config()
    cfg_path = Path(path)
    if cfg_path.exists():
        parser = configparser.ConfigParser()
        parser.read(path)
        section = parser["DEFAULT"]
        cfg.keywords = section.get("keywords", cfg.keywords)
        cfg.output_dir = section.get("output_dir", cfg.output_dir)
        cfg.drive_folder = section.get("drive_folder", cfg.drive_folder)
        cfg.credential_path = section.get("credential_path", cfg.credential_path)
        cfg.check_interval = section.getint("check_interval", cfg.check_interval)
        cfg.use_ocr = section.getboolean("use_ocr", cfg.use_ocr)
        cfg.ocr_backend = section.get("ocr_backend", cfg.ocr_backend)
        cfg.pages_per_chunk = section.getint("pages_per_chunk", cfg.pages_per_chunk)
        cfg.max_mb = section.getint("max_mb", cfg.max_mb)
        cfg.merge_backend = section.get("merge_backend", cfg.merge_backend)
        cfg.workers = section.getint("workers", cfg.workers)
    return cfg
