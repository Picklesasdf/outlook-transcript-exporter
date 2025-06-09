from dataclasses import dataclass
import os


@dataclass
class Config:
    drive_folder: str = os.getenv("OUTLOOK_EXPORTER_GDRIVE_FOLDER", "")
    credential_path: str = os.getenv("OUTLOOK_EXPORTER_CRED_PATH", "")
    check_interval: int = int(os.getenv("OUTLOOK_EXPORTER_CHECK_INTERVAL", "10"))
    use_ocr: bool = os.getenv("OUTLOOK_EXPORTER_USE_OCR", "false").lower() == "true"
