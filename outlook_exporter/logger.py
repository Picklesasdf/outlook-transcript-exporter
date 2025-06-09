import logging
from logging.handlers import RotatingFileHandler
import os

LOG_LEVEL = os.getenv("OUTLOOK_EXPORTER_LOG_LEVEL", "INFO")
LOG_FILE = os.getenv("OUTLOOK_EXPORTER_LOG_FILE", "outlook_exporter.log")

logger = logging.getLogger("outlook_exporter")
logger.setLevel(LOG_LEVEL)
_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
_handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3)
_handler.setFormatter(_formatter)
logger.addHandler(_handler)

__all__ = ["logger"]
