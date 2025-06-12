"""Google Drive helper functions."""

from typing import List
from tqdm import tqdm


def download_transcripts(folder_id: str, credential_path: str) -> List[str]:
    """Placeholder download showing a progress bar."""
    files: List[str] = []
    for _ in tqdm(range(0), desc="download", unit="file"):
        pass
    return files
