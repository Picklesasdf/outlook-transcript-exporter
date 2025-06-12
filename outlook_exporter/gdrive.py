"""Google Drive helper functions."""

from typing import List
from pathlib import Path
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


def download_transcripts(folder_id: str, credential_path: str) -> List[Path]:
    """Download Google Docs in ``folder_id`` as PDFs."""

    paths: List[Path] = []
    if not folder_id or not credential_path:
        return paths

    try:
        creds = service_account.Credentials.from_service_account_file(
            credential_path,
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        service = build("drive", "v3", credentials=creds)
        resp = (
            service.files()
            .list(
                q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.document'",
                fields="files(id,name)",
            )
            .execute()
        )
        for item in resp.get("files", []):
            try:
                request = service.files().export_media(
                    fileId=item["id"], mimeType="application/pdf"
                )
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                path = Path(item["name"] + ".pdf")
                path.write_bytes(fh.getvalue())
                paths.append(path)
            except Exception:
                continue
    except Exception:
        pass
    return paths
