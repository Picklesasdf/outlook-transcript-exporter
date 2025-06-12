from pathlib import Path
from typing import List, Optional
import typer
from . import main
from .config import load_config

app = typer.Typer(add_completion=False)


@app.command("run")
def run(
    keywords: Optional[List[str]] = typer.Argument(
        None, help="Outlook search keywords"
    ),
    output_dir: Optional[str] = typer.Option(None, help="Where to write PDFs"),
    drive_folder: Optional[str] = typer.Option(None, help="Google Drive folder ID"),
    credential_path: Optional[str] = typer.Option(
        None, help="Path to client_secret.json"
    ),
    check_interval: Optional[int] = typer.Option(
        None, help="Seconds between PDF split checks"
    ),
    use_ocr: Optional[bool] = typer.Option(
        None, help="Run OCR on PDF pages without text"
    ),
    ocr_backend: Optional[str] = typer.Option(None, help="ocrmypdf|gpu"),
    pages_per_chunk: Optional[int] = typer.Option(None, help="Pages per OCR chunk"),
    max_mb: Optional[int] = typer.Option(None, help="Split PDFs larger than MB"),
    merge_backend: Optional[str] = typer.Option(None, help="pymupdf|pypdf"),
    workers: Optional[int] = typer.Option(None, help="Parallel workers for OCR"),
):
    cfg = load_config("outlook_exporter.ini")
    if not keywords:
        if cfg.keywords:
            keywords = cfg.keywords.split()
        else:
            raw = input("Enter Outlook search keywords (space-separated): ")
            keywords = raw.strip().split()

    if output_dir is not None:
        cfg.output_dir = output_dir
    if drive_folder is not None:
        cfg.drive_folder = drive_folder
    if credential_path is not None:
        cfg.credential_path = credential_path
    if check_interval is not None:
        cfg.check_interval = check_interval
    if use_ocr is not None:
        cfg.use_ocr = use_ocr
    if ocr_backend is not None:
        cfg.ocr_backend = ocr_backend
    if pages_per_chunk is not None:
        cfg.pages_per_chunk = pages_per_chunk
    if max_mb is not None:
        cfg.max_mb = max_mb
    if merge_backend is not None:
        cfg.merge_backend = merge_backend
    if workers is not None:
        cfg.workers = workers

    main.export(" ".join(keywords), Path(cfg.output_dir), cfg)


if __name__ == "__main__":
    app()
