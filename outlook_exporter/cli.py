import typer
from . import main
from .config import Config

app = typer.Typer(add_completion=False)


@app.command()
def run(
    keywords: str = typer.Argument(..., help="Space-separated Outlook search keywords"),
    output_dir: str = typer.Option(".", help="Where to write PDFs"),
    drive_folder: str = typer.Option("", help="Google Drive folder ID"),
    credential_path: str = typer.Option("", help="Path to client_secret.json"),
    check_interval: int = typer.Option(10, help="Seconds between PDF split checks"),
    use_ocr: bool = typer.Option(False, help="Run OCR on PDF pages without text"),
):
    config = Config(
        drive_folder=drive_folder,
        credential_path=credential_path,
        check_interval=check_interval,
        use_ocr=use_ocr,
    )
    main.export(keywords, output_dir, config)


if __name__ == "__main__":
    app()
