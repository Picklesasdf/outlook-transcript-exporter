# Outlook Transcript Exporter

A lightweight tool to export Outlook emails into PDFs. This refactored version uses a small package with a command line interface.

## Quick start

```bash
poetry install
poetry run outlook-exporter "invoice" --output-dir results
```

**CUDA users**
```bash
pip install outlook-exporter[gpu]
```
This pulls EasyOCR and the CuPy wheel for CUDA 12.x. If your system uses CUDA 11,
first install:
```bash
pip install cupy-cuda11x
pip install outlook-exporter[gpu]
```

See `--help` for all options.

## Configuration

Create an ``outlook_exporter.ini`` file in the current directory to store default settings:

```ini
[DEFAULT]
keywords = invoice project
output_dir = results
use_ocr = true
ocr_backend = ocrmypdf
pages_per_chunk = 10
workers = 4
merge_backend = pymupdf
drive_folder = abc123
credential_path = creds.json
check_interval = 10
max_mb = 25
```

Any CLI flag will override the file value.

## Performance flags

| Flag | Description | Default |
|------|-------------|---------|
| `--ocr-backend` | ocrmypdf or gpu | ocrmypdf |
| `--pages-per-chunk` | Pages per OCR chunk | 10 |
| `--max-mb` | Split PDFs larger than this many MB | 25 |
| `--merge-backend` | pymupdf or pypdf | pymupdf |
| `--workers` | Parallel OCR workers | CPU count |

Example usage:

```bash
# interactive
poetry run outlook-exporter run

# default with keywords
poetry run outlook-exporter run "IR OAC" --use-ocr

# tuned
poetry run outlook-exporter run "IR" --use-ocr --pages-per-chunk 5 --workers 12

# GPU OCR (pip install outlook-exporter[gpu])
poetry run outlook-exporter run "IR" --use-ocr --ocr-backend gpu
```
