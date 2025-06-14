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
# default
poetry run outlook-exporter "IR OAC" --use-ocr

# tuned
poetry run outlook-exporter "IR" --use-ocr --pages-per-chunk 5 --workers 12

# GPU OCR (pip install outlook-exporter[gpu])
poetry run outlook-exporter "IR" --use-ocr --ocr-backend gpu
```
