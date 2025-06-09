import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from outlook_exporter import pdf_utils


def test_merge_pdfs(tmp_path):
    p1 = tmp_path / "a.pdf"
    p2 = tmp_path / "b.pdf"
    p1.write_bytes(b"PDF1")
    p2.write_bytes(b"PDF2")
    out = tmp_path / "merged.pdf"
    pdf_utils.merge_pdfs([str(p1), str(p2)], str(out))
    assert out.read_bytes() == b"PDF1PDF2"


def test_fast_merge_size(tmp_path):
    p1 = tmp_path / "a.pdf"
    p2 = tmp_path / "b.pdf"
    p1.write_bytes(b"1" * 10)
    p2.write_bytes(b"2" * 10)
    out = tmp_path / "fast.pdf"
    pdf_utils.fast_merge([str(p1), str(p2)], str(out))
    assert out.stat().st_size >= 20


def test_smart_ocr_runs(tmp_path, monkeypatch):
    src = tmp_path / "src.pdf"
    src.write_bytes(b"IMAGE")
    dst = tmp_path / "dst.pdf"

    called = False

    def fake_ocr(src, dst, jobs):
        nonlocal called
        called = True
        Path(dst).write_bytes(b"OCR")

    monkeypatch.setattr(pdf_utils, "_ocrmypdf", fake_ocr)
    pdf_utils.smart_ocr(str(src), str(dst), pages_per_chunk=1, jobs=1)
    assert called
    assert dst.read_bytes() == b"OCR"


def test_smart_ocr_skip(tmp_path, monkeypatch):
    src = tmp_path / "text.pdf"
    src.write_bytes(b"TEXT")
    dst = tmp_path / "out.pdf"

    called = False

    def fake_ocr(src, dst, jobs):
        nonlocal called
        called = True
        Path(dst).write_bytes(b"OCR")

    monkeypatch.setattr(pdf_utils, "_ocrmypdf", fake_ocr)
    pdf_utils.smart_ocr(str(src), str(dst), pages_per_chunk=1, jobs=1)
    assert dst.read_bytes() == b"TEXT"
    assert not called
