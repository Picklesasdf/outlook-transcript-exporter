import os
import sys
from pathlib import Path
import types

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from outlook_exporter import pdf_utils
try:
    from pypdf import PdfWriter, PdfReader
except ModuleNotFoundError:  # offline fallback
    import types, sys as _sys

    stub = types.ModuleType("pypdf")

    class PdfReader:
        def __init__(self, path):
            self.pages = [type("P", (), {"extract_text": lambda self: ""})()]

    class PdfWriter:
        def __init__(self):
            self.pages = []

        def add_blank_page(self, width: int, height: int):
            self.pages.append(None)

        def add_page(self, page):
            self.pages.append(page)

        def write(self, fh):
            fh.write(b"%PDF-1.4\n%EOF")

    stub.PdfReader = PdfReader
    stub.PdfWriter = PdfWriter
    _sys.modules["pypdf"] = stub
    from pypdf import PdfWriter
    pdf_utils.PdfReader = PdfReader


def _fake_ocr(src: str, dst: str, jobs: int) -> None:
    Path(dst).write_bytes(b"OCR")


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
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    with open(src, "wb") as f:
        writer.write(f)
    dst = tmp_path / "dst.pdf"

    called = False

    def record_ocr(src, dst, jobs):
        nonlocal called
        called = True
        _fake_ocr(src, dst, jobs)

    monkeypatch.setattr(pdf_utils, "_ocrmypdf", record_ocr)

    class DummyExec:
        def __init__(self, *a, **kw):
            pass

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            pass

        def submit(self, fn, *args, **kwargs):
            from concurrent.futures import Future
            fut = Future()
            fut.set_result(fn(*args, **kwargs))
            return fut

    monkeypatch.setattr(pdf_utils, "ProcessPoolExecutor", DummyExec)
    pdf_utils.smart_ocr(str(src), str(dst), pages_per_chunk=1, jobs=1)
    assert called
    assert dst.read_bytes() == b"OCR"


def test_smart_ocr_skip(tmp_path, monkeypatch):
    src = tmp_path / "text.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    with open(src, "wb") as f:
        writer.write(f)
    dst = tmp_path / "out.pdf"

    called = False

    def record_ocr(src, dst, jobs):
        nonlocal called
        called = True
        _fake_ocr(src, dst, jobs)

    monkeypatch.setattr(pdf_utils, "_ocrmypdf", record_ocr)
    # skip OCR because page already has text
    monkeypatch.setattr(
        sys.modules["pypdf"],
        "PdfReader",
        lambda p: types.SimpleNamespace(
            pages=[type("P", (), {"extract_text": lambda self: "hello"})()]
        ),
    )

    class DummyExec:
        def __init__(self, *a, **kw):
            pass

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            pass

        def submit(self, fn, *args, **kwargs):
            from concurrent.futures import Future
            fut = Future()
            fut.set_result(fn(*args, **kwargs))
            return fut

    monkeypatch.setattr(pdf_utils, "ProcessPoolExecutor", DummyExec)
    pdf_utils.smart_ocr(str(src), str(dst), pages_per_chunk=1, jobs=1)
    assert dst.read_bytes() == Path(src).read_bytes()
    assert not called
