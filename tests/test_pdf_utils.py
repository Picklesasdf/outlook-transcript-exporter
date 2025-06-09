import os
import sys
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


def test_fast_merge(tmp_path):
    p1 = tmp_path / "a.pdf"
    p2 = tmp_path / "b.pdf"
    p1.write_bytes(b"ONE")
    p2.write_bytes(b"TWO")
    out = tmp_path / "fast.pdf"
    pdf_utils.fast_merge([str(p1), str(p2)], str(out))
    assert b"ONE" in out.read_bytes()


def test_smart_ocr(tmp_path):
    src = tmp_path / "src.pdf"
    src.write_bytes(b"IMAGE")
    dst = tmp_path / "dst.pdf"
    pdf_utils.smart_ocr(str(src), str(dst))
    assert dst.read_bytes().endswith(b"OCR")


def test_smart_ocr_skip(tmp_path):
    src = tmp_path / "text.pdf"
    src.write_bytes(b"TEXT")
    dst = tmp_path / "out.pdf"
    pdf_utils.smart_ocr(str(src), str(dst))
    assert dst.read_bytes() == b"TEXT"
