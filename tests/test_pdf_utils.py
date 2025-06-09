from outlook_exporter import pdf_utils


def test_merge_pdfs(tmp_path):
    p1 = tmp_path / "a.pdf"
    p2 = tmp_path / "b.pdf"
    p1.write_bytes(b"PDF1")
    p2.write_bytes(b"PDF2")
    out = tmp_path / "merged.pdf"
    pdf_utils.merge_pdfs([str(p1), str(p2)], str(out))
    assert out.read_bytes() == b"PDF1PDF2"
