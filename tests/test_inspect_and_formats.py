from pathlib import Path

import pymupdf
import pytest

from doc2word.formats import Kind, expand_inputs, kind_for
from doc2word.pdfinspect import PdfPasswordError, PdfType, inspect_pdf, page_count


def test_digital_pdf(digital_pdf):
    info = inspect_pdf(digital_pdf)
    assert info.page_count == 2
    assert info.doc_type == PdfType.DIGITAL
    assert info.scanned_pages == []


def test_scanned_pdf(scanned_pdf):
    assert inspect_pdf(scanned_pdf).doc_type == PdfType.SCANNED


def test_mixed_pdf(mixed_pdf):
    info = inspect_pdf(mixed_pdf)
    assert info.doc_type == PdfType.MIXED
    assert info.scanned_pages == [2, 3]


def test_ocrd_scan_still_counts_as_scanned(ocrd_scan_pdf):
    # Invisible OCR text must not fool us into thinking it's digital.
    assert inspect_pdf(ocrd_scan_pdf).doc_type == PdfType.SCANNED


def test_password_pdf(tmp_path):
    path = tmp_path / "locked.pdf"
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 72), "secret")
    doc.save(path, encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="pw", owner_pw="pw")
    with pytest.raises(PdfPasswordError):
        inspect_pdf(path)
    with pytest.raises(PdfPasswordError):
        page_count(path)


def test_kind_by_extension_and_magic(tmp_path, digital_pdf, png_file):
    assert kind_for(digital_pdf) == Kind.PDF
    assert kind_for(png_file) == Kind.IMAGE
    disguised = tmp_path / "actually_a_pdf.jpg"
    disguised.write_bytes(digital_pdf.read_bytes())
    assert kind_for(disguised) == Kind.PDF
    for name, kind in [("a.docx", Kind.WORD), ("a.rtf", Kind.WORD), ("a.xlsx", Kind.SPREADSHEET),
                       ("a.pptx", Kind.PRESENTATION), ("a.zip", Kind.UNSUPPORTED)]:
        p = tmp_path / name
        p.write_bytes(b"PK\x03\x04 not really")
        assert kind_for(p) == kind, name


def test_expand_inputs(tmp_path, digital_pdf, png_file):
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "notes.txt").write_text("hi")
    (sub / "~$lockfile.docx").write_text("x")
    (sub / "archive.zip").write_bytes(b"x")
    found = expand_inputs([tmp_path, digital_pdf])  # duplicate on purpose
    names = sorted(p.name for p in found)
    assert names == ["digital report.pdf", "notes.txt", "photo.png"]
