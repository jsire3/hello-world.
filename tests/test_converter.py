from pathlib import Path

import pytest
from docx import Document

from doc2word.converter import Converter, _ranges, output_path_for, unique_path
from doc2word.engines.base import ConversionError
from doc2word.formats import Kind


def test_digital_pdf_goes_to_word_without_warnings(fake_office, digital_pdf, tmp_path):
    out_dir = tmp_path / "out"
    res = Converter(fake_office).convert(digital_pdf, out_dir)
    assert res.output == out_dir / "digital report.docx"
    assert res.output.is_file()
    assert res.kind == Kind.PDF and res.detail == "Digital PDF"
    assert res.warnings == []
    # Word received a private copy, not the user's original path.
    assert fake_office.calls[0][1].name == "input.pdf"


def test_scanned_and_mixed_pdfs_warn(fake_office, scanned_pdf, mixed_pdf):
    conv = Converter(fake_office)
    assert "Scanned PDF" in conv.convert(scanned_pdf).warnings[0]
    assert "2-3" in conv.convert(mixed_pdf).warnings[0]


def test_low_text_coverage_warns(digital_pdf):
    class LossyOffice:
        def word_to_docx(self, src, dest):
            d = Document()
            d.add_paragraph("only a little text")
            d.save(dest)

    res = Converter(LossyOffice()).convert(digital_pdf)
    assert any("Only about" in w for w in res.warnings)


def test_output_never_overwrites(fake_office, digital_pdf):
    conv = Converter(fake_office)
    first = conv.convert(digital_pdf).output
    second = conv.convert(digital_pdf).output
    assert first.name == "digital report.docx"
    assert second.name == "digital report (1).docx"


def test_docx_input_does_not_overwrite_itself(fake_office, tmp_path):
    src = tmp_path / "letter.docx"
    Document().save(src)
    assert output_path_for(src, None).name == "letter (1).docx"


def test_office_formats_route_through_office(fake_office, tmp_path):
    conv = Converter(fake_office)
    for name in ("a.txt", "b.xlsx", "c.pptx"):
        p = tmp_path / name
        p.write_text("hello world")
        conv.convert(p)
    assert [c[0] for c in fake_office.calls] == ["word", "excel", "word", "ppt", "word"]


def test_images_need_no_office(png_file, jpg_file, multipage_tiff):
    conv = Converter(None)
    res = conv.convert(png_file)
    assert res.kind == Kind.IMAGE and res.output.is_file()
    doc = Document(str(res.output))
    assert doc.sections[0].page_width > doc.sections[0].page_height  # landscape image
    assert len(doc.inline_shapes) == 1

    assert len(Document(str(conv.convert(jpg_file).output)).inline_shapes) == 1

    tiff = conv.convert(multipage_tiff)
    assert "3 pages" in tiff.detail
    assert len(Document(str(tiff.output)).inline_shapes) == 3


def test_pdf_without_office_fails_clearly(digital_pdf):
    with pytest.raises(ConversionError, match="Microsoft Word is needed"):
        Converter(None).convert(digital_pdf)


def test_errors(fake_office, tmp_path):
    conv = Converter(fake_office)
    with pytest.raises(ConversionError, match="no longer exists"):
        conv.convert(tmp_path / "gone.pdf")
    bad = tmp_path / "broken.pdf"
    bad.write_bytes(b"%PDF-1.7 garbage")
    with pytest.raises(ConversionError, match="damaged"):
        conv.convert(bad)
    z = tmp_path / "a.zip"
    z.write_bytes(b"PK")
    with pytest.raises(ConversionError, match="aren't supported"):
        conv.convert(z)


def test_helpers(tmp_path):
    assert _ranges([1, 2, 3, 5, 7, 8]) == "1-3, 5, 7-8"
    assert _ranges([4]) == "4"
    p = tmp_path / "x.docx"
    assert unique_path(p) == p
    p.write_text("")
    (tmp_path / "x (1).docx").write_text("")
    assert unique_path(p).name == "x (2).docx"
