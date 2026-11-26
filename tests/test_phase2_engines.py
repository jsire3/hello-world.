"""Phase 2: pdf2docx engine, scoring and best-of-both selection."""

import time

import pytest
from docx import Document

from doc2word import scoring
from doc2word.converter import Converter, PdfMode, save_alternate
from doc2word.engines.base import ConversionError
from doc2word.engines.pdf2docx_engine import Pdf2DocxJob


def test_pdf2docx_engine_runs_in_background(table_pdf, tmp_path):
    job = Pdf2DocxJob(table_pdf, tmp_path / "out.docx")
    out = job.wait(60)
    doc = Document(str(out))
    assert len(doc.tables) == 1
    assert [c.text for c in doc.tables[0].rows[1].cells] == ["North", "10", "12", "14"]


def test_pdf2docx_engine_timeout_is_enforced(table_pdf, tmp_path):
    job = Pdf2DocxJob(table_pdf, tmp_path / "out.docx")
    start = time.monotonic()
    with pytest.raises(ConversionError, match="took too long"):
        job.wait(0.01)
    assert not job.proc.is_alive()
    assert time.monotonic() - start < 10


def test_pdf2docx_engine_reports_errors(tmp_path):
    bad = tmp_path / "bad.pdf"
    bad.write_bytes(b"%PDF-1.4 nonsense")
    with pytest.raises(ConversionError, match="pdf2docx could not convert"):
        Pdf2DocxJob(bad, tmp_path / "out.docx").wait(60)


def _flat_docx(path, text):
    """What a weak engine produces: right text, no table, one big text box-free blob."""
    d = Document()
    d.add_paragraph(text)
    d.save(path)
    return path


def test_scoring_rewards_real_tables(table_pdf, tmp_path):
    facts = scoring.pdf_facts(table_pdf)
    assert facts.tables == pytest.approx(1)

    good = Pdf2DocxJob(table_pdf, tmp_path / "good.docx").wait(60)
    flat = _flat_docx(tmp_path / "flat.docx", facts.text)

    s_good, s_flat = scoring.score_docx(facts, good), scoring.score_docx(facts, flat)
    assert s_good.tables == 1.0 and s_flat.tables == 0.0
    assert s_good.coverage > 0.95 and s_flat.coverage > 0.95
    assert scoring.pick_best({"Word": s_flat, "pdf2docx": s_good}) == "pdf2docx"


def test_scoring_catches_scrambled_reading_order(digital_pdf, tmp_path):
    facts = scoring.pdf_facts(digital_pdf)
    in_order = _flat_docx(tmp_path / "a.docx", facts.text)
    scrambled = _flat_docx(tmp_path / "b.docx", " ".join(sorted(facts.text.split())))
    a, b = scoring.score_docx(facts, in_order), scoring.score_docx(facts, scrambled)
    assert a.coverage == pytest.approx(b.coverage, abs=0.01)  # same characters...
    assert a.order > 0.95 > b.order + 0.3                      # ...very different order
    assert a.total > b.total


def test_text_boxes_are_penalised(digital_pdf, tmp_path):
    facts = scoring.pdf_facts(digital_pdf)
    boxed = tmp_path / "boxed.docx"
    d = Document()
    p = d.add_paragraph()
    # Minimal text box: w:r > w:pict > v:shape > v:textbox > w:txbxContent > w:p
    from docx.oxml import parse_xml
    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:shape><v:textbox><w:txbxContent>'
        f'<w:p><w:r><w:t>{facts.text.replace("&", "")[:2000]}</w:t></w:r></w:p>'
        '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>'
    )
    p._p.append(parse_xml(xml))
    d.save(boxed)
    assert scoring.score_docx(facts, boxed).editable < 0.05


def test_near_tie_prefers_word():
    s = scoring.Score(1, 1, 1, 1, 1)
    slightly_better = scoring.Score(1, 1, 1, 1, 1)
    slightly_better.coverage = 1.0
    s.coverage = 0.97  # 0.35 * 0.03 = ~0.01 behind
    assert scoring.pick_best({"pdf2docx": slightly_better, "Word": s}) == "Word"


def test_best_of_both_keeps_the_table_and_stores_alternate(fake_office, table_pdf, tmp_path):
    alt_dir = tmp_path / "alt"
    res = Converter(fake_office, alt_dir=alt_dir).convert(table_pdf)
    # FakeOffice flattens everything into plain paragraphs; pdf2docx keeps the table.
    assert res.engine == "pdf2docx"
    assert set(res.scores) == {"Word", "pdf2docx"}
    assert len(Document(str(res.output)).tables) == 1
    assert res.alternate_engine == "Word" and res.alternate.parent == alt_dir
    assert "pdf2docx" in res.engine_summary and "vs" in res.engine_summary

    other = save_alternate(res)
    assert other.name == "table report (Word).docx" and other.is_file()


def test_digital_text_only_pdf_prefers_word_on_tie(fake_office, digital_pdf, tmp_path):
    res = Converter(fake_office, alt_dir=tmp_path / "alt").convert(digital_pdf)
    assert res.engine == "Word", res.engine_summary


def test_modes_run_only_the_requested_engine(fake_office, table_pdf):
    res = Converter(fake_office, pdf_mode=PdfMode.WORD).convert(table_pdf)
    assert res.engine == "Word" and not res.scores and res.alternate is None
    res = Converter(fake_office).convert(table_pdf, pdf_mode=PdfMode.PDF2DOCX)
    assert res.engine == "pdf2docx" and len(fake_office.calls) == 1


def test_scanned_pdf_skips_pdf2docx(fake_office, scanned_pdf):
    res = Converter(fake_office).convert(scanned_pdf)
    assert res.engine == "Word" and not res.scores


def test_word_failure_falls_back_with_note(table_pdf):
    from conftest import FakeOffice

    res = Converter(FakeOffice(fail_on="input.pdf")).convert(table_pdf)
    assert res.engine == "pdf2docx"
    assert any("Word couldn't convert" in w for w in res.warnings)


def test_old_alternates_are_pruned(tmp_path, fake_office):
    import os

    alt = tmp_path / "alt"
    alt.mkdir()
    old, new = alt / "old.docx", alt / "new.docx"
    old.write_text("x")
    new.write_text("x")
    os.utime(old, (time.time() - 8 * 86400,) * 2)
    Converter(fake_office, alt_dir=alt)
    assert not old.exists() and new.exists()
