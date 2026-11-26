"""Shared fixtures: generated sample files and a fake Office."""

from __future__ import annotations

import io
import os
from pathlib import Path

import pymupdf
import pytest
from docx import Document
from PIL import Image, ImageDraw

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

LOREM = (
    "Quarterly report. Revenue grew in every region and costs were held flat. "
    "The table below summarises the results by team and by month. "
)


def _scan_png(w=1240, h=1754) -> bytes:
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    for y in range(120, h - 120, 40):
        d.rectangle([100, y, w - 100 - (y * 7) % 300, y + 14], fill=(40, 40, 40))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


@pytest.fixture
def digital_pdf(tmp_path: Path) -> Path:
    path = tmp_path / "digital report.pdf"
    doc = pymupdf.open()
    for _ in range(2):
        page = doc.new_page()
        page.insert_textbox(pymupdf.Rect(72, 72, 520, 760), LOREM * 6, fontsize=11)
    doc.save(path)
    return path


@pytest.fixture
def scanned_pdf(tmp_path: Path) -> Path:
    path = tmp_path / "scan.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, stream=_scan_png())
    doc.save(path)
    return path


@pytest.fixture
def mixed_pdf(tmp_path: Path) -> Path:
    path = tmp_path / "mixed.pdf"
    doc = pymupdf.open()
    p1 = doc.new_page()
    p1.insert_textbox(pymupdf.Rect(72, 72, 520, 760), LOREM * 4, fontsize=11)
    p2 = doc.new_page()
    p2.insert_image(p2.rect, stream=_scan_png())
    p3 = doc.new_page()
    p3.insert_image(p3.rect, stream=_scan_png())
    doc.save(path)
    return path


@pytest.fixture
def ocrd_scan_pdf(tmp_path: Path) -> Path:
    """A scan with an invisible OCR text layer on top (render mode 3)."""
    path = tmp_path / "ocrd.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(page.rect, stream=_scan_png())
    page.insert_textbox(pymupdf.Rect(72, 72, 520, 760), LOREM * 4, fontsize=11, render_mode=3)
    doc.save(path)
    return path


@pytest.fixture
def table_pdf(tmp_path: Path) -> Path:
    """Two text columns above a bordered 4x4 table - typical report page."""
    path = tmp_path / "table report.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((72, 60), "Quarterly Results", fontsize=20)
    page.insert_textbox(pymupdf.Rect(72, 80, 290, 330), LOREM * 3, fontsize=10)
    page.insert_textbox(pymupdf.Rect(305, 80, 523, 330), LOREM.replace("Revenue", "Profit") * 3, fontsize=10)
    rows = [["Team", "Jan", "Feb", "Mar"], ["North", "10", "12", "14"],
            ["South", "8", "9", "11"], ["West", "5", "7", "6"]]
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            rect = pymupdf.Rect(72 + c * 110, 350 + r * 22, 72 + (c + 1) * 110, 350 + (r + 1) * 22)
            page.draw_rect(rect, color=(0, 0, 0), width=0.8)
            page.insert_text((rect.x0 + 4, rect.y1 - 7), val, fontsize=10)
    doc.save(path)
    return path


@pytest.fixture
def png_file(tmp_path: Path) -> Path:
    path = tmp_path / "photo.png"
    path.write_bytes(_scan_png(1600, 900))  # landscape
    return path


@pytest.fixture
def jpg_file(tmp_path: Path) -> Path:
    path = tmp_path / "receipt.jpg"
    Image.open(io.BytesIO(_scan_png(800, 1400))).save(path, format="JPEG")
    return path


@pytest.fixture
def multipage_tiff(tmp_path: Path) -> Path:
    path = tmp_path / "fax.tiff"
    frames = [Image.open(io.BytesIO(_scan_png(600, 800))) for _ in range(3)]
    frames[0].save(path, save_all=True, append_images=frames[1:])
    return path


class FakeOffice:
    """Pretends to be Word: writes a .docx containing the PDF/doc text."""

    def __init__(self, fail_on: str | None = None, hang_s: float = 0):
        self.calls: list[tuple[str, Path]] = []
        self.fail_on = fail_on
        self.hang_s = hang_s
        self.killed = False
        self.resets = 0
        self.closed = False

    def word_to_docx(self, src: Path, dest: Path) -> None:
        import time

        from doc2word.engines.base import ConversionError

        self.calls.append(("word", src))
        if self.hang_s:
            end = time.monotonic() + self.hang_s
            while time.monotonic() < end and not self.killed:
                time.sleep(0.02)
            if self.killed:
                raise ConversionError("Office stopped responding while converting this file.")
        if self.fail_on and self.fail_on in src.name:
            raise ConversionError("The file is password protected.")
        doc = Document()
        if src.suffix == ".pdf":
            with pymupdf.open(src) as pdf:
                for page in pdf:
                    doc.add_paragraph(page.get_text())
        else:
            doc.add_paragraph(src.read_text(errors="ignore"))
        doc.save(dest)

    def spreadsheet_to_pdf(self, src: Path, pdf: Path) -> None:
        self.calls.append(("excel", src))
        _text_pdf(pdf, "Sheet1 A1 B1 " * 20)

    def presentation_to_pdf(self, src: Path, pdf: Path) -> None:
        self.calls.append(("ppt", src))
        _text_pdf(pdf, "Slide title and bullet points " * 10)

    def kill(self) -> None:
        self.killed = True

    def reset(self) -> None:
        self.resets += 1
        self.killed = False

    def close(self) -> None:
        self.closed = True


def _text_pdf(path: Path, text: str) -> None:
    doc = pymupdf.open()
    doc.new_page().insert_textbox(pymupdf.Rect(72, 72, 520, 760), text, fontsize=11)
    doc.save(path)


@pytest.fixture
def fake_office() -> FakeOffice:
    return FakeOffice()


@pytest.fixture
def qapp_and_window(tmp_path_factory):
    from PySide6.QtCore import QStandardPaths
    from PySide6.QtWidgets import QApplication

    QStandardPaths.setTestModeEnabled(True)  # keep tests out of real settings
    app = QApplication.instance() or QApplication([])
    from doc2word.gui.main_window import MainWindow

    win = MainWindow(session_factory=FakeOffice, alt_dir=tmp_path_factory.mktemp("alt"))
    win.settings.clear()
    win.save_mode.setCurrentIndex(0)
    yield app, win
    win.queue.shutdown()
    win.deleteLater()
