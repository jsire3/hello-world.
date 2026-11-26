"""Look inside a PDF and decide, page by page, whether it is digital or scanned.

A *digital* page has a real text layer (exported from Word, a browser, etc.).
A *scanned* page is essentially a picture of a page with little or no text.
A scanned page that was previously OCR'd has invisible text over an image; we
still treat that as scanned, because the layout lives in the picture.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

import pymupdf


class PdfType(str, Enum):
    DIGITAL = "Digital"
    SCANNED = "Scanned"
    MIXED = "Mixed"
    EMPTY = "Empty"


# A page with fewer visible characters than this has "no real text".
MIN_TEXT_CHARS = 25
# Images covering at least this fraction of the page make it look like a scan.
SCAN_IMAGE_COVERAGE = 0.6


@dataclass
class PdfInfo:
    page_count: int
    page_types: list[PdfType] = field(default_factory=list)
    encrypted: bool = False

    @property
    def doc_type(self) -> PdfType:
        kinds = {t for t in self.page_types if t != PdfType.EMPTY}
        if not kinds:
            return PdfType.EMPTY
        if kinds == {PdfType.DIGITAL}:
            return PdfType.DIGITAL
        if kinds == {PdfType.SCANNED}:
            return PdfType.SCANNED
        return PdfType.MIXED

    @property
    def scanned_pages(self) -> list[int]:
        """1-based page numbers that look scanned."""
        return [i + 1 for i, t in enumerate(self.page_types) if t == PdfType.SCANNED]


class PdfPasswordError(Exception):
    pass


def _visible_text_chars(page: pymupdf.Page) -> int:
    """Count characters that are actually drawn (ignores invisible OCR text)."""
    count = 0
    for trace in page.get_texttrace():
        # Render mode 3 = invisible text, typical of OCR layers.
        if trace.get("type") == 3 or trace.get("opacity", 1) == 0:
            continue
        count += sum(1 for ch in trace["chars"] if chr(ch[0]).strip())
    return count


def _image_coverage(page: pymupdf.Page) -> float:
    page_area = abs(page.rect)
    if page_area <= 0:
        return 0.0
    covered = 0.0
    for info in page.get_image_info():
        bbox = pymupdf.Rect(info["bbox"]) & page.rect
        covered += abs(bbox)
    return min(covered / page_area, 1.0)


def classify_page(page: pymupdf.Page) -> PdfType:
    chars = _visible_text_chars(page)
    if chars >= MIN_TEXT_CHARS:
        return PdfType.DIGITAL
    if _image_coverage(page) >= SCAN_IMAGE_COVERAGE:
        return PdfType.SCANNED
    if chars > 0 or page.get_drawings():
        # A sparse page (title page, diagram) - still digital.
        return PdfType.DIGITAL
    return PdfType.EMPTY


def inspect_pdf(path: Path) -> PdfInfo:
    with pymupdf.open(path) as doc:
        if doc.needs_pass:
            raise PdfPasswordError("This PDF is password protected.")
        info = PdfInfo(page_count=doc.page_count, encrypted=doc.is_encrypted)
        for page in doc:
            info.page_types.append(classify_page(page))
        return info


def extract_text(path: Path) -> str:
    """Plain text of a PDF; used to check conversion quality."""
    with pymupdf.open(path) as doc:
        return "\n".join(page.get_text() for page in doc)


def page_count(path: Path) -> int:
    with pymupdf.open(path) as doc:
        if doc.needs_pass:
            raise PdfPasswordError("This PDF is password protected.")
        return doc.page_count
