"""Score a converted .docx against its source PDF, so we can keep the better
of two engines' results.

Each signal is 0..1 (higher is better):

* coverage  - share of the PDF's visible characters present in the .docx
* order     - do the words appear in the same reading order? Catches
              multi-column pages that came out interleaved.
* tables    - share of the PDF's detected tables that became real Word tables
* images    - share of the PDF's pictures that made it across
* editable  - share of text that is NOT trapped in floating text boxes
              (text boxes look right but are painful to edit)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path

import pymupdf
from docx import Document

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic"
VML_IMAGE = "{urn:schemas-microsoft-com:vml}imagedata"

WEIGHTS = {"coverage": 0.35, "order": 0.25, "tables": 0.20, "images": 0.10, "editable": 0.10}

# Detecting tables is the slowest part; sample at most this many pages.
MAX_TABLE_PAGES = 30
# Compare reading order over at most this many words (keeps it fast).
MAX_ORDER_WORDS = 4000
# If scores are this close, prefer Word (cleaner styles, native fonts).
TIE_MARGIN = 0.02

_WORD = re.compile(r"\w+", re.UNICODE)
_WS = re.compile(r"\s+")


def _words(text: str) -> list[str]:
    return [w.lower() for w in _WORD.findall(text)]


@dataclass
class PdfFacts:
    text: str
    words: list[str]
    tables: float   # estimated table count over the whole document
    images: int


@dataclass
class Score:
    coverage: float
    order: float
    tables: float
    images: float
    editable: float

    @property
    def total(self) -> float:
        return sum(getattr(self, k) * w for k, w in WEIGHTS.items())

    def describe(self) -> str:
        return (f"{self.total:.0%} (text {self.coverage:.0%}, order {self.order:.0%}, "
                f"tables {self.tables:.0%}, images {self.images:.0%}, editable {self.editable:.0%})")


def pdf_facts(path: Path) -> PdfFacts:
    with pymupdf.open(path) as doc:
        text = "\n".join(page.get_text() for page in doc)
        images = sum(len(page.get_image_info()) for page in doc)
        n = doc.page_count
        sample = range(min(n, MAX_TABLE_PAGES))
        found = 0
        for i in sample:
            try:
                found += len(doc[i].find_tables().tables)
            except Exception:  # noqa: BLE001 - table finder is best effort
                pass
        tables = found * (n / len(sample)) if len(sample) else 0
    return PdfFacts(text=text, words=_words(text), tables=tables, images=images)


def score_docx(facts: PdfFacts, docx_path: Path) -> Score:
    body = Document(str(docx_path)).element.body

    # One pass over every text run, in document order. Runs are joined
    # directly within a paragraph (a word can be split across runs) and with
    # a newline between paragraphs.
    pieces: list[str] = []
    boxed_chars = 0
    last_para = None
    for t in body.iter(f"{W}t", f"{W}br", f"{W}cr", f"{W}tab"):
        # Line breaks and tabs separate words just like spaces do.
        s = (t.text or "") if t.tag == f"{W}t" else " "
        para, in_box = None, False
        node = t.getparent()
        while node is not None and node is not body:
            if para is None and node.tag == f"{W}p":
                para = node
            if node.tag == f"{W}txbxContent":
                in_box = True
                break
            node = node.getparent()
        if para is not last_para:
            pieces.append("\n")
            last_para = para
        pieces.append(s)
        if in_box:
            boxed_chars += len(_WS.sub("", s))
    doc_text = "".join(pieces)

    src_chars = len(_WS.sub("", facts.text))
    got_chars = len(_WS.sub("", doc_text))
    coverage = 1.0 if src_chars == 0 else min(got_chars / src_chars, 1.0)

    a = facts.words[:MAX_ORDER_WORDS]
    b = _words(doc_text)[:MAX_ORDER_WORDS]
    order = 1.0 if not a else SequenceMatcher(None, a, b, autojunk=False).ratio()

    doc_tables = sum(1 for _ in body.iter(f"{W}tbl"))
    if facts.tables >= 0.5:
        tables = min(doc_tables / facts.tables, 1.0)
    else:
        # No tables in the PDF: some engines use invisible tables for layout.
        # That's fine to read but awkward to edit, so a light penalty.
        tables = 1.0 if doc_tables == 0 else 0.85

    doc_images = sum(1 for _ in body.iter(PIC)) + sum(1 for _ in body.iter(VML_IMAGE))
    images = 1.0 if facts.images == 0 else min(doc_images / facts.images, 1.0)

    editable = 1.0 if got_chars == 0 else 1.0 - boxed_chars / got_chars

    return Score(coverage, order, tables, images, editable)


def pick_best(scores: dict[str, Score], prefer: str = "Word") -> str:
    """Name of the engine with the best total; ``prefer`` wins near-ties."""
    best = max(scores, key=lambda k: scores[k].total)
    if prefer in scores and prefer != best:
        if scores[best].total - scores[prefer].total <= TIE_MARGIN:
            return prefer
    return best
