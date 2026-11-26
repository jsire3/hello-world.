"""Rough checks that a conversion didn't lose content.

Phase 1 uses this to warn the user; Phase 2 will use the same scores to pick
the better of two engines automatically.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document

_WS = re.compile(r"\s+")
_W_T = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"


def _norm_len(text: str) -> int:
    return len(_WS.sub("", text))


def docx_text(path: Path) -> str:
    """All text in the body, including tables and text boxes."""
    body = Document(str(path)).element.body
    return "".join(t.text or "" for t in body.iter(_W_T))


def docx_table_count(path: Path) -> int:
    return len(Document(str(path)).tables)


def text_coverage(source_text: str, docx_path: Path) -> float:
    """Fraction of the source's visible characters that made it into the .docx."""
    src = _norm_len(source_text)
    if src == 0:
        return 1.0
    return min(_norm_len(docx_text(docx_path)) / src, 1.0)
