"""Which file types we accept and how each one is routed."""

from __future__ import annotations

from enum import Enum
from pathlib import Path


class Kind(str, Enum):
    PDF = "PDF"
    WORD = "Document"          # anything Word opens natively
    SPREADSHEET = "Spreadsheet"  # exported via Excel first
    PRESENTATION = "Presentation"  # exported via PowerPoint first
    IMAGE = "Image"
    UNSUPPORTED = "Unsupported"


WORD_EXTS = {
    ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm",
    ".rtf", ".odt", ".txt", ".htm", ".html", ".mht", ".mhtml",
    ".xml", ".wpd", ".wps",
}
SPREADSHEET_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv"}
PRESENTATION_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".odp"}
IMAGE_EXTS = {
    ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp",
    ".heic", ".heif",
}
PDF_EXTS = {".pdf"}

SUPPORTED_EXTS = WORD_EXTS | SPREADSHEET_EXTS | PRESENTATION_EXTS | IMAGE_EXTS | PDF_EXTS

_MAGIC = [
    (b"%PDF", Kind.PDF),
    (b"\xff\xd8\xff", Kind.IMAGE),       # JPEG
    (b"\x89PNG\r\n\x1a\n", Kind.IMAGE),  # PNG
    (b"II*\x00", Kind.IMAGE),            # TIFF little-endian
    (b"MM\x00*", Kind.IMAGE),            # TIFF big-endian
    (b"BM", Kind.IMAGE),                 # BMP
    (b"GIF8", Kind.IMAGE),
]


def kind_for(path: Path) -> Kind:
    """Classify a file by extension, falling back to magic bytes.

    Magic bytes win over the extension for PDFs and images, so a PDF saved
    as ``scan.jpg`` (it happens) still goes down the right path.
    """
    try:
        with open(path, "rb") as fh:
            head = fh.read(16)
    except OSError:
        head = b""

    for magic, kind in _MAGIC:
        if head.startswith(magic):
            # "BM" is only two bytes; trust it only if the extension agrees.
            if magic == b"BM" and path.suffix.lower() != ".bmp":
                continue
            return kind
    if head[4:12] in (b"ftypheic", b"ftypheix", b"ftypmif1", b"ftypheif"):
        return Kind.IMAGE
    if head.startswith(b"RIFF") and head[8:12] == b"WEBP":
        return Kind.IMAGE

    ext = path.suffix.lower()
    if ext in PDF_EXTS:
        return Kind.PDF
    if ext in WORD_EXTS:
        return Kind.WORD
    if ext in SPREADSHEET_EXTS:
        return Kind.SPREADSHEET
    if ext in PRESENTATION_EXTS:
        return Kind.PRESENTATION
    if ext in IMAGE_EXTS:
        return Kind.IMAGE
    return Kind.UNSUPPORTED


def is_supported(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_EXTS


def expand_inputs(paths: list[Path]) -> list[Path]:
    """Expand dropped folders recursively; keep only supported, real files."""
    out: list[Path] = []
    seen: set[Path] = set()
    for p in paths:
        candidates = sorted(p.rglob("*")) if p.is_dir() else [p]
        for c in candidates:
            if not c.is_file() or not is_supported(c):
                continue
            # Skip Office lock files like "~$report.docx".
            if c.name.startswith("~$"):
                continue
            key = c.resolve()
            if key not in seen:
                seen.add(key)
                out.append(c)
    return out
