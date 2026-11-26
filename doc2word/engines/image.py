"""Phase 1 handling for images: place each image on its own Word page.

Editable text from images (OCR + layout + tables) arrives in Phase 3. Until
then the user gets a Word file that looks exactly like the picture, sized to
the page, with the page orientation matching the image.
"""

from __future__ import annotations

import io
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Emu, Inches
from PIL import Image, ImageOps

from .base import ConversionError

try:  # HEIC/HEIF photos from iPhones - optional dependency
    from pillow_heif import register_heif_opener

    register_heif_opener()
except ImportError:  # pragma: no cover
    pass

MARGIN = Inches(0.5)
A4 = (Inches(8.27), Inches(11.69))
LETTER = (Inches(8.5), Inches(11))
MAX_DPI = 300


def _frames(img: Image.Image):
    """Yield every page of a (possibly multi-page) image, upright, RGB or greyscale."""
    for index in range(getattr(img, "n_frames", 1)):
        img.seek(index)
        frame = ImageOps.exif_transpose(img.copy())
        if frame.mode not in ("RGB", "L"):
            frame = frame.convert("RGB")
        yield frame


def _encode(frame: Image.Image, photo: bool, max_px: tuple[int, int]) -> io.BytesIO:
    """Downscale to ~300 dpi at the printed size and encode compactly."""
    if frame.width > max_px[0] or frame.height > max_px[1]:
        frame = frame.copy()
        frame.thumbnail(max_px, Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    if photo:
        frame.save(buf, format="JPEG", quality=92, optimize=True)
    else:
        frame.save(buf, format="PNG", optimize=True)
    buf.seek(0)
    return buf


def images_to_docx(src: Path, dest: Path, paper: str = "A4") -> int:
    """Write ``src`` into ``dest`` as full-page images. Returns the page count."""
    width, height = LETTER if paper.lower() == "letter" else A4
    try:
        img = Image.open(src)
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"The image could not be opened: {exc}") from exc

    doc = Document()
    pages = 0
    photo = (img.format or "").upper() in ("JPEG", "MPO", "HEIF", "WEBP")
    with img:
        for frame in _frames(img):
            section = doc.sections[0] if pages == 0 else doc.add_section()
            landscape = frame.width > frame.height
            section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
            section.page_width, section.page_height = (height, width) if landscape else (width, height)
            for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
                setattr(section, side, MARGIN)

            box_w = section.page_width - 2 * MARGIN
            box_h = section.page_height - 2 * MARGIN
            # Small safety factor so Word never pushes the picture onto a new page.
            scale = min(box_w / frame.width, box_h / frame.height) * 0.98

            max_px = (int(box_w / Inches(1) * MAX_DPI), int(box_h / Inches(1) * MAX_DPI))
            buf = _encode(frame, photo, max_px)
            para = doc.paragraphs[0] if pages == 0 and doc.paragraphs else doc.add_paragraph()
            para.paragraph_format.space_after = 0
            para.add_run().add_picture(
                buf, width=Emu(int(frame.width * scale)), height=Emu(int(frame.height * scale))
            )
            pages += 1

    if pages == 0:
        raise ConversionError("The image contains no pages.")
    doc.save(dest)
    return pages
