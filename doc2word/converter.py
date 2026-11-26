"""Route each input file to the right engine and produce a .docx.

This module has no GUI code, so it is shared by the desktop app, the
command-line tool and the tests.
"""

from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from . import pdfinspect, quality
from .engines.base import ConversionError
from .engines.image import images_to_docx
from .engines.office import TempDir
from .formats import Kind, kind_for
from .pdfinspect import PdfType

log = logging.getLogger(__name__)

# Warn if less than this share of a digital PDF's text reaches the .docx.
MIN_TEXT_COVERAGE = 0.85


class OfficeLike(Protocol):
    """What the converter needs from Office (real: engines.office.OfficeSession)."""

    def word_to_docx(self, src: Path, dest: Path) -> None: ...
    def spreadsheet_to_pdf(self, src: Path, pdf: Path) -> None: ...
    def presentation_to_pdf(self, src: Path, pdf: Path) -> None: ...


@dataclass
class Result:
    output: Path
    kind: Kind
    detail: str = ""             # e.g. "Digital", "Scanned", "3 pages"
    warnings: list[str] = field(default_factory=list)


def unique_path(path: Path) -> Path:
    """``report.docx`` -> ``report (1).docx`` if it already exists, and so on."""
    if not path.exists():
        return path
    n = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def output_path_for(src: Path, out_dir: Path | None) -> Path:
    folder = out_dir or src.parent
    return unique_path(folder / f"{src.stem}.docx")


def describe(src: Path) -> tuple[Kind, str]:
    """Quick look at a file for the queue list: (kind, short detail).

    Kept cheap (page count only) because it runs when files are added; the
    full digital/scanned check happens during conversion.
    """
    kind = kind_for(src)
    if kind != Kind.PDF:
        return kind, kind.value
    try:
        pages = pdfinspect.page_count(src)
    except pdfinspect.PdfPasswordError:
        return kind, "PDF (password protected)"
    except Exception:  # noqa: BLE001
        return kind, "PDF (unreadable)"
    return kind, f"PDF, {pages} page{'s' if pages != 1 else ''}"


class Converter:
    def __init__(self, office: OfficeLike | None):
        self.office = office

    def _need_office(self) -> OfficeLike:
        if self.office is None:
            raise ConversionError(
                "Microsoft Word is needed for this file type but isn't available on this PC."
            )
        return self.office

    def convert(self, src: Path, out_dir: Path | None = None) -> Result:
        src = Path(src)
        if not src.is_file():
            raise ConversionError("The file no longer exists.")
        if out_dir is not None:
            out_dir.mkdir(parents=True, exist_ok=True)

        kind = kind_for(src)
        if kind == Kind.UNSUPPORTED:
            raise ConversionError(f"'{src.suffix}' files aren't supported.")

        with TempDir() as tmp:
            # Work on a private copy with a short, plain name. This avoids
            # Windows path-length limits, odd characters, files locked by
            # another program and slow network drives.
            work_src = tmp / f"input{src.suffix.lower()}"
            shutil.copy2(src, work_src)
            work_out = tmp / "output.docx"

            result = self._dispatch(kind, work_src, work_out, tmp)

            if not work_out.is_file() or work_out.stat().st_size == 0:
                raise ConversionError("The converter produced no output.")
            final = output_path_for(src, out_dir)
            shutil.move(str(work_out), final)
            result.output = final
            return result

    def _dispatch(self, kind: Kind, src: Path, out: Path, tmp: Path) -> Result:
        if kind == Kind.PDF:
            return self._pdf(src, out)
        if kind == Kind.WORD:
            self._need_office().word_to_docx(src, out)
            return Result(out, kind, "Document")
        if kind == Kind.SPREADSHEET:
            pdf = tmp / "sheet.pdf"
            self._need_office().spreadsheet_to_pdf(src, pdf)
            res = self._pdf(pdf, out)
            res.kind, res.detail = kind, "Spreadsheet"
            return res
        if kind == Kind.PRESENTATION:
            pdf = tmp / "slides.pdf"
            self._need_office().presentation_to_pdf(src, pdf)
            res = self._pdf(pdf, out)
            res.kind, res.detail = kind, "Presentation"
            return res
        if kind == Kind.IMAGE:
            pages = images_to_docx(src, out)
            return Result(
                out, kind, f"Image, {pages} page{'s' if pages != 1 else ''}",
                ["Placed as a picture. Editable text from images comes with scan support (Phase 3)."],
            )
        raise ConversionError("Unsupported file type.")  # pragma: no cover

    def _pdf(self, src: Path, out: Path) -> Result:
        try:
            info = pdfinspect.inspect_pdf(src)
        except pdfinspect.PdfPasswordError as exc:
            raise ConversionError(str(exc)) from exc
        except Exception as exc:  # noqa: BLE001
            raise ConversionError("The PDF appears to be damaged.") from exc

        self._need_office().word_to_docx(src, out)

        result = Result(out, Kind.PDF, f"{info.doc_type.value} PDF")
        scanned = info.scanned_pages
        if scanned:
            if info.doc_type == PdfType.SCANNED:
                result.warnings.append(
                    "Scanned PDF: pages are kept as pictures. Editable text from scans comes in Phase 3."
                )
            else:
                result.warnings.append(
                    f"Scanned page(s) {_ranges(scanned)} kept as pictures; the rest is editable."
                )
        if info.doc_type == PdfType.DIGITAL:
            try:
                cov = quality.text_coverage(pdfinspect.extract_text(src), out)
                if cov < MIN_TEXT_COVERAGE:
                    result.warnings.append(
                        f"Only about {cov:.0%} of the PDF's text was found in the Word file - please check it."
                    )
            except Exception:  # noqa: BLE001 - a check, never a failure
                log.exception("Quality check failed for %s", src)
        return result


def _ranges(pages: list[int]) -> str:
    """[1,2,3,5,7,8] -> '1-3, 5, 7-8'"""
    out, start, prev = [], None, None
    for p in pages + [None]:
        if start is None:
            start = prev = p
        elif p is not None and p == prev + 1:
            prev = p
        else:
            out.append(f"{start}" if start == prev else f"{start}-{prev}")
            start = prev = p
    return ", ".join(out)
