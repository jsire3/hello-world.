"""Route each input file to the right engine and produce a .docx.

This module has no GUI code, so it is shared by the desktop app, the
command-line tool and the tests.
"""

from __future__ import annotations

import logging
import shutil
import threading
import time
import uuid
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Protocol

from . import pdfinspect, quality, scoring
from .engines.base import ConversionError
from .engines.image import images_to_docx
from .engines.office import TempDir
from .engines.pdf2docx_engine import ENGINE_NAME as P2D, Pdf2DocxJob
from .formats import Kind, kind_for
from .pdfinspect import PdfType

log = logging.getLogger(__name__)

# Warn if less than this share of a digital PDF's text reaches the .docx.
MIN_TEXT_COVERAGE = 0.85
WORD = "Word"
# The losing engine's result is kept this long so users can still fetch it.
ALTERNATE_MAX_AGE_S = 7 * 24 * 3600


class PdfMode(str, Enum):
    BEST = "best"        # run both engines, keep the better result
    WORD = "word"        # Word's PDF reflow only
    PDF2DOCX = "pdf2docx"  # pdf2docx only (also works without Word)


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
    engine: str = ""             # engine whose output was kept
    scores: dict[str, scoring.Score] = field(default_factory=dict)
    alternate: Path | None = None        # the other engine's result, if any
    alternate_engine: str = ""

    @property
    def engine_summary(self) -> str:
        if not self.engine:
            return ""
        if len(self.scores) > 1:
            parts = [f"{n} {sc.total:.0%}" for n, sc in sorted(
                self.scores.items(), key=lambda kv: -kv[1].total)]
            return f"{self.engine} (" + " vs ".join(parts) + ")"
        return self.engine


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
    def __init__(
        self,
        office: OfficeLike | None,
        pdf_mode: PdfMode = PdfMode.BEST,
        alt_dir: Path | None = None,
        timeout_s: float = 300,
    ):
        self.office = office
        self.pdf_mode = PdfMode(pdf_mode)
        self.timeout_s = timeout_s
        self.alt_dir = alt_dir
        self._lock = threading.Lock()
        self._active: Pdf2DocxJob | None = None
        self._deadline = 0.0
        if alt_dir is not None:
            _prune(alt_dir, ALTERNATE_MAX_AGE_S)

    def kill(self) -> None:
        """Stop any background engine process (called on cancel/timeout)."""
        with self._lock:
            job = self._active
        if job:
            job.kill()

    def _need_office(self) -> OfficeLike:
        if self.office is None:
            raise ConversionError(
                "Microsoft Word is needed for this file type but isn't available on this PC."
            )
        return self.office

    def convert(self, src: Path, out_dir: Path | None = None, pdf_mode: PdfMode | None = None) -> Result:
        src = Path(src)
        if not src.is_file():
            raise ConversionError("The file no longer exists.")
        if out_dir is not None:
            out_dir.mkdir(parents=True, exist_ok=True)

        kind = kind_for(src)
        if kind == Kind.UNSUPPORTED:
            raise ConversionError(f"'{src.suffix}' files aren't supported.")

        mode = PdfMode(pdf_mode) if pdf_mode else self.pdf_mode
        self._deadline = time.monotonic() + self.timeout_s

        with TempDir() as tmp:
            # Work on a private copy with a short, plain name. This avoids
            # Windows path-length limits, odd characters, files locked by
            # another program and slow network drives.
            work_src = tmp / f"input{src.suffix.lower()}"
            shutil.copy2(src, work_src)
            work_out = tmp / "output.docx"

            result = self._dispatch(kind, work_src, work_out, tmp, mode)

            if not work_out.is_file() or work_out.stat().st_size == 0:
                raise ConversionError("The converter produced no output.")
            final = output_path_for(src, out_dir)
            shutil.move(str(work_out), final)
            result.output = final
            return result

    def _dispatch(self, kind: Kind, src: Path, out: Path, tmp: Path, mode: PdfMode) -> Result:
        if kind == Kind.PDF:
            return self._pdf(src, out, tmp, mode)
        if kind == Kind.WORD:
            self._need_office().word_to_docx(src, out)
            return Result(out, kind, "Document", engine=WORD)
        if kind == Kind.SPREADSHEET:
            pdf = tmp / "sheet.pdf"
            self._need_office().spreadsheet_to_pdf(src, pdf)
            res = self._pdf(pdf, out, tmp, mode)
            res.kind, res.detail = kind, "Spreadsheet"
            return res
        if kind == Kind.PRESENTATION:
            pdf = tmp / "slides.pdf"
            self._need_office().presentation_to_pdf(src, pdf)
            res = self._pdf(pdf, out, tmp, mode)
            res.kind, res.detail = kind, "Presentation"
            return res
        if kind == Kind.IMAGE:
            pages = images_to_docx(src, out)
            return Result(
                out, kind, f"Image, {pages} page{'s' if pages != 1 else ''}",
                ["Placed as a picture. Editable text from images comes with scan support (Phase 3)."],
            )
        raise ConversionError("Unsupported file type.")  # pragma: no cover

    def _engines_for(self, info: pdfinspect.PdfInfo, mode: PdfMode) -> tuple[bool, bool]:
        use_word = mode in (PdfMode.BEST, PdfMode.WORD)
        use_p2d = mode in (PdfMode.BEST, PdfMode.PDF2DOCX)
        if mode == PdfMode.BEST and info.doc_type == PdfType.SCANNED:
            use_p2d = False  # both just keep pictures of scans; don't waste time
        if self.office is None:
            if mode == PdfMode.WORD:
                self._need_office()  # raises with a clear message
            use_word, use_p2d = False, True
        return use_word, use_p2d

    def _pdf(self, src: Path, out: Path, tmp: Path, mode: PdfMode) -> Result:
        try:
            info = pdfinspect.inspect_pdf(src)
        except pdfinspect.PdfPasswordError as exc:
            raise ConversionError(str(exc)) from exc
        except Exception as exc:  # noqa: BLE001
            raise ConversionError("The PDF appears to be damaged.") from exc

        use_word, use_p2d = self._engines_for(info, mode)
        candidates: dict[str, Path] = {}
        errors: dict[str, ConversionError] = {}

        # Start pdf2docx in the background first, so it runs while Word works.
        job = Pdf2DocxJob(src, tmp / "pdf2docx.docx") if use_p2d else None
        with self._lock:
            self._active = job
        try:
            if use_word:
                try:
                    word_out = tmp / "word.docx"
                    self._need_office().word_to_docx(src, word_out)
                    candidates[WORD] = word_out
                except ConversionError as exc:
                    errors[WORD] = exc
            if job:
                try:
                    remaining = max(1.0, self._deadline - time.monotonic())
                    candidates[P2D] = job.wait(remaining)
                except ConversionError as exc:
                    errors[P2D] = exc
        finally:
            if job:
                job.kill()
            with self._lock:
                self._active = None

        if not candidates:
            # Prefer Word's error message - it's usually the more meaningful one.
            raise errors.get(WORD) or next(iter(errors.values()))
        for name, exc in errors.items():
            log.info("%s failed for %s, using the other engine: %s", name, src.name, exc)
        failed_note = ""
        if errors:
            failed = next(iter(errors))
            used = next(iter(candidates))
            failed_note = f"{failed} couldn't convert this file, so {used}'s version was used."

        result = Result(out, Kind.PDF, f"{info.doc_type.value} PDF")
        chosen = next(iter(candidates))
        if len(candidates) > 1:
            try:
                facts = scoring.pdf_facts(src)
                result.scores = {n: scoring.score_docx(facts, p) for n, p in candidates.items()}
                chosen = scoring.pick_best(result.scores, prefer=WORD)
                for n, sc in result.scores.items():
                    log.info("%s: %s scored %s", src.name, n, sc.describe())
            except Exception:  # noqa: BLE001 - scoring must never break a conversion
                log.exception("Scoring failed; keeping Word's result")
                chosen = WORD if WORD in candidates else chosen

        shutil.move(str(candidates.pop(chosen)), out)
        result.engine = chosen
        if candidates and self.alt_dir is not None:
            alt_name, alt_path = next(iter(candidates.items()))
            self.alt_dir.mkdir(parents=True, exist_ok=True)
            kept = self.alt_dir / f"{uuid.uuid4().hex}.docx"
            shutil.move(str(alt_path), kept)
            result.alternate, result.alternate_engine = kept, alt_name

        if failed_note:
            result.warnings.append(failed_note)
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
                if chosen in result.scores:
                    cov = result.scores[chosen].coverage
                else:
                    cov = quality.text_coverage(pdfinspect.extract_text(src), out)
                if cov < MIN_TEXT_COVERAGE:
                    result.warnings.append(
                        f"Only about {cov:.0%} of the PDF's text was found in the Word file - please check it."
                    )
            except Exception:  # noqa: BLE001 - a check, never a failure
                log.exception("Quality check failed for %s", src)
        return result


def _prune(folder: Path, max_age_s: float) -> None:
    """Delete stored alternate results older than ``max_age_s``."""
    if not folder.is_dir():
        return
    cutoff = time.time() - max_age_s
    for f in folder.glob("*.docx"):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
        except OSError:
            pass


def save_alternate(result: Result) -> Path:
    """Copy the other engine's version next to the chosen output.

    ``report.docx`` -> ``report (pdf2docx).docx``. Returns the new path.
    """
    if not result.alternate or not result.alternate.is_file():
        raise ConversionError("The other version is no longer available.")
    dest = unique_path(result.output.with_name(f"{result.output.stem} ({result.alternate_engine}).docx"))
    shutil.copy2(result.alternate, dest)
    return dest


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
