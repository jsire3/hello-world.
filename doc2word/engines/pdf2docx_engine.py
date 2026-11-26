"""pdf2docx engine, run in a separate process.

Why a separate process:
* It can run **at the same time** as Word, so "best of both" costs about the
  time of the slower engine, not the sum of both.
* A pathological PDF can make it run for a very long time; a process can be
  terminated cleanly, a thread can't.
* A crash in native code (PyMuPDF/OpenCV) can't take the app down with it.

The child writes the .docx to ``dest`` and, on failure, a one-line reason to
``dest.with_suffix('.err')``. Exit code 0 means success.
"""

from __future__ import annotations

import logging
import multiprocessing as mp
import time
from pathlib import Path

from .base import ConversionError

ENGINE_NAME = "pdf2docx"


def _child(src: str, dest: str) -> None:
    # Runs in the child process.
    import os
    import sys
    import warnings

    warnings.filterwarnings("ignore")
    logging.disable(logging.CRITICAL)  # pdf2docx logs every page at INFO
    try:  # lower priority so the PC stays responsive
        import psutil

        p = psutil.Process()
        p.nice(psutil.BELOW_NORMAL_PRIORITY_CLASS if sys.platform == "win32" else 10)
    except Exception:  # noqa: BLE001
        pass
    try:
        from pdf2docx import Converter

        cv = Converter(src)
        try:
            cv.convert(dest, multi_processing=False)
        finally:
            cv.close()
    except BaseException as exc:  # noqa: BLE001
        try:
            Path(dest).with_suffix(".err").write_text(f"{type(exc).__name__}: {exc}"[:500], encoding="utf-8")
        finally:
            os._exit(1)
    os._exit(0)


class Pdf2DocxJob:
    """One pdf2docx conversion running in the background."""

    def __init__(self, src: Path, dest: Path):
        self.dest = dest
        ctx = mp.get_context("spawn")  # Windows only supports spawn anyway
        self.proc = ctx.Process(target=_child, args=(str(src), str(dest)), daemon=True,
                                name="doc2word-pdf2docx")
        self.proc.start()
        self.started = time.monotonic()

    def wait(self, timeout: float | None) -> Path:
        """Block until done; return the .docx or raise ConversionError."""
        self.proc.join(timeout)
        if self.proc.is_alive():
            self.kill()
            raise ConversionError(f"{ENGINE_NAME} took too long and was stopped.")
        if self.proc.exitcode == 0 and self.dest.is_file() and self.dest.stat().st_size > 0:
            return self.dest
        err = self.dest.with_suffix(".err")
        reason = err.read_text(encoding="utf-8") if err.is_file() else f"exit code {self.proc.exitcode}"
        raise ConversionError(f"{ENGINE_NAME} could not convert this file ({reason}).")

    def kill(self) -> None:
        if self.proc.is_alive():
            self.proc.kill()
            self.proc.join(5)
