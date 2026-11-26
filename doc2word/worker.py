"""Background conversion queue.

One worker thread owns the Office session (COM objects must stay on the
thread that created them). A watchdog thread kills Office if a single file
takes too long, which un-sticks the worker so the queue carries on.

Plain threads + callbacks keep this usable without Qt; the GUI wraps the
callbacks in Qt signals.
"""

from __future__ import annotations

import itertools
import logging
import queue
import threading
import time
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable

from .converter import Converter, Result
from .engines.base import ConversionError

log = logging.getLogger(__name__)

DEFAULT_TIMEOUT_S = 300  # per file


class State(str, Enum):
    QUEUED = "Waiting"
    WORKING = "Converting…"
    DONE = "Done"
    WARNING = "Done - check"
    FAILED = "Failed"
    CANCELLED = "Cancelled"


@dataclass
class Job:
    id: int
    src: Path
    out_dir: Path | None
    state: State = State.QUEUED
    message: str = ""
    result: Result | None = None


Callback = Callable[[Job], None]
SessionFactory = Callable[[], object | None]


def default_session_factory():
    """Real Office on Windows; ``None`` elsewhere or if Word isn't installed."""
    from .engines import office

    if not office.office_available():
        log.warning("Microsoft Word not found - only images can be converted.")
        return None
    return office.OfficeSession()


class ConversionQueue:
    def __init__(
        self,
        on_update: Callback,
        session_factory: SessionFactory = default_session_factory,
        timeout_s: float = DEFAULT_TIMEOUT_S,
    ):
        self._on_update = on_update
        self._session_factory = session_factory
        self._timeout_s = timeout_s
        self._q: queue.Queue[Job | None] = queue.Queue()
        self._ids = itertools.count(1)
        self._cancelled: set[int] = set()
        self._lock = threading.Lock()
        self._current: Job | None = None
        self._started_at = 0.0
        self._killed = False
        self._session = None
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, name="doc2word-worker", daemon=True)
        self._watchdog = threading.Thread(target=self._watch, name="doc2word-watchdog", daemon=True)
        self._thread.start()
        self._watchdog.start()

    # -- public API (any thread) --------------------------------------------
    def submit(self, src: Path, out_dir: Path | None) -> Job:
        job = Job(next(self._ids), Path(src), out_dir)
        self._q.put(job)
        return job

    def cancel_all(self) -> None:
        """Cancel everything waiting, and stop the file currently converting."""
        while True:
            try:
                job = self._q.get_nowait()
            except queue.Empty:
                break
            if job is None:
                self._q.put(None)
                break
            job.state, job.message = State.CANCELLED, ""
            self._emit(job)
        with self._lock:
            current = self._current
            if current:
                self._cancelled.add(current.id)
        if current:
            self._kill_session()

    def shutdown(self, wait: float = 10.0) -> None:
        self._stop.set()
        self.cancel_all()
        self._q.put(None)
        self._thread.join(wait)

    @property
    def busy(self) -> bool:
        with self._lock:
            return self._current is not None or not self._q.empty()

    # -- internals ------------------------------------------------------------
    def _emit(self, job: Job) -> None:
        try:
            self._on_update(job)
        except Exception:  # noqa: BLE001
            log.exception("update callback failed")

    def _kill_session(self) -> None:
        session = self._session
        if session is not None and hasattr(session, "kill"):
            self._killed = True
            session.kill()

    def _watch(self) -> None:
        while not self._stop.wait(1.0):
            with self._lock:
                running = self._current is not None
                elapsed = time.monotonic() - self._started_at
            if running and elapsed > self._timeout_s and not self._killed:
                log.warning("Conversion exceeded %ss - stopping Office", self._timeout_s)
                self._kill_session()

    def _run(self) -> None:
        try:
            self._session = self._session_factory()
        except Exception:  # noqa: BLE001
            log.exception("Could not start Microsoft Office")
            self._session = None
        converter = Converter(self._session)

        try:
            while True:
                job = self._q.get()
                if job is None:
                    break
                with self._lock:
                    self._current = job
                    self._started_at = time.monotonic()
                    self._killed = False
                job.state = State.WORKING
                self._emit(job)
                try:
                    job.result = converter.convert(job.src, job.out_dir)
                    job.state = State.WARNING if job.result.warnings else State.DONE
                    job.message = " ".join(job.result.warnings)
                except Exception as exc:  # noqa: BLE001
                    self._handle_failure(job, exc)
                finally:
                    with self._lock:
                        self._current = None
                        self._cancelled.discard(job.id)
                self._emit(job)
        finally:
            if self._session is not None and hasattr(self._session, "close"):
                try:
                    self._session.close()
                except Exception:  # noqa: BLE001
                    log.exception("Error closing Office")

    def _handle_failure(self, job: Job, exc: Exception) -> None:
        with self._lock:
            cancelled = job.id in self._cancelled
        if cancelled:
            job.state, job.message = State.CANCELLED, ""
        elif self._killed:
            job.state = State.FAILED
            job.message = f"Took longer than {int(self._timeout_s)} seconds and was stopped."
        else:
            job.state = State.FAILED
            job.message = str(exc) if isinstance(exc, ConversionError) else f"Unexpected error: {exc}"
            if not isinstance(exc, ConversionError):
                log.exception("Unexpected failure converting %s", job.src)
        # After any failure Office may be hung, dead or holding a half-open
        # file. Restarting costs a few seconds and guarantees a clean state.
        if self._session is not None and hasattr(self._session, "reset"):
            try:
                self._session.reset()
            except Exception:  # noqa: BLE001
                pass
