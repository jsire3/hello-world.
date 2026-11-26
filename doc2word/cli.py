"""Command-line converter, for batch jobs and testing.

    doc2word-cli report.pdf scans\\ --out C:\\Converted
"""

from __future__ import annotations

import argparse
import sys
import threading
from pathlib import Path

from . import __version__
from .formats import expand_inputs
from .converter import PdfMode
from .logsetup import alternates_dir, setup_logging
from .worker import ConversionQueue, Job, State

FINISHED = {State.DONE, State.WARNING, State.FAILED, State.CANCELLED}


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="doc2word-cli", description="Convert files to Word (.docx).")
    ap.add_argument("inputs", nargs="+", type=Path, help="files or folders")
    ap.add_argument("--out", type=Path, default=None, help="output folder (default: next to each file)")
    ap.add_argument(
        "--pdf-engine", choices=[m.value for m in PdfMode], default=PdfMode.BEST.value,
        help="best = run Word and pdf2docx and keep the better result (default)",
    )
    ap.add_argument("--timeout", type=float, default=300, help="seconds allowed per file")
    ap.add_argument("-v", "--verbose", action="store_true")
    ap.add_argument("--version", action="version", version=__version__)
    args = ap.parse_args(argv)

    setup_logging(args.verbose)
    files = expand_inputs(args.inputs)
    if not files:
        print("No supported files found.", file=sys.stderr)
        return 2

    remaining = len(files)
    failed = 0
    done = threading.Event()
    lock = threading.Lock()

    def on_update(job: Job) -> None:
        nonlocal remaining, failed
        if job.state not in FINISHED:
            return
        line = f"[{job.state.value}] {job.src}"
        if job.result:
            line += f" -> {job.result.output}"
            if job.result.engine_summary:
                line += f"  [{job.result.engine_summary}]"
        if job.message:
            line += f"\n    {job.message}"
        print(line, flush=True)
        with lock:
            remaining -= 1
            failed += job.state in (State.FAILED, State.CANCELLED)
            if remaining == 0:
                done.set()

    q = ConversionQueue(on_update, timeout_s=args.timeout, alt_dir=alternates_dir())
    for f in files:
        q.submit(f, args.out, PdfMode(args.pdf_engine))
    try:
        done.wait()
    except KeyboardInterrupt:
        q.cancel_all()
    q.shutdown()
    print(f"\n{len(files) - failed} of {len(files)} converted.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
