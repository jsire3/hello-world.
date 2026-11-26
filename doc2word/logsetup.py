from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import APP_NAME


def data_dir() -> Path:
    """Per-user app data: %LOCALAPPDATA%\\Doc2Word on Windows."""
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state"))
    return base / APP_NAME


def log_dir() -> Path:
    return data_dir() / "logs"


def alternates_dir() -> Path:
    """Where the losing engine's result is kept for a week."""
    return data_dir() / "alternates"


def setup_logging(verbose: bool = False) -> None:
    folder = log_dir()
    handlers: list[logging.Handler] = []
    try:
        folder.mkdir(parents=True, exist_ok=True)
        handlers.append(
            RotatingFileHandler(folder / "doc2word.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8")
        )
    except OSError:
        pass
    if sys.stderr is not None:  # None in a windowed (no console) build
        handlers.append(logging.StreamHandler())
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=handlers,
    )
