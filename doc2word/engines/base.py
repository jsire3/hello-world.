from __future__ import annotations

from pathlib import Path
from typing import Protocol


class ConversionError(Exception):
    """A conversion failed with a message that is safe to show the user."""


class Engine(Protocol):
    """Something that can turn a file into a .docx at ``dest``."""

    def to_docx(self, src: Path, dest: Path) -> None: ...

    def close(self) -> None: ...
