"""Conversion engines. Each engine turns one input file into a .docx."""

from .base import ConversionError, Engine

__all__ = ["ConversionError", "Engine"]
