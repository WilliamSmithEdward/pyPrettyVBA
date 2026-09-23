"""Reading and writing module files without disturbing their bytes.

The VBE exports modules in the machine's ANSI code page (Windows-1252 on
Western systems), with CRLF line endings; files that went through other
tools may be UTF-8, with or without a byte order mark. ``decode`` reads
either: strict UTF-8 when the bytes are valid UTF-8, otherwise the ANSI code
page. Decoding uses ``surrogateescape``, so a byte the code page does not
define survives the round trip, and ``encode`` writes the text back in the
encoding it was read with, BOM included. Formatting changes only ASCII
characters, so a file's other bytes come out exactly as they went in.
"""

from __future__ import annotations

import codecs
import locale
import sys
from dataclasses import dataclass

__all__ = ["DecodedText", "ansi_code_page", "decode", "encode"]

_UTF8_BOM = codecs.BOM_UTF8


@dataclass(frozen=True)
class DecodedText:
    text: str
    encoding: str
    bom: bool


def ansi_code_page() -> str:
    """The code page the VBE on this machine exports with.

    On Windows, the ANSI code page; elsewhere cp1252, which is what the
    great majority of exported modules use.
    """
    if sys.platform == "win32":
        encoding = locale.getencoding()
        if encoding and encoding.lower().replace("-", "") not in ("utf8", "cp65001"):
            return encoding
    return "cp1252"


def decode(data: bytes, encoding: str = "auto") -> DecodedText:
    """Decode module bytes (see the module docstring)."""
    if encoding != "auto":
        name = codecs.lookup(encoding).name
        if name == "utf-8" and data.startswith(_UTF8_BOM):
            return DecodedText(data[len(_UTF8_BOM):].decode("utf-8", "surrogateescape"), "utf-8", True)
        return DecodedText(data.decode(name, "surrogateescape"), name, False)
    if data.startswith(_UTF8_BOM):
        return DecodedText(data[len(_UTF8_BOM):].decode("utf-8", "surrogateescape"), "utf-8", True)
    try:
        return DecodedText(data.decode("utf-8"), "utf-8", False)
    except UnicodeDecodeError:
        page = ansi_code_page()
        return DecodedText(data.decode(page, "surrogateescape"), page, False)


def encode(text: str, encoding: str, bom: bool = False) -> bytes:
    """Encode text the way ``decode`` read it."""
    data = text.encode(encoding, "surrogateescape")
    return _UTF8_BOM + data if bom else data
