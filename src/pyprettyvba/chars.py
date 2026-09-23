"""Character classes shared by the lexer and the literal parser."""

from __future__ import annotations

import re
import unicodedata

__all__ = ["WSC_CHARS", "WSC_CLASS", "is_wsc", "is_ident_start", "is_ident_part"]

# WSC (MS-VBAL 3.2.2): tab, the eom character, space, the DBCS space, and the
# Unicode Zs spaces. U+2028 and U+2029 count as whitespace rather than line
# terminators, as they do in the XLIDE lexer; the VBE reads source in an ANSI
# code page where neither can occur. Spelled as code points to keep the
# source ASCII.
WSC_CHARS: frozenset[str] = frozenset(
    {
        chr(0x09),
        chr(0x19),
        chr(0x20),
        chr(0xA0),
        chr(0x1680),
        chr(0x2028),
        chr(0x2029),
        chr(0x202F),
        chr(0x205F),
        chr(0x3000),
        chr(0xFEFF),
    }
    | {chr(c) for c in range(0x2000, 0x200B)}
)

# A regex character class matching one WSC character.
WSC_CLASS = "[" + re.escape("".join(sorted(WSC_CHARS))) + "]"


def is_wsc(ch: str) -> bool:
    """True when ``ch`` is VBA whitespace (never a line terminator)."""
    return ch in WSC_CHARS


def is_ident_start(ch: str) -> bool:
    """True when ``ch`` can begin a VBA name."""
    if "A" <= ch <= "Z" or "a" <= ch <= "z":
        return True
    return ord(ch) >= 0x80 and ch.isalpha()


def is_ident_part(ch: str) -> bool:
    """True when ``ch`` can continue a VBA name.

    A combining mark continues the name it is attached to (Thai and
    Devanagari build one letter from a base and a mark); the VBE accepts such
    names.
    """
    if ch == "_" or "0" <= ch <= "9" or is_ident_start(ch):
        return True
    return ord(ch) >= 0x80 and unicodedata.category(ch).startswith("M")
