"""Numeric and date literals, spelled the way the VBE spells them.

The VBE rewrites a literal when it reads a line: `1.0` becomes `1#`, `.5`
becomes `0.5`, `&hff` becomes `&HFF`, `#2020-01-15#` becomes `#1/15/2020#`
(tests/oracle/vbe_rendering.json). These rules apply the same spellings, but
only where the result means exactly the same value and type.

Where the VBE's spelling would change the value, numeric-literals leaves
the literal alone and reports it instead: the VBE keeps 15 significant
digits of a Double, so it stores `3.141592653589793` as `3.14159265358979`,
a different number, and the next export writes that.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator

from ..document import Document, LineKind
from ..lexer import TokenKind, tokenize
from ..literals import canonical_date, canonical_number, precision_loss
from .base import Finding, Option, Rule

__all__ = ["DateLiteralRule", "NumericLiteralRule"]


class NumericLiteralRule(Rule):
    """Spell numeric literals the way the VBE does, when the value is kept exactly.

    The VBE rewrites a number when it reads a line (measured): `1.0` becomes
    `1#`, `.5` becomes `0.5`, `&hff` becomes `&HFF`, `&17` becomes `&O17`,
    and a type suffix the value makes redundant goes (`10%` is `10`, but
    `10&` keeps its `&`, since `10` alone would be an Integer). This rule
    writes the same spellings, only where the value and type stay the same.

    The VBE keeps 15 significant digits of a Double and 7 of a Single, so it
    stores `3.141592653589793` as `3.14159265358979`, a different number.
    Such a literal is left alone and reported. A number that opens a line is
    a line number, and keeps its spelling.
    """

    code = "numeric-literals"
    summary = "Spell numbers the way the VBE does."
    category = "literals"
    vbe_canonical = True
    options = (
        Option(
            "report-precision-loss",
            True,
            "Report a literal with more digits than the VBE keeps, which the VBE "
            "silently replaces with a different value.",
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        report = bool(self.settings["report-precision-loss"])
        # A number that opens a line is where a line number goes: `00010`
        # is a label, not a literal, and `10%` spelled `10` would become one.
        heads = {
            doc.tokens[line.head].start
            for line in doc.lines
            if line.kind is LineKind.CODE and line.head is not None
        }
        for j in _literal_tokens(doc, (TokenKind.INTEGER, TokenKind.FLOAT)):
            tok = doc.tokens[j]
            if tok.start in heads:
                continue
            canonical = canonical_number(tok.text)
            if canonical is not None and canonical != tok.text:
                if _stays_one_token(doc, j, canonical):
                    yield Finding(tok.start, tok.end, canonical, f"The VBE writes {tok.text} as {canonical}.")
            elif canonical is None and report:
                lost = precision_loss(tok.text)
                if lost is not None:
                    yield Finding(
                        tok.start,
                        tok.end,
                        None,
                        f"The VBE keeps fewer digits than {tok.text} has and stores {lost}, "
                        "a different value.",
                    )


class DateLiteralRule(Rule):
    """Spell date literals the way the VBE does: #M/D/YYYY h:mm:ss AM#.

    `#2020-01-15#` becomes `#1/15/2020#`, `#13:30#` becomes `#1:30:00 PM#`,
    and a midnight time is dropped from a date (measured). A literal whose
    value depends on the machine, a two-digit year or a date without a
    year, is left as written.
    """

    code = "date-literals"
    summary = "Spell dates the way the VBE does."
    category = "literals"
    vbe_canonical = True

    def run(self, doc: Document) -> Iterable[Finding]:
        for j in _literal_tokens(doc, (TokenKind.DATE,)):
            tok = doc.tokens[j]
            canonical = canonical_date(tok.text)
            if canonical is not None and canonical != tok.text:
                yield Finding(tok.start, tok.end, canonical, f"The VBE writes {tok.text} as {canonical}.")


def _literal_tokens(doc: Document, kinds: tuple[TokenKind, ...]) -> Iterator[int]:
    for line in doc.lines:
        if line.kind not in (LineKind.CODE, LineKind.DIRECTIVE):
            continue
        for j in range(line.first, line.stop):
            if doc.tokens[j].kind in kinds:
                yield j


_GLUE_RISK = frozenset((
    TokenKind.INTEGER, TokenKind.FLOAT, TokenKind.DATE, TokenKind.IDENTIFIER, TokenKind.KEYWORD,
    TokenKind.BRACKETED, TokenKind.TYPE_SUFFIX, TokenKind.UNKNOWN,
))


def _stays_one_token(doc: Document, j: int, replacement: str) -> bool:
    """True when token j respelled still lexes apart from its neighbours.

    `10%0` is two numbers; without its redundant `%`, `100` is one.
    """
    tokens = doc.tokens
    for k in (j - 1, j + 1):
        # A neighbour glued on that may be respelled in the same pass
        # (`&hff%.5`: `&HFF` and `0.5` together read `&HFF0.5`).
        if 0 <= k < len(tokens) and tokens[k].kind in _GLUE_RISK:
            return False
    before = tokens[j - 1].text if j > 0 else ""
    after = tokens[j + 1].text if j + 1 < len(tokens) else ""
    relexed = [tok.text for tok in tokenize(before + replacement + after)]
    return relexed == [text for text in (before, replacement, after) if text]
