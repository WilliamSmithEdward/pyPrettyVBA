"""Wrap long lines with line continuations.

A line longer than the limit is broken after a comma, a spaced binary
operator or the `Then` of a single-line If, at the rightmost such point
that leaves the line, with its ` _`, within the limit, and again on what
is left. The continuation lines sit one level in from the line. The
break never falls inside a string or a comment, and never at a `.`:
`Foo _` on one line and `.Bar` on the next passes a With member where
`Foo.Bar` called one. A line with no such point within the limit is left
as written, for max-line-length to report. A line already continued by
hand is left alone while each of its physical lines fits; one that does
not is laid out again from the start.
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass

from ..document import Document, LineKind, LogicalLine, StatementKind
from ..lexer import Token, TokenKind
from .base import Finding, Option, Rule
from .spacing import display_width

__all__ = ["WrapLinesRule"]

# Operators a line may break after, when they are spaced on both sides:
# glued, `&` is a type character and `-` may be a sign.
_OPERATORS = frozenset(("&", "+", "-", "*", "/"))
_OPERATOR_WORDS = frozenset(("and", "or", "xor"))


@dataclass(frozen=True)
class _Piece:
    """One token of the line as laid out again, and whether a break may follow it."""

    text: str
    space: bool  # whitespace, dropped at the start of a continuation line
    breakable: bool


class WrapLinesRule(Rule):
    """Wrap lines longer than a limit with line continuations.

    The line is broken after a comma, a spaced binary operator (`&`, `+`,
    `-`, `*`, `/`, `And`, `Or`, `Xor`) or the `Then` of a single-line If,
    at the rightmost point that keeps the line, with its ` _`, within
    `max`, and again on what is left; the continuation lines sit one
    level in. A break never falls inside a string or a comment, and never
    at a `.`, which would change what the line means. A trailing comment
    rides on the last line. A line with no break point within the limit
    is left as written, for max-line-length to report, and a line already
    continued by hand is left alone while each of its lines fits. A style
    the VBE does not impose; off unless enabled.
    """

    code = "wrap-lines"
    summary = "Wrap lines longer than a limit."
    category = "layout"
    default_enabled = False
    options = (Option("max", 120, "The longest a line may be, in columns.", minimum=20),)

    def run(self, doc: Document) -> Iterable[Finding]:
        limit = int(self.settings["max"])
        tab = self.context.tab_width
        for line in doc.lines:
            if line.kind is not LineKind.CODE or line.label is not None or not line.statements:
                continue
            if not self._too_long(doc, line, limit):
                continue
            tokens = doc.tokens[line.first : line.stop]
            if any(t.kind is TokenKind.UNKNOWN for t in tokens) or any(
                t.kind is TokenKind.COMMENT and ("\n" in t.text or "\r" in t.text) for t in tokens
            ):
                continue
            indent = tokens[0].text if tokens[0].kind is TokenKind.WHITESPACE else ""
            pieces = _pieces(doc, line)
            if not pieces:
                continue
            newline = tokens[-1].text if tokens[-1].kind is TokenKind.NEWLINE else self.context.newline
            continuation = self.context.indent_text(display_width(indent, 0, tab) + self.context.indent_width)
            laid_out = _lay_out(pieces, indent, continuation, limit, tab)
            if laid_out is None:
                continue
            text = newline.join(laid_out)
            end = line.end - len(newline) if tokens[-1].kind is TokenKind.NEWLINE else line.end
            if text != doc.text[line.start : end]:
                yield Finding(line.start, end, text, f"Wrap the line within {limit} columns.")

    @staticmethod
    def _too_long(doc: Document, line: LogicalLine, limit: int) -> bool:
        """True when any physical line of ``line`` is over the limit."""
        starts = doc.physical_starts
        text = doc.text
        for physical in range(line.first_physical, line.last_physical + 1):
            start = starts[physical]
            end = starts[physical + 1] if physical + 1 < len(starts) else len(text)
            if display_width(text[start:end].rstrip("\r\n"), 0, 4) > limit:
                return True
        return False


def _pieces(doc: Document, line: LogicalLine) -> list[_Piece]:
    """The line's tokens after its indentation, with each continuation and
    the whitespace around it collapsed to one space."""
    tokens = doc.tokens
    then_at = _then_of_single_if(doc, line)
    pieces: list[_Piece] = []
    k = line.first
    if k < line.stop and tokens[k].kind is TokenKind.WHITESPACE:
        k += 1
    pending_space = False
    for j in range(k, line.stop):
        tok = tokens[j]
        if tok.kind is TokenKind.NEWLINE:
            break
        if tok.kind is TokenKind.WHITESPACE:
            if pending_space:
                continue
            pieces.append(_Piece(tok.text, True, False))
            continue
        if tok.kind is TokenKind.CONTINUATION:
            if pieces and pieces[-1].space:
                pieces[-1] = _Piece(" ", True, False)
            elif pieces:
                pieces.append(_Piece(" ", True, False))
            pending_space = True
            continue
        pending_space = False
        pieces.append(_Piece(tok.text, False, _breakable(tokens, j, then_at)))
    while pieces and pieces[-1].space:
        pieces.pop()
    return pieces


def _then_of_single_if(doc: Document, line: LogicalLine) -> int | None:
    statement = line.statements[0]
    if statement.kind is not StatementKind.IF_SINGLE:
        return None
    depth = 0
    for j in statement.tokens:
        tok = doc.tokens[j]
        if tok.kind is TokenKind.PUNCTUATION and tok.text in ("(", ")"):
            depth += 1 if tok.text == "(" else -1
        elif depth == 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "then":
            return j
    return None


def _breakable(tokens: list[Token], j: int, then_at: int | None) -> bool:
    tok = tokens[j]
    if j == then_at:
        return True
    if tok.kind is TokenKind.PUNCTUATION and tok.text == ",":
        return True
    spaced = (
        j > 0
        and j + 1 < len(tokens)
        and tokens[j - 1].kind is TokenKind.WHITESPACE
        and tokens[j + 1].kind is TokenKind.WHITESPACE
    )
    if tok.kind is TokenKind.OPERATOR and tok.text in _OPERATORS:
        return spaced
    return tok.kind is TokenKind.KEYWORD and tok.lower in _OPERATOR_WORDS


def _lay_out(pieces: list[_Piece], indent: str, continuation: str, limit: int, tab: int) -> list[str] | None:
    """The physical lines of the wrapped line, or None when it cannot be wrapped."""
    lines: list[str] = []
    prefix = indent
    rest = pieces
    broke = False
    while rest:
        taken: list[_Piece] = []
        split_at: int | None = None
        for piece in rest:
            taken.append(piece)
            width = display_width(prefix + "".join(p.text for p in taken), 0, tab)
            if width <= limit:
                continue
            split_at = _last_break(prefix, taken, limit, tab)
            if split_at is None:
                break
            lines.append(prefix + "".join(p.text for p in taken[: split_at + 1]).rstrip() + " _")
            rest = rest[split_at + 1 :]
            while rest and rest[0].space:
                rest = rest[1:]
            prefix = continuation
            broke = True
            break
        else:
            lines.append(prefix + "".join(p.text for p in taken))
            rest = []
        if split_at is None and rest:
            # Nothing on this line can take a break: keep what there is.
            lines.append(prefix + "".join(p.text for p in rest))
            rest = []
    return lines if broke else None


def _last_break(prefix: str, taken: list[_Piece], limit: int, tab: int) -> int | None:
    """The rightmost piece of ``taken`` a break may follow, with the line
    up to it and its ` _` within the limit, and something left to carry."""
    for index in range(len(taken) - 1, -1, -1):
        piece = taken[index]
        if not piece.breakable:
            continue
        if all(p.space or p.text.startswith("'") for p in taken[index + 1 :]) and index == len(taken) - 1:
            continue
        carried = [p for p in taken[index + 1 :] if not p.space]
        if carried and all(p.text.startswith("'") or p.text[:3].lower() == "rem" for p in carried):
            continue
        head = prefix + "".join(p.text for p in taken[: index + 1]).rstrip() + " _"
        if display_width(head, 0, tab) <= limit:
            return index
    return None
