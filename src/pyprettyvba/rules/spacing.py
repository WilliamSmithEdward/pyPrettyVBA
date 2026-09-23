"""Spacing inside a line: the gaps the VBE writes between tokens (see SpacingRule)."""

from __future__ import annotations

import bisect
from collections.abc import Iterable, Iterator

from ..document import Document, LineKind, LogicalLine, Statement, StatementKind
from ..keywords import (
    BINARY_OPERATOR_WORDS,
    FUNCTION_WORDS,
    OPERAND_WORDS,
    RESERVED_TYPE_IDENTIFIERS,
    UNARY_OPERATOR_WORDS,
)
from ..lexer import Token, TokenKind
from .base import Finding, Option, Rule

__all__ = ["SpacingRule", "LineAnalysis", "display_width", "is_name_like"]

# Gap policies.
NONE = "none"  # no whitespace
ONE = "one"  # exactly one space
KEEP = "keep"  # leave exactly as written
# One space, unless the gap was wider: then the token after it keeps its column.
KEEP_COLUMN = "keep-column"

# The VBE spells the reversed comparison operators the usual way (measured).
OPERATOR_SPELLING = {"=>": ">=", "=<": "<=", "><": "<>"}

_SUFFIX_LIKE = frozenset(("&", "!", "#", "^"))
TYPE_WORDS = frozenset(word.lower() for word in RESERVED_TYPE_IDENTIFIERS) | {"any"}


def is_name_like(tok: Token) -> bool:
    """A name, or a reserved word that is used as a value (Me, Debug, Date)."""
    if tok.kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED, TokenKind.TYPE_SUFFIX):
        return True
    return tok.kind is TokenKind.KEYWORD and tok.lower in OPERAND_WORDS


def takes_type_character(tok: Token) -> bool:
    """A token a glued `& ! # ^` can be the type character of.

    A name, but not one that already carries a `$ % @`: `s$&"x"` is a
    concatenation, the VBE writes `s$ & "x"`.
    """
    return tok.kind is not TokenKind.TYPE_SUFFIX and is_name_like(tok)


def display_width(text: str, column: int, tab_width: int) -> int:
    """Columns ``text`` spans when it starts at ``column``."""
    col = column
    for ch in text:
        if ch == "\t":
            col += tab_width - (col % tab_width)
        else:
            col += 1
    return col - column


def _is(tok: Token, kind: TokenKind, text: str) -> bool:
    return tok.kind is kind and tok.text == text


class LineAnalysis:
    """The facts about one logical line that decide its gaps."""

    def __init__(
        self, doc: Document, line: LogicalLine, *, type_member: bool = False, enum_member: bool = False
    ) -> None:
        self.member_line = type_member or enum_member
        self.doc = doc
        tokens = doc.tokens
        self.statement_of: dict[int, Statement] = {}
        members: list[int] = []
        for statement in line.statements:
            for j in statement.tokens:
                members.append(j)
                self.statement_of[j] = statement
        # Separating colons belong to no statement; merge them in.
        for j in range(line.first, line.stop):
            if tokens[j].kind is TokenKind.COLON and j != line.label_colon and j not in self.statement_of:
                members.append(j)
        self.order = sorted(members)
        # Keywords used as member names (`wb.Close(`, `.Print`) are names.
        self.member_keywords: set[int] = {
            j for j in self.order if tokens[j].kind is TokenKind.KEYWORD and doc.is_member_name(j)
        }
        self.unary: set[int] = set()
        self.call_paren: set[int] = set()
        self.call_comma: set[int] = set()
        self.suffix_ops: set[int] = set()
        self.keep_column: set[int] = set()
        for k, j in enumerate(self.order):
            tok = tokens[j]
            if tok.kind is TokenKind.OPERATOR and tok.text in _SUFFIX_LIKE and k > 0:
                prev_index = self.order[k - 1]
                if prev_index == j - 1 and takes_type_character(tokens[prev_index]):
                    self.suffix_ops.add(j)
        for statement in line.statements:
            self._statement(statement)
        if type_member and line.statements:
            self._mark_as_columns(line.statements[0].tokens)

    def _statement(self, statement: Statement) -> None:
        tokens = self.doc.tokens
        indices = statement.tokens
        # A Type or Enum member declares a name; it is not a call statement.
        if statement.kind is StatementKind.OTHER and not self.member_line:
            self._call_statement(indices)
        elif statement.kind is StatementKind.IF_SINGLE:
            # The statements after Then and Else are statements too:
            # `If x Then Foo (1)` (measured).
            for segment in _single_if_segments(tokens, indices):
                self._call_statement(segment)
        for k, j in enumerate(indices):
            tok = tokens[j]
            if tok.kind is TokenKind.OPERATOR and tok.text in ("-", "+") and j not in self.unary:
                if k == 0 or self._starts_operand(indices[k - 1]):
                    self.unary.add(j)
                elif statement.kind is StatementKind.FOR and tokens[indices[k - 1]].lower == "step":
                    # `Step` is a contextual keyword, lexed as a name.
                    self.unary.add(j)
        if statement.kind is StatementKind.VARIABLES and indices and tokens[indices[0]].lower != "redim":
            self._mark_as_columns(indices)

    def _call_statement(self, indices: list[int]) -> None:
        """Mark the gaps a call statement written without `Call` decides."""
        tokens = self.doc.tokens
        head_end = self._call_head_end(indices)
        if head_end is None or head_end + 1 >= len(indices):
            return
        nxt = indices[head_end + 1]
        ntok = tokens[nxt]
        if ntok.kind is TokenKind.OPERATOR and ntok.text in ("-", "+"):
            # `Foo -1`: the sign opens the first argument.
            self.unary.add(nxt)
        elif _is(ntok, TokenKind.PUNCTUATION, ","):
            # `Foo , 2`: an omitted first argument. The VBE writes the space
            # even where it was left out, and reads both the same (measured).
            self.call_comma.add(nxt)
        elif _is(ntok, TokenKind.PUNCTUATION, "("):
            close = _matching_paren(tokens, indices, head_end + 1)
            after = tokens[indices[close + 1]] if close is not None and close + 1 < len(indices) else None
            continues = after is not None and (
                (after.kind is TokenKind.OPERATOR and after.text in ("=", "!"))
                or (after.kind is TokenKind.PUNCTUATION and after.text in (".", "("))
            )
            if close is not None and not continues:
                # The parenthesis is the first argument, not an index.
                self.call_paren.add(nxt)

    def _mark_as_columns(self, indices: list[int]) -> None:
        """Keep the column of each `As` at parenthesis depth 0."""
        tokens = self.doc.tokens
        depth = 0
        for j in indices:
            tok = tokens[j]
            if tok.kind is TokenKind.PUNCTUATION and tok.text == "(":
                depth += 1
            elif tok.kind is TokenKind.PUNCTUATION and tok.text == ")":
                depth -= 1
            elif depth == 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "as":
                self.keep_column.add(j)

    def _starts_operand(self, prev_index: int) -> bool:
        """True when a sign after the token at ``prev_index`` is unary."""
        prev = self.doc.tokens[prev_index]
        if prev.kind is TokenKind.OPERATOR:
            return prev_index not in self.suffix_ops
        if prev.kind is TokenKind.PUNCTUATION:
            return prev.text in ("(", ",", ";")
        if prev.kind in (TokenKind.COLON, TokenKind.DIRECTIVE):
            return True
        if prev.kind is TokenKind.KEYWORD:
            return prev.lower not in OPERAND_WORDS
        return False

    def _call_head_end(self, indices: list[int]) -> int | None:
        """Last position of the name chain a call statement opens with.

        `Foo`, `obj.Method`, `.Member`, `Me!Ctl.Value`: names joined by `.`
        or `!` written without spaces.
        """
        tokens = self.doc.tokens
        n = len(indices)
        if n == 0:
            return None
        k = 0
        if _is(tokens[indices[0]], TokenKind.PUNCTUATION, "."):
            k = 1
        elif not is_name_like(tokens[indices[0]]):
            return None
        end: int | None = None
        while k < n:
            tok = tokens[indices[k]]
            member = k > 0 and indices[k] == indices[k - 1] + 1 and _is_member_joint(tokens[indices[k - 1]])
            if not (is_name_like(tok) or (member and tok.kind is TokenKind.KEYWORD)):
                break
            end = k
            k += 1
            if k < n and tokens[indices[k]].kind is TokenKind.TYPE_SUFFIX:
                end = k
                k += 1
            if k < n and indices[k] == indices[k - 1] + 1 and _is_member_joint(tokens[indices[k]]):
                k += 1
                continue
            break
        return end


def _single_if_segments(tokens: list[Token], indices: list[int]) -> list[list[int]]:
    """The statements a single-line If runs: after Then, after Else, after `:`."""
    segments: list[list[int]] = []
    current: list[int] = []
    depth = 0
    started = False
    for j in indices:
        tok = tokens[j]
        if tok.kind is TokenKind.PUNCTUATION and tok.text in "()":
            depth += 1 if tok.text == "(" else -1
        if depth == 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "then" and not started:
            started = True
            continue
        if not started:
            continue
        if depth == 0 and (tok.kind is TokenKind.COLON or (tok.kind is TokenKind.KEYWORD and tok.lower == "else")):
            segments.append(current)
            current = []
            continue
        current.append(j)
    segments.append(current)
    return [segment for segment in segments if segment]


def _is_member_joint(tok: Token) -> bool:
    return _is(tok, TokenKind.PUNCTUATION, ".") or _is(tok, TokenKind.OPERATOR, "!")


def _matching_paren(tokens: list[Token], indices: list[int], k: int) -> int | None:
    depth = 0
    for m in range(k, len(indices)):
        tok = tokens[indices[m]]
        if tok.kind is TokenKind.PUNCTUATION:
            if tok.text == "(":
                depth += 1
            elif tok.text == ")":
                depth -= 1
                if depth == 0:
                    return m
    return None


def gap_policy(ana: LineAnalysis, a_index: int, b_index: int) -> str:
    """How much whitespace belongs between tokens a and b."""
    tokens = ana.doc.tokens
    a = tokens[a_index]
    b = tokens[b_index]
    glued = b_index == a_index + 1

    # Whitespace that decides the parse is never touched.
    if a.kind is TokenKind.UNKNOWN or b.kind is TokenKind.UNKNOWN:
        return KEEP
    if b.kind is TokenKind.TYPE_SUFFIX or b_index in ana.suffix_ops:
        return KEEP
    if a_index in ana.suffix_ops and not (
        b.kind in (TokenKind.OPERATOR, TokenKind.COLON)
        or (b.kind is TokenKind.PUNCTUATION and b.text in (",", ";", ")"))
    ):
        # `rs!Field`, `s&t`: glued on both sides, whatever it means.
        return KEEP
    if _is(a, TokenKind.OPERATOR, "!") and a_index not in ana.suffix_ops:
        return KEEP
    if _is(b, TokenKind.OPERATOR, "!"):
        return KEEP
    if _is(a, TokenKind.PUNCTUATION, "."):
        return KEEP
    if _is(b, TokenKind.PUNCTUATION, ".") and (
        is_name_like(a) or a_index in ana.member_keywords or _is(a, TokenKind.PUNCTUATION, ")")
    ):
        # `Foo.Bar` stays glued; `Foo .Bar` (a With member argument) keeps
        # its one space.
        return NONE if glued else ONE
    if a.kind is TokenKind.DIRECTIVE:
        return NONE if glued else KEEP
    if _is(a, TokenKind.OPERATOR, "#") and a_index not in ana.suffix_ops:
        # The file number marker: `#1`.
        return KEEP

    # Separators. The statement after a `:` keeps its column when the gap
    # was widened, as `As` does (measured).
    if a.kind is TokenKind.COLON:
        return KEEP_COLUMN
    if b.kind is TokenKind.COLON:
        return NONE
    if b_index in ana.call_comma:
        return ONE
    if a.kind is TokenKind.PUNCTUATION and a.text in (",", ";"):
        return NONE if _is(b, TokenKind.PUNCTUATION, ")") else ONE
    if b.kind is TokenKind.PUNCTUATION and b.text in (",", ";", ")"):
        return NONE
    if _is(a, TokenKind.PUNCTUATION, "("):
        return NONE
    if _is(a, TokenKind.OPERATOR, ":=") or _is(b, TokenKind.OPERATOR, ":="):
        return NONE

    statement = ana.statement_of.get(b_index) or ana.statement_of.get(a_index)
    if statement is not None and statement.kind is StatementKind.DEFTYPE and "-" in (a.text, b.text):
        # DefInt A-Z: letter ranges stay tight.
        return NONE

    if b_index in ana.keep_column:
        return KEEP_COLUMN
    if _is(b, TokenKind.PUNCTUATION, "("):
        return _before_paren(ana, a, a_index, b_index)
    if _is(b, TokenKind.PUNCTUATION, "."):
        # A With member after an operator, keyword or separator.
        return ONE if a.kind is not TokenKind.PUNCTUATION else NONE
    if _is(a, TokenKind.OPERATOR, "-") or _is(a, TokenKind.OPERATOR, "+"):
        return NONE if a_index in ana.unary else ONE
    return ONE


def _before_paren(ana: LineAnalysis, a: Token, a_index: int, b_index: int) -> str:
    if b_index in ana.call_paren:
        return ONE
    if a.kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED, TokenKind.TYPE_SUFFIX) or a_index in ana.member_keywords:
        return NONE
    if _is(a, TokenKind.PUNCTUATION, ")"):
        return NONE
    if a.kind is TokenKind.KEYWORD:
        if a.lower in FUNCTION_WORDS or a.lower in TYPE_WORDS or a.lower == "me":
            # `Len(x)`, and `As Byte()`: an array of a built-in type.
            return NONE
        if a.lower in UNARY_OPERATOR_WORDS or a.lower in BINARY_OPERATOR_WORDS:
            return ONE
        return ONE
    if a.kind is TokenKind.OPERATOR and a.text in ("-", "+") and a_index in ana.unary:
        return NONE
    if a_index in ana.suffix_ops:
        return NONE
    return ONE


class SpacingRule(Rule):
    """Insert the spaces the VBE writes between tokens and collapse the rest.

    What the VBE does to a line it parses (tests/oracle/vbe_rendering.json):

    * One space around `=`, the comparison operators, and the binary
      operators (`+ - * / \\ ^ & Mod And Or Xor Eqv Imp Like Is`); none
      after a unary sign (`x = -y`, `Step -1`).
    * One space after `,` and `;`, none before; none inside parentheses.
      A call statement whose first argument is left out keeps a space
      before the comma: the VBE writes `Foo , 2`.
    * No space before the parenthesis of a call or an index inside an
      expression (`x = Foo(a)`), but one where a statement calls a
      procedure without `Call` and the parenthesis is its first argument:
      the VBE writes `Foo (x)` and `Debug.Print (x)`. After a keyword it is
      one space: `Not (x)`, `If (a) Then`.
    * No space around `:=`, `.`, `!`, and a type character glued to its
      name.
    * After a `:` separator the gap is kept, and at least one space.
    * A run of spaces collapses to one, except indentation, the gap before
      an end-of-line comment, and the column of `As` in a Dim, Private,
      Public, Global, Static or Const declaration or a Type member. The VBE
      keeps that column, which is how aligned declarations survive. Tabs
      become spaces.
    * `=>`, `=<` and `><` are written `>=`, `<=` and `<>`.

    The VBE stores a line it cannot parse as written. The rule is careful
    where whitespace decides the parse: it never adds or removes the space
    that tells a type character from an operator (`s&t` against `s & t`),
    a member from a With member (`Foo.Bar` against `Foo .Bar`), or a bang
    from a syntax error (`rs!Field` against `rs ! Field`).

    With `collapse = false` the rule only inserts missing spaces and never
    removes one, as XLIDE's Format Document does.
    """

    code = "spacing"
    summary = "Space tokens the way the VBE writes them."
    category = "spacing"
    vbe_canonical = True
    options = (
        Option(
            "collapse",
            True,
            "Collapse runs of spaces the way the VBE does. When false, only insert "
            "missing spaces and never remove one, as XLIDE's Format Document does.",
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        collapse = bool(self.settings["collapse"])
        block: StatementKind | None = None  # TYPE_START or ENUM_START while inside one
        for line in doc.lines:
            if line.kind is LineKind.CODE and line.statements:
                first = line.statements[0].kind
                if first in (StatementKind.TYPE_START, StatementKind.ENUM_START):
                    block = first
                elif first in (StatementKind.TYPE_END, StatementKind.ENUM_END):
                    block = None
            if line.kind not in (LineKind.CODE, LineKind.DIRECTIVE):
                continue
            member = (
                block is not None
                and line.kind is LineKind.CODE
                and bool(line.statements)
                and line.statements[0].kind is StatementKind.OTHER
            )
            ana = LineAnalysis(
                doc,
                line,
                type_member=member and block is StatementKind.TYPE_START,
                enum_member=member and block is StatementKind.ENUM_START,
            )
            yield from self._line(doc, ana, collapse)
            if collapse:
                for j in ana.order:
                    tok = doc.tokens[j]
                    if tok.kind is TokenKind.OPERATOR and tok.text in OPERATOR_SPELLING:
                        wanted = OPERATOR_SPELLING[tok.text]
                        yield Finding(tok.start, tok.end, wanted, f"The VBE writes {tok.text} as {wanted}.")

    def _line(self, doc: Document, ana: LineAnalysis, collapse: bool) -> Iterator[Finding]:
        tokens = doc.tokens
        order = ana.order
        if len(order) < 2:
            return
        tab = self.context.tab_width
        first = tokens[order[0]]
        line_start = doc.physical_starts[bisect.bisect_right(doc.physical_starts, first.start) - 1]
        old_col = display_width(doc.text[line_start : first.start], 0, tab)
        new_col = old_col
        # Columns just after token a, in the old text and the new.
        old_after = old_col + display_width(first.text, old_col, tab)
        new_after = new_col + display_width(first.text, new_col, tab)
        for k in range(1, len(order)):
            a_index, b_index = order[k - 1], order[k]
            b = tokens[b_index]
            gap_tokens = tokens[a_index + 1 : b_index]
            continuation = next(
                (t for t in gap_tokens if t.kind is TokenKind.CONTINUATION), None
            )
            if continuation is not None:
                before = gap_tokens[0] if gap_tokens[0].kind is TokenKind.WHITESPACE else None
                if collapse and before is not None and before.text != " ":
                    yield Finding(before.start, before.end, " ", "Use one space before a line continuation.")
                # A new physical line: indentation is not ours to change.
                col = display_width(doc.text[continuation.end : b.start], 0, tab)
                old_after = col + display_width(b.text, col, tab)
                new_after = old_after
                continue
            start, end = tokens[a_index].end, b.start
            current = doc.text[start:end]
            old_b = old_after + display_width(current, old_after, tab)
            policy = gap_policy(ana, a_index, b_index)
            width = old_b - old_after
            if not collapse:
                # Insert-only: add a missing space, never remove one.
                wanted = " " if policy in (ONE, KEEP_COLUMN) and current == "" else current
            elif policy == KEEP:
                wanted = current
            elif policy == NONE:
                wanted = ""
            elif policy == ONE or width <= 1:
                # The VBE keeps a wider gap only where the user widened it: a
                # single space before `As` stays a single space even when the
                # text before it shrinks (measured).
                wanted = " "
            else:
                # KEEP_COLUMN: the token after the gap stays in its column.
                wanted = " " * max(1, old_b - new_after)
            if wanted != current:
                yield Finding(start, end, wanted, _message(tokens[a_index], b, wanted))
            new_b = new_after + display_width(wanted, new_after, tab)
            old_after = old_b + display_width(b.text, old_b, tab)
            new_after = new_b + display_width(b.text, new_b, tab)


def _message(a: Token, b: Token, wanted: str) -> str:
    if wanted == "":
        return f"Remove the space between {a.text!r} and {b.text!r}."
    if wanted == " ":
        return f"Use one space between {a.text!r} and {b.text!r}."
    return f"Keep {b.text!r} in its column."
