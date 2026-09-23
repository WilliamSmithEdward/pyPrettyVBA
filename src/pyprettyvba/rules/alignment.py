"""Aligning the `As` of consecutive declarations.

The VBE keeps the column of `As` in a declaration when the gap before it is
wider than one space (tests/oracle/vbe_rendering.json), so an aligned block
survives an import and export. It does not keep aligned `=` in an Enum or a
Const without `As`; those are collapsed, so this rule does not align them.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator

from ..document import Document, LineKind, LogicalLine, StatementKind
from ..lexer import Token, TokenKind
from .base import Finding, Option, Rule
from .spacing import display_width

__all__ = ["AlignDeclarationsRule"]


class AlignDeclarationsRule(Rule):
    """Line up `As` in runs of single-variable declarations.

    A run is consecutive lines, at one indentation, that each declare one
    variable or constant with an `As` clause (`Dim`, `Private`, `Public`,
    `Global`, `Static`, `Const`), or consecutive members of a Type. A blank
    line, a comment line, or any other statement ends the run.
    """

    code = "align-declarations"
    summary = "Align As in runs of declarations."
    category = "spacing"
    default_enabled = False
    options = (
        Option("min-gap", 1, "Fewest spaces before the aligned As.", minimum=1, maximum=40),
        Option(
            "max-column",
            60,
            "Leave a line out of the alignment when its name reaches past this column.",
            minimum=10,
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        run: list[tuple[Token, Token, int, int]] = []
        indent: str | None = None
        last_index = -2
        in_type = False
        for line in doc.lines:
            if line.kind is LineKind.CODE and line.statements:
                kind = line.statements[0].kind
                if kind is StatementKind.TYPE_START:
                    in_type = True
                elif kind is StatementKind.TYPE_END:
                    in_type = False
            item = self._candidate(doc, line, in_type)
            this_indent = doc.indent_of(line) if item is not None else None
            if item is not None and line.index == last_index + 1 and this_indent == indent:
                run.append(item)
            else:
                yield from self._flush(doc, run)
                run = [item] if item is not None else []
                indent = this_indent
            last_index = line.index if item is not None else -2
        yield from self._flush(doc, run)

    def _candidate(self, doc: Document, line: LogicalLine, in_type: bool) -> tuple[Token, Token, int, int] | None:
        """(token before As, the As, name end column, As column) for a line that can align."""
        if line.kind is not LineKind.CODE or len(line.statements) != 1 or line.label is not None:
            return None
        statement = line.statements[0]
        tokens = doc.tokens
        if statement.kind is StatementKind.VARIABLES:
            if tokens[statement.tokens[0]].lower == "redim":
                return None
        elif not (in_type and statement.kind is StatementKind.OTHER):
            return None
        depth = 0
        as_at: int | None = None
        for position, j in enumerate(statement.tokens):
            tok = tokens[j]
            if tok.kind is TokenKind.PUNCTUATION:
                if tok.text == "(":
                    depth += 1
                elif tok.text == ")":
                    depth -= 1
                elif tok.text == "," and depth == 0:
                    return None  # more than one variable
            elif depth == 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "as" and as_at is None:
                as_at = position
        if as_at is None or as_at == 0:
            return None
        before = tokens[statement.tokens[as_at - 1]]
        as_token = tokens[statement.tokens[as_at]]
        gap = doc.text[before.end : as_token.start]
        if "\n" in gap or "\r" in gap:
            return None
        start = doc.physical_starts[doc.physical_line_of(before.end)]
        tab = self.context.tab_width
        end_col = display_width(doc.text[start : before.end], 0, tab)
        if end_col >= int(self.settings["max-column"]):
            return None
        return before, as_token, end_col, end_col + display_width(gap, end_col, tab)

    def _flush(self, doc: Document, run: list[tuple[Token, Token, int, int]]) -> Iterator[Finding]:
        if len(run) < 2:
            return
        target = max(end for _b, _a, end, _col in run) + int(self.settings["min-gap"])
        for before, as_token, end, _col in run:
            wanted = " " * (target - end)
            if doc.text[before.end : as_token.start] != wanted:
                yield Finding(before.end, as_token.start, wanted, f"Align As at column {target + 1}.")
