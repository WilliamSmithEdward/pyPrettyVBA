"""Statement-level style: one statement per line, and the optional Let.

Both are off unless enabled; neither is something the VBE does.

split-statements never splits a single-line If. Everything after its Then
belongs to it, colons included: `If a Then b: c` runs `c` only when `a`
holds (measured), so putting `c` on a line of its own would run it always.
The document model gives a single-line If as one statement, so there is no
colon inside it to split at.
"""

from __future__ import annotations

from collections.abc import Iterable

from ..document import Document, LineKind, LogicalLine, Statement, StatementKind
from ..lexer import TokenKind
from .base import Finding, Option, Rule

__all__ = ["LetKeywordRule", "SplitStatementsRule", "StatementFormRule"]


class SplitStatementsRule(Rule):
    """Put each statement on a line of its own.

    `x = 1: y = 2` becomes two lines. A single-line If stays whole, a label
    keeps the statement that follows it unless ``labels`` is `split`, and
    `Case 1: x = 1` is split unless ``case`` is `keep`. A colon that ends a
    line with nothing after it is removed.
    """

    code = "split-statements"
    summary = "Put each statement on a line of its own."
    category = "statements"
    default_enabled = False
    options = (
        Option(
            "case",
            "split",
            "`split` moves the statements after `Case x:` to lines of their own; "
            "`keep` leaves a one-line Case arm as it is.",
            choices=("split", "keep"),
        ),
        Option(
            "labels",
            "keep",
            "`keep` leaves `Label: statement` on one line; `split` moves the "
            "statement below the label.",
            choices=("keep", "split"),
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        tokens = doc.tokens
        for line in doc.lines:
            if line.kind is not LineKind.CODE:
                continue
            if any(tokens[j].kind is TokenKind.UNKNOWN for j in range(line.first, line.stop)):
                # Text the lexer has no rule for: its neighbours stay as written.
                continue
            newline = self._newline(doc, line)
            indent = doc.indent_of(line)
            statements = line.statements
            if (
                self.settings["case"] == "keep"
                and statements
                and statements[0].kind is StatementKind.CASE
            ):
                continue
            if (
                self.settings["labels"] == "split"
                and line.label is not None
                and statements
                and not _reads_differently_at_line_start(doc, statements[0])
            ):
                after_label = line.label_colon if line.label_colon is not None else line.label
                first = statements[0].tokens[0]
                yield Finding(
                    tokens[after_label].end,
                    tokens[first].start,
                    newline + indent,
                    "Put the statement after the label on its own line.",
                )
            for index, statement in enumerate(statements):
                if not statement.colon_after:
                    continue
                last = statement.tokens[-1]
                colon = doc.next_code(last)
                if colon is None or tokens[colon].kind is not TokenKind.COLON:
                    continue
                if index + 1 < len(statements):
                    if _reads_differently_at_line_start(doc, statements[index + 1]):
                        continue
                    nxt = statements[index + 1].tokens[0]
                    yield Finding(
                        tokens[last].end,
                        tokens[nxt].start,
                        newline + indent,
                        "Put each statement on a line of its own.",
                    )
                elif line.comment is not None and tokens[line.comment].text[:3].lower() == "rem":
                    # `x = 1: Rem note`: without the colon, Rem is not a comment.
                    continue
                else:
                    # Colons that end the line (a comment may follow them),
                    # all of them at once: `x = 1::` has two.
                    end = colon
                    after = doc.next_code(end)
                    while after is not None and tokens[after].kind is TokenKind.COLON:
                        end = after
                        after = doc.next_code(end)
                    yield Finding(tokens[last].end, tokens[end].end, "", "Remove the colon at the end of the line.")

    def _newline(self, doc: Document, line: LogicalLine) -> str:
        tokens = doc.tokens
        if line.stop > line.first and tokens[line.stop - 1].kind is TokenKind.NEWLINE:
            return tokens[line.stop - 1].text
        return self.context.newline


def _reads_differently_at_line_start(doc: Document, statement: Statement) -> bool:
    """True when a statement moved to the start of a line would read as something else.

    `x = 1: Foo: Bar` calls Foo; on a line of its own, `Foo:` defines a
    label. A number that starts a line is a line number, a `#` a directive,
    and `Attribute` an attribute line.
    """
    first = doc.tokens[statement.tokens[0]]
    if first.kind is TokenKind.INTEGER and first.text.isdigit():
        return True
    if first.text.startswith("#") or first.lower == "attribute":
        return True
    return len(statement.tokens) == 1 and first.kind is TokenKind.IDENTIFIER and statement.colon_after


class StatementFormRule(Rule):
    """Write two statements the way the VBE rewrites them.

    `Call Foo()` becomes `Call Foo`: the VBE drops the empty parentheses. A
    block `Else` followed by a statement on the same line gets a colon:
    `Else x = 1` becomes `Else: x = 1`. Both are measured
    (tests/oracle/vbe_rendering.json) and neither changes what runs.
    """

    code = "statement-form"
    summary = "Write Call Foo() and Else x the way the VBE does."
    category = "statements"
    vbe_canonical = True

    def run(self, doc: Document) -> Iterable[Finding]:
        tokens = doc.tokens
        for line in doc.lines:
            if line.kind is not LineKind.CODE:
                continue
            for index, statement in enumerate(line.statements):
                indices = statement.tokens
                first = tokens[indices[0]]
                if statement.kind is StatementKind.ELSE and not statement.colon_after:
                    if index + 1 < len(line.statements):
                        nxt = tokens[line.statements[index + 1].tokens[0]]
                        # Only between Else and a statement it is apart from:
                        # a colon glued to `Else=` or `Else$` makes new tokens.
                        if nxt.start > first.end and (
                            nxt.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD, TokenKind.BRACKETED)
                            or (nxt.kind is TokenKind.PUNCTUATION and nxt.text == ".")
                        ):
                            yield Finding(first.end, first.end, ":", "The VBE writes a colon after Else here.")
                    continue
                if (
                    first.kind is TokenKind.KEYWORD
                    and first.lower == "call"
                    and len(indices) >= 4
                    and tokens[indices[-2]].text == "("
                    and tokens[indices[-1]].text == ")"
                    and _call_target_ends_before(doc, indices, len(indices) - 2)
                ):
                    yield Finding(
                        tokens[indices[-3]].end,
                        tokens[indices[-1]].end,
                        "",
                        "The VBE drops the empty parentheses after Call.",
                    )


def _call_target_ends_before(doc: Document, indices: list[int], position: int) -> bool:
    """True when indices[1:position] is a plain name chain (`Foo`, `obj.Method`)."""
    tokens = doc.tokens
    expect_name = True
    for j in indices[1:position]:
        tok = tokens[j]
        if expect_name:
            if tok.kind not in (TokenKind.IDENTIFIER, TokenKind.KEYWORD, TokenKind.BRACKETED):
                return False
            expect_name = False
        elif tok.kind is TokenKind.TYPE_SUFFIX:
            continue
        elif (tok.kind is TokenKind.PUNCTUATION and tok.text == ".") or (
            tok.kind is TokenKind.OPERATOR and tok.text == "!"
        ):
            expect_name = True
        else:
            return False
    return not expect_name


class LetKeywordRule(Rule):
    """Drop the optional `Let` from assignments: `Let x = 1` becomes `x = 1`.

    Only a statement that is plainly an assignment loses its `Let`: a name
    after it and an `=` outside parentheses. A style the VBE does not
    impose; off unless enabled.
    """

    code = "let-keyword"
    summary = "Drop the optional Let from assignments."
    category = "statements"
    default_enabled = False

    def run(self, doc: Document) -> Iterable[Finding]:
        tokens = doc.tokens
        for line in doc.lines:
            if line.kind is not LineKind.CODE:
                continue
            for statement in line.statements:
                if len(statement.tokens) < 2:
                    continue
                first = tokens[statement.tokens[0]]
                if first.kind is TokenKind.KEYWORD and first.lower == "let" and _is_assignment(doc, statement):
                    nxt = tokens[statement.tokens[1]]
                    yield Finding(first.start, nxt.start, "", "Drop the optional Let.")


def _is_assignment(doc: Document, statement: Statement) -> bool:
    """True for `Let target = value`: a name after Let, and `=` outside parentheses.

    Anything else is left alone; `Let 10` without its Let would be a line number.
    """
    tokens = doc.tokens
    target = tokens[statement.tokens[1]]
    if target.kind not in (TokenKind.IDENTIFIER, TokenKind.BRACKETED) and target.lower != "me":
        return False
    depth = 0
    for j in statement.tokens[2:]:
        tok = tokens[j]
        if tok.kind is TokenKind.PUNCTUATION and tok.text in ("(", ")"):
            depth += 1 if tok.text == "(" else -1
        elif depth == 0 and tok.kind is TokenKind.OPERATOR and tok.text == "=":
            return True
    return False
