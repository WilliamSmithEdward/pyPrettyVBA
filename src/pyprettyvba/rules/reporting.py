"""Rules that report without fixing: untyped variables, long lines and
unreadable directives."""

from __future__ import annotations

from collections.abc import Iterable

from ..document import Document, LineKind, declaration_items
from ..lexer import TokenKind
from ..suppression import DIRECTIVE_CODE, scan_suppressions
from .base import Finding, Option, Rule
from .spacing import display_width

__all__ = ["DeclaredTypesRule", "MaxLineLengthRule", "SuppressionDirectiveRule"]

# The type characters that end a name: `s$`, `n&`, `f!`, `d#`, `c@`, `p^`.
# `$` and `@` lex as a type suffix, the others as an operator glued on.
_TYPE_CHARACTERS = frozenset(("&", "!", "#", "^"))


class DeclaredTypesRule(Rule):
    """Report a variable declared with no type.

    `Dim x` makes `x` a Variant, which is rarely what was meant. The rule
    says so and leaves the type to you: `As Variant` would only restate
    the default, and the real type is a decision. A name that ends in a
    type character (`s$`, `n&`) has a type. A `Const` takes its value's
    type and is not reported, and a `ReDim` resizes a variable declared
    elsewhere. Off unless enabled.
    """

    code = "declared-types"
    summary = "Report a variable declared with no type."
    category = "statements"
    default_enabled = False
    fixable = False

    def run(self, doc: Document) -> Iterable[Finding]:
        tokens = doc.tokens
        for line in doc.lines:
            if line.kind is not LineKind.CODE:
                continue
            for statement in line.statements:
                declaration = declaration_items(tokens, statement)
                if declaration is None:
                    continue
                keywords, items = declaration
                if any(tokens[j].lower == "const" for j in keywords):
                    continue
                for item in items:
                    name = next(
                        (j for j in item if tokens[j].kind in (TokenKind.IDENTIFIER, TokenKind.BRACKETED)), None
                    )
                    if name is None or _typed(doc, item, name):
                        continue
                    tok = tokens[name]
                    yield Finding(tok.start, tok.end, None, f"{tok.text} has no type; it is a Variant.")


def _typed(doc: Document, item: list[int], name: int) -> bool:
    """True when a declared item has an `As` clause or a type character on its name."""
    tokens = doc.tokens
    if any(tokens[j].kind is TokenKind.KEYWORD and tokens[j].lower == "as" for j in item):
        return True
    after = tokens[name + 1] if name + 1 < len(tokens) else None
    if after is None or after.start != tokens[name].end:
        return False
    return after.kind is TokenKind.TYPE_SUFFIX or (after.kind is TokenKind.OPERATOR and after.text in _TYPE_CHARACTERS)


class MaxLineLengthRule(Rule):
    """Report lines longer than a limit.

    Nothing is fixed: breaking a line needs a line continuation, and where to
    put one is a judgement this rule leaves to you.
    """

    code = "max-line-length"
    summary = "Report lines longer than a limit."
    category = "layout"
    default_enabled = False
    fixable = False
    options = (
        Option("max", 120, "The longest a line may be, in columns.", minimum=20),
        Option("ignore-comment-lines", False, "Do not report lines that hold only a comment."),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        limit = int(self.settings["max"])
        skip_comments = bool(self.settings["ignore-comment-lines"])
        tab = self.context.tab_width
        text = doc.text
        starts = doc.physical_starts
        kinds = {}
        for line in doc.lines:
            for physical in range(line.first_physical, line.last_physical + 1):
                kinds[physical] = line.kind
        for physical, start in enumerate(starts):
            end = starts[physical + 1] if physical + 1 < len(starts) else len(text)
            content = text[start:end].rstrip("\r\n")
            kind = kinds.get(physical)
            if kind is LineKind.ATTRIBUTE or (skip_comments and kind is LineKind.COMMENT):
                continue
            width = display_width(content, 0, tab)
            if width > limit:
                yield Finding(start, start + len(content), None, f"Line is {width} columns long; the limit is {limit}.")


class SuppressionDirectiveRule(Rule):
    """Report `'@prettyvba-ignore` directives that cannot be read.

    An unknown verb or rule name, an `-ignore-file` below the first line of
    code, or an `-end` without a `-start` suppresses nothing, which is easy to
    miss; this rule says so. No directive can suppress it.
    """

    code = DIRECTIVE_CODE
    summary = "Report suppression directives that cannot be read."
    category = "meta"
    fixable = False

    def run(self, doc: Document) -> Iterable[Finding]:
        index = scan_suppressions(doc, self.context.known_codes)
        for start, end, message in index.issues:
            yield Finding(start, end, None, message)
