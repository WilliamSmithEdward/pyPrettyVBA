"""Comments: where an end-of-line comment sits, and how comments are written.

The VBE keeps an end-of-line comment in its column when the gap before it
is wider than one space, and otherwise writes exactly one space before it:
`x = 1'note` becomes `x = 1 'note`, while `x = 1      ' note` keeps its
alignment (tests/oracle/vbe_rendering.json). trailing-comments does that by
default, measuring the column in the source as written, so a comment keeps
its place even when indentation moves the code in front of it.

comment-space and rem-comments are style choices the VBE does not make, and
are off unless enabled.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator

from ..document import Document, LineKind, LogicalLine
from ..lexer import Token, TokenKind
from .base import Finding, Option, Rule
from .spacing import display_width

__all__ = ["CommentSpaceRule", "RemCommentsRule", "TrailingCommentsRule"]


def _trailing_comments(doc: Document) -> Iterator[tuple[LogicalLine, Token, Token]]:
    """(line, the token before the comment, the comment) for each end-of-line comment."""
    for line in doc.lines:
        if line.comment is None or line.kind not in (LineKind.CODE, LineKind.DIRECTIVE):
            continue
        comment = doc.tokens[line.comment]
        before = doc.prev_code(line.comment)
        if before is None:
            continue
        between = doc.tokens[before + 1 : line.comment]
        if any(t.kind is TokenKind.CONTINUATION for t in between):
            continue
        yield line, doc.tokens[before], comment


class TrailingCommentsRule(Rule):
    """Place end-of-line comments: in their column, one space after the code, or aligned.

    The VBE keeps an end-of-line comment in its column when the gap before
    it is wider than one space, and otherwise writes exactly one space:
    `x = 1'note` becomes `x = 1 'note`, while `x = 1      ' note` keeps its
    alignment (measured). The default, `keep-column`, does the same,
    measuring the column in the source as written, so a comment stays put
    when indentation moves the code in front of it. A comment one space after
    its code stays one space after it.
    """

    code = "trailing-comments"
    summary = "Place end-of-line comments consistently."
    category = "spacing"
    vbe_canonical = True
    options = (
        Option(
            "position",
            "keep-column",
            "`keep-column` keeps a comment in the column it was written at when it "
            "was set apart by more than one space, as the VBE does; `one-space` puts "
            "every comment one space after its code; `align` lines up the comments "
            "of consecutive lines.",
            choices=("keep-column", "one-space", "align"),
        ),
        Option("min-gap", 1, "Fewest spaces between code and its comment.", minimum=1, maximum=40),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        position = self.settings["position"]
        if position == "align":
            yield from self._align(doc)
            return
        tab = self.context.tab_width
        min_gap = int(self.settings["min-gap"])
        origin = self.context.origin
        for _line, before, comment in _trailing_comments(doc):
            gap = doc.text[before.end : comment.start]
            code_end = self._column(doc, before.end)
            if position == "one-space":
                wanted = " " * min_gap
            else:
                original_col = origin.column(comment.start, tab)
                original_text = origin.text
                original_start = origin.offset(comment.start)
                line_start = max(original_text.rfind("\n", 0, original_start), original_text.rfind("\r", 0, original_start)) + 1
                code_text = original_text[line_start:original_start].rstrip(" \t")
                original_code_end = display_width(code_text, 0, tab)
                if original_col - original_code_end <= 1:
                    wanted = " " * min_gap
                else:
                    wanted = " " * max(min_gap, original_col - code_end)
            if wanted != gap:
                yield Finding(before.end, comment.start, wanted, _message(wanted, code_end))

    def _align(self, doc: Document) -> Iterator[Finding]:
        min_gap = int(self.settings["min-gap"])
        # Comments on consecutive lines form a group aligned to one column.
        group: list[tuple[Token, Token, int]] = []
        last_line = -2
        for line, before, comment in _trailing_comments(doc):
            if line.index != last_line + 1:
                yield from _aligned(doc, group, min_gap)
                group = []
            group.append((before, comment, self._column(doc, before.end)))
            last_line = line.index
        yield from _aligned(doc, group, min_gap)

    def _column(self, doc: Document, offset: int) -> int:
        start = doc.physical_starts[doc.physical_line_of(offset)]
        return display_width(doc.text[start:offset], 0, self.context.tab_width)


def _message(wanted: str, code_end: int) -> str:
    if len(wanted) == 1:
        return "Put one space before the comment."
    return f"Keep the comment at column {code_end + len(wanted) + 1}."


def _aligned(doc: Document, group: list[tuple[Token, Token, int]], min_gap: int) -> Iterator[Finding]:
    """Align a group of comments (token before, comment, code end column)."""
    if len(group) > 1:
        target = max(end for _b, _c, end in group) + min_gap
        for before, comment, end in group:
            wanted = " " * (target - end)
            if doc.text[before.end : comment.start] != wanted:
                yield Finding(before.end, comment.start, wanted, f"Align the comment at column {target + 1}.")
    elif group:
        before, comment, _end = group[0]
        if doc.text[before.end : comment.start] == "":
            yield Finding(before.end, comment.start, " " * min_gap, "Put a space before the comment.")


class CommentSpaceRule(Rule):
    """Put a space between the apostrophe and the text of a comment.

    `'note` becomes `' note`. Documentation comments (`'''`), annotations and
    directives (`'@...`), and comments that open with a character listed in
    ``ignore`` (separator lines like `'-----`) are left as they are.
    """

    code = "comment-space"
    summary = "Put a space after the apostrophe of a comment."
    category = "comments"
    default_enabled = False
    options = (
        Option(
            "ignore",
            ["'", "@", "!", "#", "-", "=", "*", "~", "/", "<", ">", "|", "+", "_", "."],
            "Comments whose text starts with one of these are left alone.",
            types=(list,),
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        ignore = tuple(str(p) for p in self.settings["ignore"])
        for tok in doc.tokens:
            if tok.kind is not TokenKind.COMMENT or not tok.text.startswith("'"):
                continue
            body = tok.text[1:]
            if not body or body[0] in (" ", "\t") or body.startswith(ignore):
                continue
            yield Finding(tok.start + 1, tok.start + 1, " ", "Put a space after the apostrophe.")


class RemCommentsRule(Rule):
    """Write `Rem` comments with an apostrophe: `Rem note` becomes `' note`.

    Both mark the rest of the line as a comment, and both continue onto the
    next line after a trailing ` _` (measured), so the change is only one of
    style. A style the VBE does not impose; off unless enabled.
    """

    code = "rem-comments"
    summary = "Write Rem comments with an apostrophe."
    category = "comments"
    default_enabled = False

    def run(self, doc: Document) -> Iterable[Finding]:
        for tok in doc.tokens:
            if tok.kind is TokenKind.COMMENT and tok.text[:3].lower() == "rem":
                yield Finding(tok.start, tok.start + 3, "'", "Write the comment with an apostrophe.")
