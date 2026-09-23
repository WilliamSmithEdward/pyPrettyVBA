"""Rules that report without fixing: long lines and unreadable directives."""

from __future__ import annotations

from collections.abc import Iterable

from ..document import Document, LineKind
from ..suppression import DIRECTIVE_CODE, scan_suppressions
from .base import Finding, Option, Rule
from .spacing import display_width

__all__ = ["MaxLineLengthRule", "SuppressionDirectiveRule"]


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
