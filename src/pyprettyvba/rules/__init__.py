"""The rule catalogue, in the order the pipeline runs it.

The order is part of the design: each rule runs on the output of the rules
before it, and is placed so that no later rule undoes it. Token rewrites
come first, then statement splitting, then spacing inside lines, then
indentation (which splitting needs), then the placement that depends on
columns (alignment, end-of-line comments), then whole-line hygiene. Rules
that only report run last, on the formatted text.
"""

from __future__ import annotations

from .alignment import AlignDeclarationsRule
from .base import Finding, FormatContext, Option, Rule
from .casing import IdentifierCaseRule, KeywordCaseRule
from .comments import CommentSpaceRule, RemCommentsRule, TrailingCommentsRule
from .layout import (
    BlankLinesRule,
    ContinuationIndentRule,
    EndOfFileRule,
    IndentRule,
    LineEndingsRule,
    TrailingWhitespaceRule,
)
from .literals import DateLiteralRule, NumericLiteralRule
from .reporting import MaxLineLengthRule, SuppressionDirectiveRule
from .spacing import SpacingRule
from .statements import LetKeywordRule, SplitStatementsRule, StatementFormRule

__all__ = ["RULES", "RULES_BY_CODE", "Finding", "FormatContext", "Option", "Rule"]

RULES: tuple[type[Rule], ...] = (
    LineEndingsRule,
    RemCommentsRule,
    CommentSpaceRule,
    LetKeywordRule,
    KeywordCaseRule,
    IdentifierCaseRule,
    NumericLiteralRule,
    DateLiteralRule,
    StatementFormRule,
    SplitStatementsRule,
    SpacingRule,
    IndentRule,
    ContinuationIndentRule,
    AlignDeclarationsRule,
    TrailingCommentsRule,
    TrailingWhitespaceRule,
    BlankLinesRule,
    EndOfFileRule,
    MaxLineLengthRule,
    SuppressionDirectiveRule,
)

RULES_BY_CODE: dict[str, type[Rule]] = {rule.code: rule for rule in RULES}
