"""Where a contextual keyword is a keyword.

`Explicit`, `Lib`, `Step` and their kind are ordinary names to the lexer:
VBA reserves none of them. The VBE still capitalizes each one inside the
statement that gives it meaning (`Option Explicit`, `Declare ... Lib`,
`For ... Step`) and leaves the same word alone anywhere else, so
`ws.text` stays `ws.text` while `Option Compare text` becomes
`Option Compare Text` (tests/oracle/vbe_rendering.json).

``contextual_keywords`` finds those positions. keyword-case capitalizes
them; identifier-case leaves them to keyword-case.
"""

from __future__ import annotations

from .document import Document, LineKind, Statement, StatementKind
from .keywords import CONTEXTUAL_KEYWORDS
from .lexer import TokenKind

__all__ = ["contextual_keywords"]

_OPTION_WORDS = frozenset(("explicit", "base", "compare", "binary", "text", "database", "module"))
_OPEN_WORDS = frozenset(("append", "binary", "output", "random", "access", "read"))


def contextual_keywords(doc: Document) -> dict[int, str]:
    """Token index -> the VBE's spelling, for each contextual keyword."""
    found: dict[int, str] = {}
    for line in doc.lines:
        if line.kind is not LineKind.CODE:
            continue
        for statement in line.statements:
            _statement(doc, statement, found)
    return found


def _statement(doc: Document, statement: Statement, found: dict[int, str]) -> None:
    tokens = doc.tokens
    indices = statement.tokens
    if not indices:
        return

    def mark(j: int) -> None:
        tok = tokens[j]
        if tok.kind is TokenKind.IDENTIFIER and tok.lower in CONTEXTUAL_KEYWORDS:
            found[j] = CONTEXTUAL_KEYWORDS[tok.lower]

    def word(k: int) -> str:
        if 0 <= k < len(indices) and tokens[indices[k]].is_word:
            return tokens[indices[k]].lower
        return ""

    # `As Object` anywhere: Object is a type name, not a reserved word.
    for k in range(1, len(indices)):
        if tokens[indices[k]].lower == "object" and tokens[indices[k - 1]].lower in ("as", "new"):
            mark(indices[k])

    kind = statement.kind
    first = word(0)
    if kind is StatementKind.OPTION:
        for j in indices[1:]:
            if tokens[j].lower in _OPTION_WORDS:
                mark(j)
        return
    if kind is StatementKind.DECLARE:
        for j in indices:
            tok = tokens[j]
            if tok.kind is TokenKind.PUNCTUATION and tok.text == "(":
                break
            if tok.lower in ("ptrsafe", "lib", "alias"):
                mark(j)
        return
    if kind in (StatementKind.PROC_START, StatementKind.PROC_END):
        for j in indices:
            if tokens[j].lower == "property":
                mark(j)
                break
        return
    if kind is StatementKind.FOR:
        depth = 0
        for j in indices:
            tok = tokens[j]
            if tok.kind is TokenKind.PUNCTUATION and tok.text in "()":
                depth += 1 if tok.text == "(" else -1
            elif depth == 0 and tok.lower == "step":
                mark(j)
        return
    if kind is StatementKind.DEFTYPE:
        return
    if first == "exit" and word(1) == "property":
        mark(indices[1])
    elif first == "on" and word(1) == "error":
        mark(indices[1])
    elif first == "error" and len(indices) > 1 and tokens[indices[1]].kind is not TokenKind.PUNCTUATION:
        # The Error statement (`Error 5`), not the Error function.
        mark(indices[0])
    elif first == "open":
        for j in indices[1:]:
            if tokens[j].lower in _OPEN_WORDS:
                mark(j)
    elif first == "line" and word(1) == "input":
        mark(indices[0])
    elif first == "width" and len(indices) > 1 and tokens[indices[1]].text == "#":
        mark(indices[0])
    elif first == "name" and len(indices) > 1 and tokens[indices[1]].kind in (
        TokenKind.IDENTIFIER, TokenKind.STRING, TokenKind.KEYWORD,
    ):
        if any(tokens[j].lower == "as" for j in indices[2:]):
            mark(indices[0])
