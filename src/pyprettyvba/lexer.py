"""A lossless VBA lexer.

Every character of the input belongs to exactly one token, so joining the
token texts reproduces the source exactly. Whitespace, line continuations and
line breaks are tokens of their own; the formatter works by changing them.

The grammar is MS-VBAL v20250520 section 3, read the way the VBE reads it.
tests/oracle/vbe_rendering.json records the evidence for each point below:

* A comment runs to the end of its logical line. A comment that ends in
  ` _` continues onto the next physical line, for apostrophe and `Rem`
  comments alike (MS-VBAL 3.3.1 comment-body). The VBE agrees: a statement
  on the line after `' note _` does not run.
* A line continuation may carry whitespace after the underscore. The spec
  does not allow it, but the VBE trims trailing whitespace from every line it
  stores, so ` _   ` continues the line exactly as ` _` does.
* `=>`, `=<` and `><` lex as one operator. They are two tokens in the
  spec's syntactic grammar; reading them as one keeps the formatter from
  ever separating them.

The lexer never fails. Text it has no rule for becomes an UNKNOWN token, and
the formatter leaves the gaps around such a token alone.
"""

from __future__ import annotations

import enum
import re
from dataclasses import dataclass

from .chars import WSC_CHARS, WSC_CLASS, is_ident_part, is_ident_start
from .keywords import canonical_keyword
from .literals import parse_date_body

__all__ = ["Token", "TokenKind", "name_key", "tokenize"]

_ASCII_LOWER = str.maketrans("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz")


def name_key(name: str) -> str:
    """A name with its Latin letters in lower case, and nothing else changed.

    MS-VBAL 3.3.5 makes names equal when they differ only in the case of
    Latin letters; it defines no case equivalence for the code-page letters
    names may also contain (Cyrillic, Greek, ...). So only A-Z fold, and a
    name in another script is never respelled.
    """
    return name.translate(_ASCII_LOWER)


class TokenKind(enum.StrEnum):
    """The category of a token."""

    NEWLINE = "newline"
    WHITESPACE = "whitespace"
    # An underscore, any whitespace after it, and the line terminator. The
    # whitespace before the underscore is a WHITESPACE token of its own.
    CONTINUATION = "continuation"
    COMMENT = "comment"
    KEYWORD = "keyword"
    IDENTIFIER = "identifier"
    BRACKETED = "bracketed"
    # A `$`, `%` or `@` glued to a name. The other type characters (`&`, `!`,
    # `#`, `^`) double as operators, so they lex as OPERATOR, and the
    # formatter leaves a glued one exactly where it is.
    TYPE_SUFFIX = "type-suffix"
    INTEGER = "integer"
    FLOAT = "float"
    DATE = "date"
    STRING = "string"
    OPERATOR = "operator"
    PUNCTUATION = "punctuation"
    COLON = "colon"
    # The `#` that opens a conditional-compilation directive line.
    DIRECTIVE = "directive"
    UNKNOWN = "unknown"


@dataclass(frozen=True, slots=True)
class Token:
    """One token: its kind, exact text, and where it sits in the source."""

    kind: TokenKind
    text: str
    start: int
    end: int
    # 0-based physical line of the first character.
    line: int

    @property
    def lower(self) -> str:
        return self.text.lower()

    @property
    def key(self) -> str:
        """The name's identity: VBA folds the case of Latin letters only."""
        return name_key(self.text)

    @property
    def is_word(self) -> bool:
        return self.kind is TokenKind.KEYWORD or self.kind is TokenKind.IDENTIFIER

    @property
    def is_trivia(self) -> bool:
        return self.kind is TokenKind.WHITESPACE or self.kind is TokenKind.CONTINUATION

    @property
    def canonical(self) -> str | None:
        """The VBE spelling of a keyword token, None for anything else."""
        if self.kind is TokenKind.KEYWORD:
            return canonical_keyword(self.text)
        return None


_WSC_RUN = re.compile(WSC_CLASS + "+")
# An underscore, optional trailing whitespace, and a line terminator.
_CONTINUATION = re.compile("_" + WSC_CLASS + "*(?:\r\n|\r|\n)")
_IDENT_ASCII_RUN = re.compile(r"[A-Za-z0-9_]*")
_TO_LINE_END = re.compile(r"[^\r\n]*")
_DIGITS = re.compile(r"[0-9]*")
_HEX_DIGITS = re.compile(r"[0-9A-Fa-f]*")
_OCT_DIGITS = re.compile(r"[0-7]*")
# The tail of a physical line that continues it: whitespace, an underscore,
# optional whitespace, then the end of the line.
_CONTINUED_TAIL = re.compile(WSC_CLASS + "+_" + WSC_CLASS + "*$")

_OPERATOR_CHARS = frozenset("+-*/\\^=&!?<>")
_PUNCTUATION_CHARS = frozenset(".,;()")
_TWO_CHAR_OPERATORS = frozenset(("<=", ">=", "<>", "=>", "=<", "><"))


def tokenize(src: str) -> list[Token]:
    """Split ``src`` into tokens whose texts concatenate back to ``src``."""
    tokens: list[Token] = []
    append = tokens.append
    n = len(src)
    pos = 0
    line = 0
    # True at the start of a logical line: the only place `#` opens a
    # directive.
    at_line_start = True
    # True where a statement may begin: line start, after a `:` separator,
    # and after a leading line number. Governs whether `Rem` opens a comment.
    at_statement_start = True

    while pos < n:
        ch = src[pos]
        start = pos

        if ch == "\r" or ch == "\n":
            pos += 2 if ch == "\r" and pos + 1 < n and src[pos + 1] == "\n" else 1
            append(Token(TokenKind.NEWLINE, src[start:pos], start, pos, line))
            line += 1
            at_line_start = at_statement_start = True
            continue

        if ch in WSC_CHARS:
            pos = _WSC_RUN.match(src, pos).end()  # type: ignore[union-attr]
            append(Token(TokenKind.WHITESPACE, src[start:pos], start, pos, line))
            cont = _CONTINUATION.match(src, pos)
            if cont is not None:
                end = cont.end()
                append(Token(TokenKind.CONTINUATION, src[pos:end], pos, end, line))
                pos = end
                line += 1
            continue

        if ch == "'":
            pos = _scan_comment(src, pos)
            append(Token(TokenKind.COMMENT, src[start:pos], start, pos, line))
            line += _count_breaks(src, start, pos)
            at_line_start = at_statement_start = False
            continue

        if is_ident_start(ch):
            pos = _IDENT_ASCII_RUN.match(src, pos + 1).end()  # type: ignore[union-attr]
            while pos < n and is_ident_part(src[pos]):
                pos = _IDENT_ASCII_RUN.match(src, pos + 1).end()  # type: ignore[union-attr]
            word = src[start:pos]
            if at_statement_start and word.lower() == "rem":
                pos = _scan_comment(src, pos)
                append(Token(TokenKind.COMMENT, src[start:pos], start, pos, line))
                line += _count_breaks(src, start, pos)
                at_line_start = at_statement_start = False
                continue
            kind = TokenKind.KEYWORD if canonical_keyword(word) else TokenKind.IDENTIFIER
            append(Token(kind, word, start, pos, line))
            at_line_start = at_statement_start = False
            if pos < n and src[pos] in "$%@":
                append(Token(TokenKind.TYPE_SUFFIX, src[pos], pos, pos + 1, line))
                pos += 1
            continue

        if (
            "0" <= ch <= "9"
            or (ch == "." and pos + 1 < n and "0" <= src[pos + 1] <= "9")
            or (ch == "&" and _starts_radix_literal(src, pos + 1))
        ):
            kind, pos = _scan_number(src, pos)
            append(Token(kind, src[start:pos], start, pos, line))
            # A line number leaves the statement after it at a statement
            # start, so `10 Rem note` is a comment.
            at_statement_start = at_line_start and kind is TokenKind.INTEGER
            at_line_start = False
            continue

        if ch == '"':
            pos = _scan_string(src, pos)
            append(Token(TokenKind.STRING, src[start:pos], start, pos, line))
            at_line_start = at_statement_start = False
            continue

        if ch == "[":
            close = src.find("]", pos + 1)
            stop = _line_end(src, pos + 1)
            pos = close + 1 if 0 <= close < stop else stop
            append(Token(TokenKind.BRACKETED, src[start:pos], start, pos, line))
            at_line_start = at_statement_start = False
            if pos < n and src[pos] in "$%@":
                append(Token(TokenKind.TYPE_SUFFIX, src[pos], pos, pos + 1, line))
                pos += 1
            continue

        if ch == "#":
            if at_line_start:
                pos += 1
                append(Token(TokenKind.DIRECTIVE, "#", start, pos, line))
                at_line_start = at_statement_start = False
                continue
            close = src.find("#", pos + 1)
            stop = _line_end(src, pos + 1)
            if 0 <= close < stop and _is_date_body(src[pos + 1 : close]):
                pos = close + 1
                append(Token(TokenKind.DATE, src[start:pos], start, pos, line))
            else:
                pos += 1
                append(Token(TokenKind.OPERATOR, "#", start, pos, line))
            at_line_start = at_statement_start = False
            continue

        if ch == ":":
            if pos + 1 < n and src[pos + 1] == "=":
                pos += 2
                append(Token(TokenKind.OPERATOR, ":=", start, pos, line))
                at_line_start = at_statement_start = False
            else:
                pos += 1
                append(Token(TokenKind.COLON, ":", start, pos, line))
                at_line_start = False
                at_statement_start = True
            continue

        if ch in _OPERATOR_CHARS:
            pos += 2 if src[pos : pos + 2] in _TWO_CHAR_OPERATORS else 1
            append(Token(TokenKind.OPERATOR, src[start:pos], start, pos, line))
            at_line_start = at_statement_start = False
            continue

        if ch in _PUNCTUATION_CHARS:
            pos += 1
            append(Token(TokenKind.PUNCTUATION, ch, start, pos, line))
            at_line_start = at_statement_start = False
            continue

        pos += 1
        append(Token(TokenKind.UNKNOWN, ch, start, pos, line))
        at_line_start = at_statement_start = False

    return tokens


def _line_end(src: str, pos: int) -> int:
    """Offset of the next line terminator at or after ``pos``, or len(src)."""
    return _TO_LINE_END.match(src, pos).end()  # type: ignore[union-attr]


def _count_breaks(src: str, start: int, end: int) -> int:
    """Physical line breaks inside src[start:end] (CRLF counts once)."""
    chunk = src[start:end]
    return chunk.count("\n") + chunk.count("\r") - chunk.count("\r\n")


def _scan_comment(src: str, pos: int) -> int:
    """End of a comment that starts at ``pos``: the end of its logical line."""
    n = len(src)
    while True:
        end = _line_end(src, pos)
        if end >= n or _CONTINUED_TAIL.search(src, pos, end) is None:
            return end
        # The line ends in ` _`: step over the terminator into the next line.
        pos = end + (2 if src[end] == "\r" and end + 1 < n and src[end + 1] == "\n" else 1)


def _scan_string(src: str, pos: int) -> int:
    """End of a string literal that starts at the `"` at ``pos``.

    A string closes at its unpaired `"`. Left open it runs to the end of the
    physical line, except that a line continuation ending that line ends the
    string before the whitespace it starts with (MS-VBAL 3.3.4).
    """
    n = len(src)
    p = pos + 1
    while p < n:
        c = src[p]
        if c == '"':
            if p + 1 < n and src[p + 1] == '"':
                p += 2
                continue
            return p + 1
        if c == "\r" or c == "\n":
            break
        p += 1
    if p < n:
        tail = _CONTINUED_TAIL.search(src, pos + 1, p)
        if tail is not None:
            return tail.start()
    return p


def _starts_radix_literal(src: str, pos: int) -> bool:
    """True when src[pos:] continues an `&` into a hex or octal literal."""
    n = len(src)
    if pos >= n:
        return False
    c = src[pos]
    if c in "hH":
        return pos + 1 < n and src[pos + 1] in "0123456789abcdefABCDEF"
    if c in "oO":
        return pos + 1 < n and "0" <= src[pos + 1] <= "7"
    return "0" <= c <= "7"


def _scan_number(src: str, pos: int) -> tuple[TokenKind, int]:
    """Lex the numeric literal at ``pos`` (MS-VBAL 3.3.2)."""
    n = len(src)
    if src[pos] == "&":
        radix = src[pos + 1]
        if radix in "hH":
            pos = _HEX_DIGITS.match(src, pos + 2).end()  # type: ignore[union-attr]
        elif radix in "oO":
            pos = _OCT_DIGITS.match(src, pos + 2).end()  # type: ignore[union-attr]
        else:
            pos = _OCT_DIGITS.match(src, pos + 1).end()  # type: ignore[union-attr]
        if pos < n and src[pos] in "%&^":
            pos += 1
        return TokenKind.INTEGER, pos

    is_float = False
    pos = _DIGITS.match(src, pos).end()  # type: ignore[union-attr]
    if pos < n and src[pos] == ".":
        after = src[pos + 1] if pos + 1 < n else " "
        # Fractional digits are optional (`1.` is a Double), but a name after
        # the dot would be member access, which a literal cannot have.
        if "0" <= after <= "9" or _exponent_at(src, pos + 1) or not is_ident_start(after):
            is_float = True
            pos = _DIGITS.match(src, pos + 1).end()  # type: ignore[union-attr]
    if _exponent_at(src, pos):
        is_float = True
        pos += 1
        if src[pos] in "+-":
            pos += 1
        pos = _DIGITS.match(src, pos).end()  # type: ignore[union-attr]
    if pos < n:
        c = src[pos]
        if c in "!#@":
            is_float = True
            pos += 1
        elif not is_float and c in "%&^":
            pos += 1
    return (TokenKind.FLOAT if is_float else TokenKind.INTEGER), pos


def _exponent_at(src: str, pos: int) -> bool:
    """True when src[pos:] starts an exponent: [DdEe] [+-] 1*digit."""
    n = len(src)
    if pos >= n or src[pos] not in "DdEe":
        return False
    p = pos + 1
    if p < n and src[p] in "+-":
        p += 1
    return p < n and "0" <= src[p] <= "9"


def _is_date_body(body: str) -> bool:
    """True when the text between a `#` pair is a date literal body."""
    return parse_date_body(body) is not None
