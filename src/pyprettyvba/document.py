"""The module model the rules read: logical lines, statements, and headers.

A VBA module file can carry text that is not VBA: the VERSION/BEGIN block of
an exported class, a UserForm's designer block, and the module's own
Attribute lines. ``split_header`` sets that aside so it is never touched.

The rest is cut into logical lines (a physical line plus the lines its
continuations join to it), each logical line into statements at its `:`
separators, and each statement classified by what it opens, closes or
declares. A single-line If owns everything after its Then, colons and Else
included, so nothing later splits it: `If a Then b: c` runs `c` only when
`a` holds (measured, tests/oracle/vbe_rendering.json).
"""

from __future__ import annotations

import enum
import re
from dataclasses import dataclass, field

from .lexer import Token, TokenKind, tokenize

__all__ = [
    "Document",
    "LineKind",
    "LogicalLine",
    "Statement",
    "StatementKind",
    "split_header",
]


# --- export headers ------------------------------------------------------

_VERSION_RE = re.compile(r"^VERSION\s+\d", re.IGNORECASE)
_OBJECT_RE = re.compile(r"^Object\s*=", re.IGNORECASE)
_BEGIN_RE = re.compile(r"^(?:BEGIN|BeginProperty)(?:\s|$)", re.IGNORECASE)
_END_RE = re.compile(r"^(?:END|EndProperty)\s*$", re.IGNORECASE)
_ATTRIBUTE_RE = re.compile(r"^Attribute\s+VB_", re.IGNORECASE)


def split_header(text: str) -> tuple[str, str]:
    """Split an exported module into (header, body).

    The header is the VERSION line, any VB6 `Object =` lines, the BEGIN ...
    END block (a class's settings, a form's designer controls), and the
    `Attribute VB_...` lines that follow. Source without a VERSION line keeps
    only its leading Attribute lines as a header. ``header + body == text``.
    """
    lines = text.splitlines(keepends=True)
    i = 0
    if lines and _VERSION_RE.match(lines[0]):
        i = 1
        while i < len(lines) and _OBJECT_RE.match(lines[i].strip()):
            i += 1
        if i < len(lines) and _BEGIN_RE.match(lines[i].strip()):
            depth = 0
            while i < len(lines):
                stripped = lines[i].strip()
                if _BEGIN_RE.match(stripped):
                    depth += 1
                elif _END_RE.match(stripped):
                    depth -= 1
                i += 1
                if depth <= 0:
                    break
    while i < len(lines) and _ATTRIBUTE_RE.match(lines[i]):
        i += 1
    cut = sum(len(line) for line in lines[:i])
    return text[:cut], text[cut:]


# --- lines and statements ------------------------------------------------

class LineKind(enum.StrEnum):
    BLANK = "blank"
    COMMENT = "comment"
    CODE = "code"
    DIRECTIVE = "directive"
    ATTRIBUTE = "attribute"


class StatementKind(enum.StrEnum):
    OTHER = "other"
    PROC_START = "proc-start"
    PROC_END = "proc-end"
    DECLARE = "declare"
    TYPE_START = "type-start"
    TYPE_END = "type-end"
    ENUM_START = "enum-start"
    ENUM_END = "enum-end"
    IF_BLOCK = "if-block"
    IF_SINGLE = "if-single"
    ELSEIF = "elseif"
    ELSE = "else"
    END_IF = "end-if"
    SELECT = "select"
    CASE = "case"
    END_SELECT = "end-select"
    FOR = "for"
    NEXT = "next"
    DO = "do"
    LOOP = "loop"
    WHILE = "while"
    WEND = "wend"
    WITH = "with"
    END_WITH = "end-with"
    VARIABLES = "variables"  # Dim / Private / Public / Global / Static / Const / ReDim
    OPTION = "option"
    DEFTYPE = "deftype"
    EVENT = "event"
    IMPLEMENTS = "implements"
    ATTRIBUTE = "attribute"
    CC_IF = "cc-if"
    CC_ELSEIF = "cc-elseif"
    CC_ELSE = "cc-else"
    CC_END_IF = "cc-end-if"
    CC_CONST = "cc-const"
    CC_OTHER = "cc-other"


# Statements that open a block, close one, or sit in the middle of one.
OPENERS = frozenset(
    {
        StatementKind.PROC_START, StatementKind.TYPE_START, StatementKind.ENUM_START,
        StatementKind.IF_BLOCK, StatementKind.SELECT, StatementKind.FOR,
        StatementKind.DO, StatementKind.WHILE, StatementKind.WITH,
    }
)
CLOSERS = frozenset(
    {
        StatementKind.PROC_END, StatementKind.TYPE_END, StatementKind.ENUM_END,
        StatementKind.END_IF, StatementKind.END_SELECT, StatementKind.NEXT,
        StatementKind.LOOP, StatementKind.WEND, StatementKind.END_WITH,
    }
)


@dataclass
class Statement:
    """A run of significant tokens between separators.

    ``tokens`` are indices into ``Document.tokens`` and skip whitespace,
    continuations and comments.
    """

    tokens: list[int]
    kind: StatementKind = StatementKind.OTHER
    # For NEXT: how many loops it closes (`Next i, j` closes two).
    count: int = 1
    # Index into ``tokens`` of the name a declaration introduces, where one
    # does (procedure, Declare, Type, Enum, Event).
    name: int | None = None
    # The procedure keyword for PROC_START / PROC_END: sub, function, property.
    proc_kind: str | None = None
    # True when a `:` separator follows the statement on its line.
    colon_after: bool = False


@dataclass
class LogicalLine:
    """A physical line and every line its continuations join to it."""

    index: int
    # Token index range [first, stop): stop is one past the NEWLINE, or the
    # end of the stream on a last line without one.
    first: int
    stop: int
    start: int  # character offset of the line's first character
    end: int  # character offset one past its terminator
    first_physical: int
    last_physical: int
    kind: LineKind = LineKind.BLANK
    statements: list[Statement] = field(default_factory=list)
    # Token index of a leading line label (identifier) or line number.
    label: int | None = None
    # Token index of the `:` after a label, when present.
    label_colon: int | None = None
    # Token index of the comment that ends the line, if any.
    comment: int | None = None
    # Token index of the first significant token (label, code, directive
    # marker or comment), None on a blank line.
    head: int | None = None


class Document:
    """A module body cut into tokens, logical lines and statements."""

    def __init__(self, text: str) -> None:
        self.text = text
        self.tokens: list[Token] = tokenize(text)
        self.lines: list[LogicalLine] = []
        # Character offset where each physical line starts.
        self.physical_starts: list[int] = [0]
        for tok in self.tokens:
            if tok.kind is TokenKind.NEWLINE or tok.kind is TokenKind.CONTINUATION:
                self.physical_starts.append(tok.end)
            elif tok.kind is TokenKind.COMMENT and ("\n" in tok.text or "\r" in tok.text):
                self.physical_starts.extend(_breaks_in(tok))
        if self.physical_starts[-1] == len(text) and len(self.physical_starts) > 1 and text:
            # A final terminator does not start a physical line.
            self.physical_starts.pop()
        self._build_lines()

    # -- construction --------------------------------------------------------

    def prev_code(self, j: int) -> int | None:
        """The nearest significant token before j on its logical line."""
        tokens = self.tokens
        k = j - 1
        while k >= 0:
            kind = tokens[k].kind
            if kind is TokenKind.NEWLINE:
                return None
            if kind is not TokenKind.WHITESPACE and kind is not TokenKind.CONTINUATION:
                return k
            k -= 1
        return None

    def next_code(self, j: int) -> int | None:
        """The nearest significant token after j on its logical line."""
        tokens = self.tokens
        k = j + 1
        while k < len(tokens):
            kind = tokens[k].kind
            if kind is TokenKind.NEWLINE:
                return None
            if kind is not TokenKind.WHITESPACE and kind is not TokenKind.CONTINUATION:
                return k
            k += 1
        return None

    def is_member_name(self, j: int) -> bool:
        """True when token j follows `.` or `!`: a member, not a free name."""
        k = self.prev_code(j)
        if k is None:
            return False
        prev = self.tokens[k]
        return (prev.kind is TokenKind.PUNCTUATION and prev.text == ".") or (
            prev.kind is TokenKind.OPERATOR and prev.text == "!"
        )

    def _build_lines(self) -> None:
        tokens = self.tokens
        count = len(tokens)
        i = 0
        while i < count or not self.lines:
            first = i
            while i < count and tokens[i].kind is not TokenKind.NEWLINE:
                i += 1
            stop = i + 1 if i < count else i
            start = tokens[first].start if first < count else len(self.text)
            end = tokens[stop - 1].end if stop > first else start
            first_physical = tokens[first].line if first < count else len(self.physical_starts) - 1
            last_physical = first_physical
            for tok in tokens[first:stop]:
                last_physical = max(last_physical, tok.line + _break_count(tok))
            if stop > first and tokens[stop - 1].kind is TokenKind.NEWLINE:
                last_physical = tokens[stop - 1].line
            line = LogicalLine(
                index=len(self.lines),
                first=first,
                stop=stop,
                start=start,
                end=end,
                first_physical=first_physical,
                last_physical=last_physical,
            )
            self._analyze(line)
            self.lines.append(line)
            i = stop
            if i >= count:
                break

    def _analyze(self, line: LogicalLine) -> None:
        tokens = self.tokens
        significant = [
            j
            for j in range(line.first, line.stop)
            if tokens[j].kind not in _NON_SIGNIFICANT
        ]
        if not significant:
            line.kind = LineKind.BLANK
            return
        line.head = significant[0]
        if tokens[significant[-1]].kind is TokenKind.COMMENT:
            line.comment = significant[-1]
            significant = significant[:-1]
        if not significant:
            line.kind = LineKind.COMMENT
            return
        head = tokens[significant[0]]
        if head.kind is TokenKind.DIRECTIVE:
            line.kind = LineKind.DIRECTIVE
            statement = Statement(tokens=significant)
            statement.kind = _classify_directive(tokens, significant)
            line.statements = [statement]
            return
        if head.kind in (TokenKind.IDENTIFIER, TokenKind.KEYWORD) and head.lower == "attribute":
            line.kind = LineKind.ATTRIBUTE
            line.statements = [Statement(tokens=significant, kind=StatementKind.ATTRIBUTE)]
            return
        line.kind = LineKind.CODE
        rest = significant
        # A line number, or an identifier and a colon, at the start of a line
        # is a label (MS-VBAL 5.4.1: an identifier followed by ":" at the
        # beginning of a line is always a label).
        if (
            head.kind is TokenKind.INTEGER
            and head.text.isdigit()
            and significant[0] == line.first + _leading_ws(tokens, line)
        ):
            line.label = significant[0]
            rest = significant[1:]
            if rest and tokens[rest[0]].kind is TokenKind.COLON:
                line.label_colon = rest[0]
                rest = rest[1:]
        elif (
            head.kind is TokenKind.IDENTIFIER
            and len(significant) > 1
            and tokens[significant[1]].kind is TokenKind.COLON
        ):
            line.label = significant[0]
            line.label_colon = significant[1]
            rest = significant[2:]
        line.statements = _split_statements(tokens, rest)
        for statement in line.statements:
            _classify(tokens, statement)

    # -- queries -------------------------------------------------------------

    def line_text(self, line: LogicalLine) -> str:
        return self.text[line.start : line.end]

    def physical_line_of(self, offset: int) -> int:
        """0-based physical line containing character ``offset``."""
        import bisect

        return max(0, bisect.bisect_right(self.physical_starts, offset) - 1)

    def significant(self, line: LogicalLine) -> list[int]:
        """Indices of the line's significant tokens, the comment included."""
        return [j for j in range(line.first, line.stop) if self.tokens[j].kind not in _NON_SIGNIFICANT]

    def indent_of(self, line: LogicalLine) -> str:
        """The whitespace a logical line starts with."""
        first = self.tokens[line.first] if line.first < len(self.tokens) else None
        if first is not None and first.kind is TokenKind.WHITESPACE:
            return first.text
        return ""


_NON_SIGNIFICANT = frozenset((TokenKind.WHITESPACE, TokenKind.CONTINUATION, TokenKind.NEWLINE))


def _breaks_in(tok: Token) -> list[int]:
    """Offsets just after each line break inside a multi-line comment."""
    out: list[int] = []
    text = tok.text
    k = 0
    while k < len(text):
        c = text[k]
        if c == "\r":
            k += 2 if k + 1 < len(text) and text[k + 1] == "\n" else 1
            out.append(tok.start + k)
            continue
        if c == "\n":
            k += 1
            out.append(tok.start + k)
            continue
        k += 1
    return out


def _break_count(tok: Token) -> int:
    if tok.kind is TokenKind.CONTINUATION:
        return 1
    if tok.kind is TokenKind.COMMENT:
        return len(_breaks_in(tok))
    return 0


def _leading_ws(tokens: list[Token], line: LogicalLine) -> int:
    return 1 if line.first < line.stop and tokens[line.first].kind is TokenKind.WHITESPACE else 0


def _split_statements(tokens: list[Token], indices: list[int]) -> list[Statement]:
    """Cut a line's significant tokens into statements.

    Statements end at `:` separators. A single-line If takes every token
    after its Then, colons and Else included, since all of them run only
    when its condition holds. A block `Else` or `ElseIf ... Then` ends where
    its keywords do, so the statement written after it on the same line is a
    statement of its own: `Else If d Then` opens a nested block If.
    """
    statements: list[Statement] = []
    current: list[int] = []
    k = 0
    while k < len(indices):
        j = indices[k]
        tok = tokens[j]
        if tok.kind is TokenKind.COLON:
            if current:
                statements.append(Statement(tokens=current, colon_after=True))
            elif statements:
                statements[-1].colon_after = True
            current = []
            k += 1
            continue
        if not current and tok.kind is TokenKind.KEYWORD:
            word = tok.lower
            if word == "if":
                then_at = _find_then(tokens, indices, k + 1)
                if then_at is not None and then_at + 1 < len(indices):
                    # Tokens after Then: a single-line If, which owns the rest.
                    statements.append(Statement(tokens=list(indices[k:])))
                    return statements
            elif word == "else" and k + 1 < len(indices) and tokens[indices[k + 1]].kind is not TokenKind.COLON:
                statements.append(Statement(tokens=[j]))
                k += 1
                continue
            elif word == "elseif":
                then_at = _find_then(tokens, indices, k + 1)
                if then_at is not None and then_at + 1 < len(indices):
                    statements.append(Statement(tokens=list(indices[k : then_at + 1])))
                    k = then_at + 1
                    continue
        current.append(j)
        k += 1
    if current:
        statements.append(Statement(tokens=current))
    return statements


def _find_then(tokens: list[Token], indices: list[int], k: int) -> int | None:
    """Position in ``indices`` of the Then that ends an If condition."""
    depth = 0
    while k < len(indices):
        tok = tokens[indices[k]]
        if tok.kind is TokenKind.PUNCTUATION:
            if tok.text == "(":
                depth += 1
            elif tok.text == ")":
                depth -= 1
        elif depth <= 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "then":
            return k
        k += 1
    return None


_PROC_MODIFIERS = frozenset(("public", "private", "friend", "global", "static"))
_TYPE_MODIFIERS = frozenset(("public", "private", "global"))
_VARIABLE_WORDS = frozenset(("dim", "static", "const", "redim", "private", "public", "global"))


def _word(tokens: list[Token], statement: Statement, k: int) -> str:
    if 0 <= k < len(statement.tokens):
        tok = tokens[statement.tokens[k]]
        if tok.is_word:
            return tok.lower
    return ""


def _classify(tokens: list[Token], statement: Statement) -> None:
    """Set ``statement.kind`` (and its name, count and procedure kind)."""
    w0 = _word(tokens, statement, 0)
    w1 = _word(tokens, statement, 1)
    first = tokens[statement.tokens[0]]
    # A name after `.` is a member, never a statement keyword: `.Print x`.
    if first.kind is not TokenKind.KEYWORD and first.kind is not TokenKind.IDENTIFIER:
        return
    if w0 == "end" or w0 == "endif":
        closer = {
            "sub": StatementKind.PROC_END,
            "function": StatementKind.PROC_END,
            "property": StatementKind.PROC_END,
            "if": StatementKind.END_IF,
            "select": StatementKind.END_SELECT,
            "with": StatementKind.END_WITH,
            "type": StatementKind.TYPE_END,
            "enum": StatementKind.ENUM_END,
        }
        if w0 == "endif":
            statement.kind = StatementKind.END_IF
        elif w1 in closer:
            statement.kind = closer[w1]
            if statement.kind is StatementKind.PROC_END:
                statement.proc_kind = w1
        return
    if w0 == "if":
        # _split_statements gave a single-line If the whole rest of the line.
        then_at = _find_then(tokens, statement.tokens, 1)
        if then_at is not None and then_at + 1 < len(statement.tokens):
            statement.kind = StatementKind.IF_SINGLE
        elif then_at is not None and statement.colon_after:
            # `If x Then:` is a single-line If with nothing to run: the line
            # after it runs unconditionally (measured).
            statement.kind = StatementKind.IF_SINGLE
        else:
            statement.kind = StatementKind.IF_BLOCK
        return
    if w0 == "elseif":
        statement.kind = StatementKind.ELSEIF
        return
    if w0 == "else":
        statement.kind = StatementKind.ELSE
        return
    if w0 == "select" and w1 == "case":
        statement.kind = StatementKind.SELECT
        return
    if w0 == "case":
        statement.kind = StatementKind.CASE
        return
    if w0 == "for":
        statement.kind = StatementKind.FOR
        return
    if w0 == "next":
        statement.kind = StatementKind.NEXT
        statement.count = 1 + sum(
            1 for j in statement.tokens[1:] if tokens[j].kind is TokenKind.PUNCTUATION and tokens[j].text == ","
        )
        return
    if w0 == "do":
        statement.kind = StatementKind.DO
        return
    if w0 == "loop":
        statement.kind = StatementKind.LOOP
        return
    if w0 == "while":
        statement.kind = StatementKind.WHILE
        return
    if w0 == "wend":
        statement.kind = StatementKind.WEND
        return
    if w0 == "with":
        statement.kind = StatementKind.WITH
        return
    if w0 == "option":
        statement.kind = StatementKind.OPTION
        return
    if w0 in _DEFTYPE_WORDS:
        statement.kind = StatementKind.DEFTYPE
        return
    if w0 == "implements":
        statement.kind = StatementKind.IMPLEMENTS
        return

    # Declarations: modifiers, then the word that says what is declared.
    k = 0
    while _word(tokens, statement, k) in _PROC_MODIFIERS:
        k += 1
    declared = _word(tokens, statement, k)
    if declared in ("sub", "function"):
        statement.kind = StatementKind.PROC_START
        statement.proc_kind = declared
        statement.name = k + 1 if k + 1 < len(statement.tokens) else None
        return
    if declared == "property" and _word(tokens, statement, k + 1) in ("get", "let", "set"):
        statement.kind = StatementKind.PROC_START
        statement.proc_kind = "property"
        statement.name = k + 2 if k + 2 < len(statement.tokens) else None
        return
    if declared == "declare":
        statement.kind = StatementKind.DECLARE
        m = k + 1
        if _word(tokens, statement, m) == "ptrsafe":
            m += 1
        if _word(tokens, statement, m) in ("sub", "function"):
            statement.name = m + 1 if m + 1 < len(statement.tokens) else None
        return
    if declared == "event":
        statement.kind = StatementKind.EVENT
        statement.name = k + 1 if k + 1 < len(statement.tokens) else None
        return
    if declared in ("type", "enum"):
        modifiers = [_word(tokens, statement, m) for m in range(k)]
        if all(word in _TYPE_MODIFIERS for word in modifiers):
            name_at = k + 1
            # Any word can name one: `Private Enum LongPtr` is how 32-bit code
            # declares the 64-bit type name for itself.
            if name_at < len(statement.tokens) and tokens[statement.tokens[name_at]].kind in (
                TokenKind.IDENTIFIER,
                TokenKind.KEYWORD,
                TokenKind.BRACKETED,
            ):
                statement.kind = StatementKind.TYPE_START if declared == "type" else StatementKind.ENUM_START
                statement.name = name_at
                return
    if w0 in _VARIABLE_WORDS:
        statement.kind = StatementKind.VARIABLES
        return


_DEFTYPE_WORDS = frozenset(
    (
        "defbool", "defbyte", "defcur", "defdate", "defdbl", "defint", "deflng",
        "deflnglng", "deflngptr", "defobj", "defsng", "defstr", "defvar", "defdec",
    )
)


def _classify_directive(tokens: list[Token], indices: list[int]) -> StatementKind:
    word = tokens[indices[1]].lower if len(indices) > 1 and tokens[indices[1]].is_word else ""
    after = tokens[indices[2]].lower if len(indices) > 2 and tokens[indices[2]].is_word else ""
    if word == "if":
        return StatementKind.CC_IF
    if word == "elseif":
        return StatementKind.CC_ELSEIF
    if word == "else":
        return StatementKind.CC_ELSE
    if word == "endif" or (word == "end" and after == "if"):
        return StatementKind.CC_END_IF
    if word == "const":
        return StatementKind.CC_CONST
    return StatementKind.CC_OTHER
