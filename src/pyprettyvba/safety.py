"""The check that formatting did not change what a module means.

Two versions of a module are equivalent when they reduce to the same
signature: the sequence of statements, each a sequence of normalized tokens.
The normalization forgives exactly the changes the rules are allowed to
make, and nothing else:

* whitespace, line continuations and blank lines;
* the letter case of keywords and names, except the name of a Declare with
  no Alias, which is the DLL entry point and is looked up case-sensitively;
* a numeric literal's spelling, as long as its type and value are the same;
  a date literal's spelling, as long as its value is the same;
* `:` against a line break between statements, except inside a single-line
  If, where every colon belongs to the If;
* a comment's marker (`Rem` or `'`) and the spaces around its text;
* a `Let` that opens an assignment, and `EndIf` against `End If`;
* a `Call` against the plain call it means: `Call Foo(a, b)` is `Foo a, b`;
* a declaration of several variables against one declaration each:
  `Dim a, b As Long` declares `a` and `b As Long` just as two Dims would;
* a single-line If against the block If it means: `If a Then b Else c`
  runs what `If a Then` / `b` / `Else` / `c` / `End If` runs.

Whitespace is not always insignificant in VBA, so the signature also records
where it is not: whether `&`, `!`, `#`, `^` and `.` touch the token before
and after them (`s&` is a typed name, `s &` a concatenation; `Foo .Bar`
passes a With member, `Foo.Bar` calls a member).

The engine computes both signatures and refuses to return output whose
signature differs from the input's.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

from .document import Document, LineKind, StatementKind, declaration_items
from .keywords import OPERAND_WORDS
from .lexer import Token, TokenKind
from .literals import date_identity, number_value

__all__ = ["SafetyError", "call_argument_form", "call_argument_parens", "first_difference", "glue_signature", "signature"]


class SafetyError(Exception):
    """Formatting would have changed the meaning of the code."""


_GLUE_SENSITIVE = frozenset(("&", "!", "#", "^", "."))
_SEP = ("sep",)
# The reversed comparison operators mean what the usual ones mean.
_SAME_OPERATOR = {"=>": ">=", "=<": "<=", "><": "<>"}


def signature(text: str) -> list[tuple[Any, ...]]:
    """The normalized statement stream of a module body."""
    doc = Document(text)
    return list(_collapse(_items(doc)))


def first_difference(before: str, after: str) -> str | None:
    """Describe the first place two module bodies differ in meaning."""
    a = signature(before)
    b = signature(after)
    for index, (x, y) in enumerate(zip(a, b, strict=False)):
        if x != y:
            return f"item {index}: {x!r} became {y!r}"
    if len(a) != len(b):
        return f"statement stream length changed from {len(a)} to {len(b)}"
    return None


def _collapse(items: Iterator[tuple[Any, ...]]) -> Iterator[tuple[Any, ...]]:
    previous: tuple[Any, ...] | None = _SEP
    for item in items:
        if item == _SEP and previous == _SEP:
            continue
        yield item
        previous = item
    # A trailing separator is not meaningful.


def _items(doc: Document) -> Iterator[tuple[Any, ...]]:
    tokens = doc.tokens
    for line in doc.lines:
        if line.kind is LineKind.BLANK:
            continue
        if line.kind is LineKind.ATTRIBUTE:
            # No rule edits an Attribute line; its layout (line breaks after
            # a ` _`, spaces) is all that can change.
            yield ("attribute", tuple(
                _comment_body(tokens[j].text) if tokens[j].kind is TokenKind.COMMENT else tokens[j].text
                for j in range(line.first, line.stop)
                if not tokens[j].is_trivia and tokens[j].kind is not TokenKind.NEWLINE
            ))
            yield _SEP
            continue
        if line.label is not None:
            label = tokens[line.label]
            if label.kind is TokenKind.INTEGER:
                yield ("line-number", int(label.text))
            else:
                yield ("label", label.key)
            yield _SEP
        for statement in line.statements:
            declaration = declaration_items(tokens, statement)
            if declaration is not None and len(declaration[1]) > 1:
                # One statement per item, as the declaration means.
                keywords, items = declaration
                for item in items:
                    for j in keywords + item:
                        yield from _token_items(doc, j, exact=False)
                    yield _SEP
                continue
            pinned = _pinned_name(tokens, statement)
            single_if = statement.kind is StatementKind.IF_SINGLE
            if single_if:
                expanded = _single_if_as_block(doc, statement)
                if expanded is not None:
                    yield from expanded
                    continue
            first = True
            indices = call_argument_form(tokens, statement.tokens)
            for position, j in enumerate(indices):
                tok = tokens[j]
                if first and tok.kind is TokenKind.KEYWORD and tok.lower == "let":
                    first = False
                    continue
                first = False
                if tok.kind is TokenKind.COLON:
                    # Only a single-line If keeps colons inside a statement.
                    yield ("colon",) if single_if else _SEP
                    continue
                yield from _token_items(doc, j, exact=position == pinned)
            yield _SEP
        if line.comment is not None:
            yield ("comment", _comment_body(tokens[line.comment].text))
        yield _SEP


def call_argument_form(tokens: list[Token], indices: list[int]) -> list[int]:
    """A `Call` statement as the plain call it means: `Call Foo(a, b)` is
    `Foo a, b`, `Call Log((x))` is `Log (x)`, and `Call Foo()` is `Foo`.
    The Call goes, and so does the parenthesis pair that wraps the whole
    argument list: the first `(` after the callee, when its match is the
    statement's last token. Anything else is returned as it is."""
    if len(indices) < 2 or tokens[indices[0]].kind is not TokenKind.KEYWORD or tokens[indices[0]].lower != "call":
        return indices
    rest = indices[1:]
    opening = call_argument_parens(tokens, rest)
    if opening is None:
        return rest
    return rest[:opening] + rest[opening + 1 : -1]


def call_argument_parens(tokens: list[Token], indices: list[int]) -> int | None:
    """The position in ``indices`` of the `(` that opens a call's argument
    list, when its match is the last token; None when there is no such pair."""
    if not indices or tokens[indices[-1]].text != ")":
        return None
    depth = 0
    opening = 0
    for position, j in enumerate(indices):
        tok = tokens[j]
        if tok.kind is not TokenKind.PUNCTUATION:
            continue
        if tok.text == "(":
            if depth == 0:
                opening = position
            depth += 1
        elif tok.text == ")":
            depth -= 1
            if depth == 0:
                return opening if position == len(indices) - 1 else None
    return None


def _single_if_as_block(doc: Document, statement: Any) -> list[tuple[Any, ...]] | None:
    """A single-line If as the block If it means: `If a Then b Else c` is
    `If a Then`, `b`, `Else`, `c`, `End If`, with the colons of each arm
    separating its statements. A body that is itself an If is left as
    written, since which If an Else belongs to is not read here."""
    tokens = doc.tokens
    indices = statement.tokens
    depth = 0
    then_at = None
    for position, j in enumerate(indices):
        tok = tokens[j]
        if tok.kind is TokenKind.PUNCTUATION and tok.text in ("(", ")"):
            depth += 1 if tok.text == "(" else -1
        elif depth == 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "then":
            then_at = position
            break
    if then_at is None:
        return None
    body = indices[then_at + 1 :]
    if body and tokens[body[0]].kind is TokenKind.KEYWORD and tokens[body[0]].lower == "if":
        return None
    items: list[tuple[Any, ...]] = []
    for j in indices[: then_at + 1]:
        items.extend(_token_items(doc, j, exact=False))
    items.append(_SEP)
    depth = 0
    for j in body:
        tok = tokens[j]
        if tok.kind is TokenKind.PUNCTUATION and tok.text in ("(", ")"):
            depth += 1 if tok.text == "(" else -1
        if tok.kind is TokenKind.COLON:
            items.append(_SEP)
        elif depth == 0 and tok.kind is TokenKind.KEYWORD and tok.lower == "else":
            items += [_SEP, ("word", "else"), _SEP]
        else:
            items.extend(_token_items(doc, j, exact=False))
    items += [_SEP, ("word", "end"), ("word", "if"), _SEP]
    return items


def _pinned_name(tokens: list[Token], statement: Any) -> int | None:
    """Position in the statement of a Declare name that must keep its case."""
    if statement.kind is not StatementKind.DECLARE or statement.name is None:
        return None
    for j in statement.tokens:
        if tokens[j].kind is TokenKind.IDENTIFIER and tokens[j].lower == "alias":
            return None
    return int(statement.name)


def _token_items(doc: Document, j: int, *, exact: bool) -> Iterator[tuple[Any, ...]]:
    tok = doc.tokens[j]
    kind = tok.kind
    if kind is TokenKind.KEYWORD or kind is TokenKind.IDENTIFIER:
        if exact:
            yield ("name", tok.text)
        elif kind is TokenKind.KEYWORD and tok.lower == "endif":
            yield ("word", "end")
            yield ("word", "if")
        else:
            yield ("word", tok.key)
        return
    if kind is TokenKind.INTEGER or kind is TokenKind.FLOAT:
        value = number_value(tok.text)
        yield ("number", value) if value is not None else ("number-text", tok.text)
        return
    if kind is TokenKind.DATE:
        yield ("date", date_identity(tok.text))
        return
    if kind is TokenKind.OPERATOR or kind is TokenKind.PUNCTUATION:
        if tok.text in _GLUE_SENSITIVE:
            yield (kind.value, tok.text, *glue_signature(doc.tokens, j))
        else:
            yield (kind.value, _SAME_OPERATOR.get(tok.text, tok.text))
        return
    if kind is TokenKind.UNKNOWN:
        yield ("unknown", tok.text, _glued(doc.tokens, j, -1), _glued(doc.tokens, j, +1))
        return
    yield (kind.value, tok.text)


def glue_signature(tokens: list[Token], j: int) -> tuple[bool | None, bool | None]:
    """Whether a glue-sensitive token touches its neighbours, where it matters.

    Returns (before, after); None where whitespace cannot change the reading.
    After a name, `& ! # ^` glued on are type characters and `!` `.` are
    member access; spaced, they are operators, or a With member passed as an
    argument. After anything else `&`, `#` and `^` are plain operators and
    `.` a With member, whatever the spacing. The token after `! # ^ .` is
    recorded too: `rs!Field`, `#1` and `.Name` read differently apart.
    """
    tok = tokens[j]
    prev = _neighbour(tokens, j, -1)
    nxt = _neighbour(tokens, j, +1)
    name_before = prev is not None and (
        _name_like(prev) or (tok.text in ("!", ".") and prev.text == ")")
    )
    before = _glued(tokens, j, -1) if name_before else None
    if nxt is None or nxt.kind in (TokenKind.COLON, TokenKind.COMMENT):
        # Nothing follows on the line for the token to join.
        after: bool | None = None
    elif tok.text == ".":
        after = _glued(tokens, j, +1)
    elif before:
        # A type character or a bang: `rs!Field` and `rs! Field` differ, but
        # `x& = 1` and `x&= 1` do not.
        after = _glued(tokens, j, +1) if nxt is not None and _operand_like(nxt) else None
    elif tok.text == "!":
        after = _glued(tokens, j, +1)
    else:
        after = None
    return before, after


def _operand_like(tok: Token) -> bool:
    return tok.kind in (
        TokenKind.IDENTIFIER, TokenKind.KEYWORD, TokenKind.BRACKETED, TokenKind.INTEGER,
        TokenKind.FLOAT, TokenKind.STRING, TokenKind.DATE,
    )


def _neighbour(tokens: list[Token], j: int, step: int) -> Token | None:
    """The nearest non-whitespace token on one side, on the same logical line."""
    k = j + step
    while 0 <= k < len(tokens):
        tok = tokens[k]
        if tok.kind is TokenKind.NEWLINE:
            return None
        if tok.kind is not TokenKind.WHITESPACE and tok.kind is not TokenKind.CONTINUATION:
            return tok
        k += step
    return None


def _name_like(tok: Token) -> bool:
    """A token after which `&`, `!`, `#` or `^` can be a type character."""
    if tok.kind is TokenKind.IDENTIFIER or tok.kind is TokenKind.BRACKETED:
        return True
    return tok.kind is TokenKind.KEYWORD and tok.lower in OPERAND_WORDS


def _glued(tokens: list[Token], j: int, step: int) -> bool:
    """True when token j touches its neighbour on one side (no whitespace)."""
    k = j + step
    if not 0 <= k < len(tokens):
        return False
    return tokens[k].kind not in (
        TokenKind.WHITESPACE,
        TokenKind.CONTINUATION,
        TokenKind.NEWLINE,
        TokenKind.COMMENT,
    )


def _comment_body(text: str) -> str:
    """A comment's text, without its marker and the whitespace around each of its lines."""
    if text[:3].lower() == "rem":
        body = text[3:]
    elif text.startswith("'"):
        body = text[1:]
    else:
        body = text
    # A comment continued with ` _` spans physical lines; their breaks and
    # trailing whitespace are layout.
    lines = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return "\n".join(line.rstrip(" \t") for line in lines).strip()
