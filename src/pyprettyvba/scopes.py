"""What a module declares, and where each declaration is visible.

identifier-case spells a name the way its declaration spells it. The VBE
does the same when it reads a module in: `Dim total` makes every later
`TOTAL` and `Total` read `total`, even a use written before the
declaration (measured, tests/oracle/vbe_rendering.json).

Declarations are collected per procedure (parameters, locals, labels) and
per module (variables, constants, procedures, Declares, Types, Enums and
their members, Events, `#Const`s). A procedure's own names shadow the
module's. When a name is declared twice (in two `#If` arms, or as locals of
two procedures), the last declaration gives its spelling, as in the VBE,
which keeps one spelling per name and lets each declaration it reads
replace the one before (measured: `Dim zqCount` in one procedure and
`Dim ZQCOUNT` in the next leave every use reading `ZQCOUNT`). A `ReDim` of
a name declared elsewhere declares nothing.

The members of a Type are declarations too: the VBE spells every use of
the name the way the member does, even a call of the library function
`Second` when a member is called `second` (measured).

The name of a Declare without an Alias is pinned: it is also the DLL entry
point, which Windows looks up case-sensitively, so respelling it would
break the call at run time.
"""

from __future__ import annotations

import bisect
from dataclasses import dataclass, field

from .document import Document, LineKind, Statement, StatementKind
from .lexer import TokenKind

__all__ = ["Declarations", "ProcedureScope", "collect_declarations"]

_LEADING_DECLARATION_WORDS = frozenset(
    ("dim", "private", "public", "global", "static", "const", "redim", "preserve", "withevents", "shared")
)
_PARAMETER_WORDS = frozenset(("optional", "byval", "byref", "paramarray"))


@dataclass
class ProcedureScope:
    """A procedure's own names and the logical lines it spans."""

    first_line: int
    last_line: int
    names: dict[str, str] = field(default_factory=dict)


@dataclass
class Declarations:
    module: dict[str, str] = field(default_factory=dict)
    procedures: list[ProcedureScope] = field(default_factory=list)
    # Token indices that must keep their spelling.
    pinned: set[int] = field(default_factory=set)
    # Names other modules can use: Public (or default-public) declarations.
    public: dict[str, str] = field(default_factory=dict)
    # Every name declared anywhere in the module, spelled by its last declaration.
    anywhere: dict[str, str] = field(default_factory=dict)
    # Token indices of the names Type members declare.
    members: set[int] = field(default_factory=set)
    # Names code can reach after a `.`: Type and Enum members, and whatever
    # the module does not declare Private (a class's properties and methods,
    # a standard module's Public procedures, reached as `Module1.Foo`).
    member_names: dict[str, str] = field(default_factory=dict)
    # The parameters of the module's procedures: the names `name:=` uses.
    parameters: dict[str, str] = field(default_factory=dict)

    def scope_at(self, line_index: int) -> ProcedureScope | None:
        starts = [scope.first_line for scope in self.procedures]
        i = bisect.bisect_right(starts, line_index) - 1
        if i >= 0 and self.procedures[i].first_line <= line_index <= self.procedures[i].last_line:
            return self.procedures[i]
        return None

    def spelling(self, lower: str, line_index: int) -> str | None:
        scope = self.scope_at(line_index)
        if scope is not None and lower in scope.names:
            return scope.names[lower]
        return self.module.get(lower)


def collect_declarations(doc: Document, module_kind: str | None = None) -> Declarations:
    """Collect what ``doc`` declares."""
    decl = Declarations()
    tokens = doc.tokens
    scope: ProcedureScope | None = None
    enum_public: bool | None = None  # inside an Enum: is it Public?
    in_type = False

    def add(target: dict[str, str], j: int, *, replaces: bool = True) -> None:
        tok = tokens[j]
        if tok.kind is TokenKind.IDENTIFIER:
            if replaces:
                target[tok.key] = tok.text
                decl.anywhere[tok.key] = tok.text
            else:
                target.setdefault(tok.key, tok.text)
                decl.anywhere.setdefault(tok.key, tok.text)

    def add_module(j: int, public: bool, *, member: bool = False) -> None:
        add(decl.module, j)
        if public:
            add(decl.public, j)
        if member:
            add(decl.member_names, j)

    standard = module_kind in (None, "standard")
    for line in doc.lines:
        if line.kind is LineKind.DIRECTIVE and line.statements:
            statement = line.statements[0]
            if statement.kind is StatementKind.CC_CONST and len(statement.tokens) > 2:
                add(decl.module, statement.tokens[2])
            continue
        if line.kind is not LineKind.CODE:
            continue
        if line.label is not None and scope is not None and tokens[line.label].kind is TokenKind.IDENTIFIER:
            add(scope.names, line.label)
        for statement in line.statements:
            kind = statement.kind
            words = [tokens[j].lower for j in statement.tokens if tokens[j].is_word]
            is_private = bool(words) and words[0] in ("private", "dim")
            if kind is StatementKind.PROC_START:
                name = _name_token(doc, statement)
                if name is not None:
                    add_module(name, standard and not is_private, member=not is_private)
                scope = ProcedureScope(first_line=line.index, last_line=len(doc.lines) - 1)
                decl.procedures.append(scope)
                for j in _parameter_names(doc, statement):
                    add(scope.names, j)
                    add(decl.parameters, j)
            elif kind is StatementKind.PROC_END:
                if scope is not None:
                    scope.last_line = line.index
                scope = None
            elif kind is StatementKind.VARIABLES:
                redim = bool(words) and words[0] == "redim"
                for j in _variable_names(doc, statement):
                    if scope is not None:
                        # `ReDim` declares only a name nothing else declares.
                        add(scope.names, j, replaces=not redim)
                    else:
                        exposed = bool(words) and words[0] in ("public", "global")
                        add_module(j, standard and exposed, member=exposed)
            elif kind is StatementKind.DECLARE:
                name = _name_token(doc, statement)
                if name is not None:
                    add_module(name, standard and not is_private)
                    if not any(tokens[j].lower == "alias" for j in statement.tokens):
                        decl.pinned.add(name)
                    # A Declare's parameters are in scope nowhere, but they
                    # are declarations: `name:=` in a call uses them.
                    for j in _parameter_names(doc, statement):
                        add(decl.parameters, j)
            elif kind in (StatementKind.TYPE_START, StatementKind.ENUM_START, StatementKind.EVENT):
                name = _name_token(doc, statement)
                if name is not None:
                    add_module(name, standard and not is_private, member=kind is StatementKind.EVENT)
                if kind is StatementKind.ENUM_START:
                    enum_public = standard and not is_private
                in_type = kind is StatementKind.TYPE_START
            elif kind is StatementKind.ENUM_END:
                enum_public = None
            elif kind is StatementKind.TYPE_END:
                in_type = False
            elif in_type and kind is StatementKind.OTHER and statement.tokens:
                first = statement.tokens[0]
                if tokens[first].kind is TokenKind.IDENTIFIER:
                    decl.members.add(first)
                    decl.anywhere[tokens[first].key] = tokens[first].text
                    decl.member_names[tokens[first].key] = tokens[first].text
            elif enum_public is not None and kind is StatementKind.OTHER and statement.tokens:
                first = statement.tokens[0]
                if tokens[first].kind is TokenKind.IDENTIFIER:
                    add_module(first, enum_public, member=True)
    return decl


def _name_token(doc: Document, statement: Statement) -> int | None:
    if statement.name is None or statement.name >= len(statement.tokens):
        return None
    j = statement.tokens[statement.name]
    return j if doc.tokens[j].kind is TokenKind.IDENTIFIER else None


def _variable_names(doc: Document, statement: Statement) -> list[int]:
    """The declared names in a Dim/Private/Public/Static/Const/ReDim list."""
    tokens = doc.tokens
    names: list[int] = []
    depth = 0
    expect_name = True
    for j in statement.tokens:
        tok = tokens[j]
        if tok.kind is TokenKind.PUNCTUATION:
            if tok.text == "(":
                depth += 1
            elif tok.text == ")":
                depth -= 1
            elif tok.text == "," and depth == 0:
                expect_name = True
            continue
        if not expect_name or depth != 0:
            continue
        if tok.is_word and tok.lower in _LEADING_DECLARATION_WORDS:
            continue
        if tok.kind is TokenKind.IDENTIFIER:
            names.append(j)
        expect_name = False
    return names


def _parameter_names(doc: Document, statement: Statement) -> list[int]:
    """The parameter names of a procedure declaration."""
    tokens = doc.tokens
    indices = statement.tokens
    start = statement.name if statement.name is not None else 0
    open_at = None
    for k in range(start, len(indices)):
        tok = tokens[indices[k]]
        if tok.kind is TokenKind.PUNCTUATION and tok.text == "(":
            open_at = k
            break
    if open_at is None:
        return []
    names: list[int] = []
    depth = 0
    expect_name = True
    for k in range(open_at, len(indices)):
        tok = tokens[indices[k]]
        if tok.kind is TokenKind.PUNCTUATION:
            if tok.text == "(":
                depth += 1
                continue
            if tok.text == ")":
                depth -= 1
                if depth == 0:
                    break
                continue
            if tok.text == "," and depth == 1:
                expect_name = True
                continue
        if depth != 1 or not expect_name:
            continue
        if tok.is_word and tok.lower in _PARAMETER_WORDS:
            continue
        if tok.kind is TokenKind.IDENTIFIER:
            names.append(indices[k])
        expect_name = False
    return names
