"""Letter case: keywords the way the VBE spells them, names the way they are declared.

VBA ignores letter case, and the VBE rewrites it: every keyword takes one
fixed spelling, and every name takes the spelling of its declaration (or,
for a built-in, of its library). Code kept outside the VBE drifts from those
spellings, and every import then shows up as a diff. These rules put the
spellings back, which never changes what the code does, with one exception
the rules respect: the name of a Declare without an Alias is also the DLL
entry point, which is looked up case-sensitively, so it is never respelled.
"""

from __future__ import annotations

from collections.abc import Iterable

from ..contextual import contextual_keywords
from ..document import Document, LineKind, StatementKind
from ..keywords import canonical_keyword
from ..lexer import TokenKind, name_key
from ..names import builtin_names, library_names, member_names
from ..scopes import Declarations, collect_declarations
from .base import Finding, Option, Rule

__all__ = ["IdentifierCaseRule", "KeywordCaseRule"]


class KeywordCaseRule(Rule):
    """Spell every keyword the way the VBE does.

    Covers reserved words (`dim` -> `Dim`, `byval` -> `ByVal`), contextual
    keywords in the statement that makes them keywords (`Option Explicit`,
    `Declare PtrSafe ... Lib ... Alias`, `For ... Step`, `On Error`), the
    `Rem` of a Rem comment, the letters of a `DefInt A-Z` range, and the
    one-word `EndIf`, which the VBE writes `End If`. A keyword used as a
    member name (`wb.Close`, `.Print`) is a name, not a keyword: this rule
    leaves it to identifier-case, except `Debug.Print`, which the VBE
    capitalizes.
    """

    code = "keyword-case"
    summary = "Spell keywords the way the VBE does."
    category = "casing"
    vbe_canonical = True
    options = (
        Option("expand-endif", True, "Write the one-word `EndIf` as `End If`, as the VBE does."),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        tokens = doc.tokens
        contextual = contextual_keywords(doc)
        expand = bool(self.settings["expand-endif"])
        for line in doc.lines:
            if line.kind in (LineKind.BLANK, LineKind.ATTRIBUTE):
                continue
            deftype = any(s.kind is StatementKind.DEFTYPE for s in line.statements)
            for j in range(line.first, line.stop):
                tok = tokens[j]
                if tok.kind is TokenKind.KEYWORD:
                    if doc.is_member_name(j) and not _debug_member(doc, j):
                        continue
                    if tok.lower == "endif" and expand:
                        wanted = "End If"
                    else:
                        wanted = canonical_keyword(tok.text) or tok.text
                elif tok.kind is TokenKind.IDENTIFIER:
                    if j in contextual:
                        wanted = contextual[j]
                    elif deftype and len(tok.text) == 1 and tok.text.isascii():
                        # Letter ranges are A-Z; other letters are not case-folded.
                        wanted = tok.text.upper()
                    else:
                        continue
                elif tok.kind is TokenKind.COMMENT and tok.text[:3].lower() == "rem" and tok.text[:3] != "Rem":
                    yield Finding(tok.start, tok.start + 3, "Rem", f"Spell {tok.text[:3]!r} as 'Rem'.")
                    continue
                else:
                    continue
                if wanted != tok.text:
                    yield Finding(tok.start, tok.end, wanted, f"Spell {tok.text!r} as {wanted!r}.")


def _debug_member(doc: Document, j: int) -> bool:
    """True for the `Print` of `Debug.Print`."""
    dot = doc.prev_code(j)
    obj = doc.prev_code(dot) if dot is not None else None
    return obj is not None and doc.tokens[obj].lower == "debug"


class IdentifierCaseRule(Rule):
    """Spell each name the way its declaration spells it.

    A name declared in the module takes the spelling of its declaration: a
    procedure's parameters, locals and labels inside that procedure, the
    module's variables, constants, procedures, Types, Enums and Enum members
    everywhere else. A name declared in another module of the project takes
    that module's spelling when the project is formatted together. A name
    from the VBA library or an Office host (`msgbox`, `vbcrlf`, `range`) takes
    the library's spelling.

    Names after `.` or `!` (members) and before `:=` (named arguments) are
    left alone by default. The members of `Debug` and `Err` are the
    exception, since their type is fixed.

    The VBE itself keeps one spelling per name for the whole project, and
    fills that table from the libraries the project references as well as
    from its declarations: every name those libraries define, member names
    included, so an undeclared `sql` is written `Sql` (Excel has a
    `QueryTable.Sql`) and `ws.cells(1, 1).value` is written
    `ws.Cells(1, 1).Value`. `library = "all"` and `members = "respell"`
    reproduce that, as the `vbe` preset does. A name declared in the module
    still wins over the libraries (measured: a Type member `second` makes
    the VBE write `second(Now)`).
    """

    code = "identifier-case"
    summary = "Spell names the way they are declared."
    category = "casing"
    vbe_canonical = True
    options = (
        Option(
            "builtins",
            True,
            "Respell names from the VBA library and the configured host libraries "
            "(the `hosts` setting) the way the library does.",
        ),
        Option(
            "scope",
            "procedure",
            "`procedure` gives a procedure's own parameters and locals their "
            "spelling inside that procedure only, as XLIDE does; `module` spells a "
            "name the same way everywhere in the module, after its last "
            "declaration, which is what the VBE does.",
            choices=("procedure", "module"),
        ),
        Option(
            "undeclared",
            "keep",
            "Names nothing declares: `keep` leaves each as written; `last` spells "
            "them all the way the last one is written, which is what the VBE does.",
            choices=("keep", "last"),
        ),
        Option(
            "library",
            "all",
            "Which library names count: `all`, every name the libraries define, "
            "members of every type included, which is what the VBE uses; "
            "`globals`, only the names code can use without a qualifier "
            "(functions, constants, objects).",
            choices=("all", "globals"),
        ),
        Option(
            "members",
            "respell",
            "Member names (after `.` or `!`) and named arguments (before `:=`): "
            "`respell` spells them like any other name, the way the VBE does; "
            "`keep` leaves them as written, apart from the members of Debug, Err "
            "and the libraries themselves, as XLIDE does. The VBE spells a member "
            "the way any declaration in the project spells that name, so when a "
            "class of your own declares a member (`Public count`), format the "
            "project together to keep its spelling.",
            choices=("respell", "keep"),
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        tokens = doc.tokens
        decl = collect_declarations(doc, self.context.module_kind)
        contextual = contextual_keywords(doc)
        hosts = tuple(self.context.hosts)
        builtins = bool(self.settings["builtins"])
        library: dict[str, str] = {}
        if builtins:
            library = library_names(hosts) if self.settings["library"] == "all" else builtin_names(hosts)
        unify = self.settings["undeclared"] == "last"
        module_scope = self.settings["scope"] == "module"
        respell_members = self.settings["members"] == "respell"
        # Undeclared names, for `undeclared = "last"`: token index and name.
        loose: list[tuple[int, str]] = []
        for line in doc.lines:
            if line.kind in (LineKind.BLANK, LineKind.COMMENT, LineKind.ATTRIBUTE):
                continue
            if any(s.kind is StatementKind.DEFTYPE for s in line.statements):
                continue
            for j in range(line.first, line.stop):
                tok = tokens[j]
                if j in contextual or j in decl.pinned:
                    continue
                # A keyword after `.` is a member name (`wb.Close`, `.End(xlUp)`),
                # which `members = "respell"` spells from the libraries too.
                member_keyword = (
                    tok.kind is TokenKind.KEYWORD and respell_members and doc.is_member_name(j)
                )
                if tok.kind is not TokenKind.IDENTIFIER and not member_keyword:
                    continue
                nxt = doc.next_code(j)
                named = nxt is not None and tokens[nxt].kind is TokenKind.OPERATOR and tokens[nxt].text == ":="
                lower = tok.key
                if (named or doc.is_member_name(j)) and not respell_members:
                    if named or not builtins:
                        continue
                    wanted, source = self._builtin_member(doc, j, lower), "as its library does"
                elif j in decl.members and not module_scope:
                    # A Type member's own declaration: not a name in scope
                    # here, and never respelled after a library name.
                    continue
                elif (named or doc.is_member_name(j)) and not module_scope:
                    wanted, source = self._member_spelling(lower, decl, library, named)
                else:
                    wanted, source = self._spelling(lower, line.index, decl, library, module_scope)
                    if wanted == "":
                        # A library spells it two ways; which one the VBE uses is unknown.
                        continue
                    if wanted is None and unify and not member_keyword:
                        loose.append((j, lower))
                if wanted is not None and wanted != tok.text and name_key(wanted) == lower:
                    yield Finding(tok.start, tok.end, wanted, f"Spell {tok.text!r} {source}: {wanted!r}.")
        if loose:
            last: dict[str, str] = {}
            for j, lower in loose:
                last[lower] = tokens[j].text
            for j, lower in loose:
                tok = tokens[j]
                if tok.text != last[lower]:
                    yield Finding(
                        tok.start, tok.end, last[lower],
                        f"Spell {tok.text!r} as its last use does: {last[lower]!r}.",
                    )

    def _spelling(
        self, lower: str, line_index: int, decl: Declarations, library: dict[str, str], wide: bool
    ) -> tuple[str | None, str]:
        """The spelling a name should have, and where it comes from.

        Declarations first, then the other modules of the project, then the
        libraries; None when nothing spells it. ``wide`` looks at every
        declaration of the module and the project, as the VBE's project-wide
        name table does; otherwise only at what is in scope at the line.
        """
        declared = decl.anywhere.get(lower) if wide else decl.spelling(lower, line_index)
        if declared is not None:
            return declared, "as declared"
        # The VBE's name table is the whole project's; a procedure's scope
        # sees only what other modules make public.
        project = self.context.project_declared if wide else self.context.project_names
        if lower in project:
            return project[lower], "as the module that declares it does"
        return library.get(lower), "as its library does"

    def _member_spelling(
        self, lower: str, decl: Declarations, library: dict[str, str], named: bool
    ) -> tuple[str | None, str]:
        """The spelling of a member name, or of a named argument, in procedure scope.

        Only declarations that can name a member count: Type and Enum
        members, what the module and the project's classes do not declare
        Private, and, for `name:=`, the module's parameters. A local
        variable named like a member does not respell it, which the VBE,
        with its one table of names, would do.
        """
        declared = decl.member_names.get(lower)
        if declared is None and named:
            declared = decl.parameters.get(lower)
        if declared is not None:
            return declared, "as declared"
        project = self.context.project_members.get(lower) or self.context.project_names.get(lower)
        if project is not None:
            return project, "as the module that declares it does"
        return library.get(lower), "as its library does"

    @staticmethod
    def _builtin_member(doc: Document, j: int, lower: str) -> str | None:
        dot = doc.prev_code(j)
        obj = doc.prev_code(dot) if dot is not None else None
        if obj is None or doc.tokens[obj].kind not in (TokenKind.IDENTIFIER, TokenKind.KEYWORD):
            return None
        if doc.is_member_name(obj):
            return None
        return member_names(doc.tokens[obj].lower).get(lower)
