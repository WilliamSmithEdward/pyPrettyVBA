"""Hypothesis strategies that write VBA modules.

Two kinds:

* ``token_soup``: any sequence of VBA fragments (keywords, names, literals,
  operators, comments, continuations, directives, line breaks), well-formed
  or not. For properties that must hold for every input.
* ``canonical_modules``: a well-formed module in the exact form the default
  preset produces, paired with a copy scrambled only in ways the default
  preset repairs (indentation, letter case of keywords and uses of names,
  whitespace inside lines, blank lines, trailing whitespace, final line
  breaks). Formatting the scrambled copy must give back the canonical text
  byte for byte.

The canonical form is written out by hand here, independently of the
formatter, so a disagreement is a finding on one side or the other.
"""

from __future__ import annotations

import random
from dataclasses import dataclass, field

from hypothesis import strategies as st

# ---------------------------------------------------------------------------
# Token soup

SOUP_FRAGMENTS = (
    # statements and keywords
    "Sub ", "Function ", "Property Get ", "End Sub", "End Function", "End Property", "End",
    "If ", " Then", " Then ", "Else", "ElseIf ", "End If", "EndIf", "For ", "Each ", " In ", " To ",
    " Step ", "Next", "Do", "Loop", "While ", "Until ", "Wend", "Select Case ", "Case ", "Case Else",
    "End Select", "With ", "End With", "Dim ", " As ", "Long", "String", "Variant", "Byte()", "Const ",
    "Private ", "Public ", "Static ", "Type ", "End Type", "Enum ", "End Enum", "Declare ", "PtrSafe ",
    "Lib ", "Alias ", "Call ", "Set ", "Let ", "Not ", " And ", " Or ", " Mod ", " Is ", " Like ",
    "New ", "Nothing", "True", "False", "Me", "Debug.Print ", "Debug.Assert ", "Option Explicit",
    "Option Compare Text", "GoTo ", "On Error ", "Resume Next", "Exit Sub", "Stop", "Open ", " For ",
    "Output", "Binary", "Access ", "Read", "Write ", "Close ", "Line Input ", "Get ", "Put ", "Input ",
    "Name ", "Mid", "LSet ", "RSet ", "Erase ", "ReDim ", "Preserve ", "Optional ", "ByVal ", "ByRef ",
    "ParamArray ", "DefLng ", "DefStr a-c", "Event ", "RaiseEvent ", "Implements ", "WithEvents ",
    "Friend ", "Print ", "?",
    # names
    "x", "y", "total", "Total", "TOTAL", "s$", "n%", "l&", "c@", "d#", "f!", "q^", "[a b]", "foo",
    "Foo", "msgbox", "MsgBox", "vbCrLf", "text", "Text", "ws", "rs!Field", "rs! Field", ".Count",
    ".Name", "Item", "Value", "ÄÖ", "СЧЁТ", "счёт", "Ωmega", "a_b", "Err.Description",
    # numbers
    "0", "1", "10", "10%", "32768", "32768&", "1.5", "1.", ".5", "1e3", "1E+03", "1.5D2", "&HFF",
    "&hff%", "&O17", "&17", "&H8000", "&H8000&", "1#", "1!", "1@", "1^", "2147483648",
    "1.0000000000000001", "123456789012345678", "-1", "- 1", "+2",
    # strings and dates
    '"a"', '"a""b"', '""', '"x  y"', '"unterminated', "#1/2/2003#", "#2003-01-02#", "#1:30 PM#",
    "#Jan 2, 2003#", "#1/2/03#", "#13/13/2003#", "#",
    # operators and punctuation
    "+", "-", "*", "/", "\\", "^", "&", "=", "<>", "<=", ">=", "=>", "=<", "><", "<", ">", ":=",
    ".", "!", "(", ")", ",", ";", ":", " = ", " & ", " , ", "( ", " )",
    # whitespace, comments, continuations, line breaks
    " ", "  ", "\t", " ", "　", " _\r\n", " _\n", " _   \r\n", "_", "\r\n", "\r\n", "\n",
    "\r", "' comment", "' note _", "Rem x", "rem ", "REM", "'@prettyvba-ignore",
    "'@prettyvba-ignore-next-line", "'@prettyvba-ignore-start", "'@prettyvba-ignore-end",
    "'@prettyvba-ignore-file", "'@prettyvba-ignore: spacing", "'@prettyvba-ignore: nosuch",
    # directives, labels, headers
    "#If ", "#Else", "#ElseIf ", "#End If", "#Const ", "VBA7", "Win64", "label:", "10 ", "20 x = 1",
    'Attribute VB_Name = "M"\r\n', "VERSION 1.0 CLASS\r\n", "BEGIN\r\n", "END\r\n",
    # characters with no rule
    "é", "\x00", "​", "﻿", "$", "%", "@", "`", "~", "{", "}", "|",
)

_CASES = (str, str.lower, str.upper, str.swapcase)


@st.composite
def token_soup(draw: st.DrawFn) -> str:
    parts = draw(
        st.lists(st.tuples(st.sampled_from(SOUP_FRAGMENTS), st.integers(0, len(_CASES) - 1)), max_size=60)
    )
    return "".join(_CASES[case](text) for text, case in parts)


# ---------------------------------------------------------------------------
# Canonical modules


@dataclass(frozen=True)
class Atom:
    """One token of a generated statement.

    ``cls`` decides the whitespace around it in the canonical form:
    word, fn (a name glued to the parenthesis after it), op (a comparison or
    arithmetic operator, safe to write without spaces), sep (an operator that
    must keep its spaces: `&` can be a type character, `-` a sign), unary,
    open, close, comma, dot (member access) and wdot (a With member).
    """

    text: str
    cls: str = "word"
    fixed_case: bool = False


def W(text: str) -> Atom:
    """A keyword, or a use of a name: its case is scrambled and must come back."""
    return Atom(text, "word")


def D(text: str, cls: str = "word") -> Atom:
    """A declaration, or a member name: its spelling is the canonical one."""
    return Atom(text, cls, fixed_case=True)


def FN(text: str) -> Atom:
    return Atom(text, "fn")


OPEN = Atom("(", "open")
CLOSE = Atom(")", "close")
COMMA = Atom(",", "comma")
DOT = Atom(".", "dot")
WDOT = Atom(".", "wdot")


def OP(text: str) -> Atom:
    return Atom(text, "op")


def SEP(text: str) -> Atom:
    return Atom(text, "sep")


def LIT(text: str) -> Atom:
    return Atom(text, "lit", fixed_case=True)


def NUM(text: str) -> Atom:
    return Atom(text, "num", fixed_case=True)


# Other spellings of a number that format to the canonical one.
NUMBER_SPELLINGS = {
    "1": ("1", "1%"),
    "2": ("2", "2%"),
    "3": ("3",),
    "10": ("10", "10%"),
    "255": ("255",),
    "&HFF": ("&HFF", "&hff", "&HFF%", "&H0FF", "&h00ff"),
    "&O377": ("&O377", "&o377", "&377", "&O0377"),
    "1.5": ("1.5", "1.50", "1.5#", "15E-1"),
}


def canonical_gap(a: Atom, b: Atom) -> str:
    if b.cls in ("close", "comma", "dot"):
        return ""
    if a.cls in ("open", "dot", "wdot", "unary"):
        return ""
    if b.cls == "open":
        return "" if a.cls == "fn" else " "
    return " "


_BLANK_FILL = ("", "", " ", "\t", "    ")


@dataclass
class GenLine:
    """A generated line, or a run of blank lines (depth None). Depth -1 is column 1."""

    depth: int | None
    atoms: tuple[Atom, ...] = ()
    # Gap indices (the gap before atoms[i]) whose width is kept as written.
    fixed_gaps: frozenset[int] = frozenset()
    comment: str | None = None
    # Blank lines: how many the canonical form has, and the range the
    # scrambled copy may hold instead.
    blanks: int = 0
    blank_range: tuple[int, int] = (0, 0)
    # Blank lines the scrambled copy may add after this line (a procedure
    # header) or before it (a procedure's End), which formatting removes.
    pad_after: bool = False
    pad_before: bool = False

    def canonical(self) -> list[str]:
        if self.depth is None:
            return [""] * self.blanks
        indent = "" if self.depth < 0 else "    " * self.depth
        if self.comment is not None:
            return [indent + self.comment]
        parts = [self.atoms[0].text]
        for i in range(1, len(self.atoms)):
            parts.append(canonical_gap(self.atoms[i - 1], self.atoms[i]))
            parts.append(self.atoms[i].text)
        return [indent + "".join(parts)]

    def scrambled(self, rnd: random.Random) -> list[str]:
        if self.depth is None:
            low, high = self.blank_range
            return [rnd.choice(_BLANK_FILL) for _ in range(rnd.randint(low, high))]
        before = [rnd.choice(_BLANK_FILL) for _ in range(rnd.randint(0, 2))] if self.pad_before else []
        after = [rnd.choice(_BLANK_FILL) for _ in range(rnd.randint(0, 2))] if self.pad_after else []
        indent = rnd.choice(("", " ", "  ", "    ", "\t", "        ", " \t", "\t\t"))
        tail = rnd.choice(("", "", "", " ", "\t", "   "))
        if self.comment is not None:
            return [*before, indent + self.comment + tail, *after]
        parts = [_scramble_atom(self.atoms[0], rnd)]
        for i in range(1, len(self.atoms)):
            a, b = self.atoms[i - 1], self.atoms[i]
            gap = canonical_gap(a, b)
            if i not in self.fixed_gaps:
                gap = _scramble_gap(a, b, gap, rnd)
            parts.append(gap)
            parts.append(_scramble_atom(b, rnd))
        return [*before, indent + "".join(parts) + tail, *after]


def _scramble_atom(atom: Atom, rnd: random.Random) -> str:
    if atom.cls == "num":
        return rnd.choice(NUMBER_SPELLINGS.get(atom.text, (atom.text,)))
    if atom.fixed_case or not any(c.isalpha() for c in atom.text):
        return atom.text
    return rnd.choice(_CASES)(atom.text)


def _scramble_gap(a: Atom, b: Atom, gap: str, rnd: random.Random) -> str:
    if gap == "":
        if a.cls == "open" or b.cls in ("close", "comma"):
            return rnd.choice(("", "", " ", "  "))
        return ""
    if a.cls == "comma":
        return rnd.choice(("", " ", " ", "  ", "\t"))
    if (a.cls == "op" and b.cls != "op") or (b.cls == "op" and a.cls != "op"):
        return rnd.choice(("", " ", " ", "  ", "\t"))
    return rnd.choice((" ", " ", "  ", "\t", "   "))


LONGS = ("alpha", "beta", "rowIndex")
STRINGS = ("text", "message")
MAX_DEPTH = 3


@dataclass
class _Procedure:
    kind: str  # "Sub" or "Function"
    name: str
    params: tuple[str, ...] = ()
    handler: bool = False


@dataclass
class _Builder:
    draw: st.DrawFn
    lines: list[GenLine] = field(default_factory=list)
    procedures: list[_Procedure] = field(default_factory=list)

    def add(self, depth: int, *atoms: Atom, fixed: tuple[int, ...] = ()) -> GenLine:
        line = GenLine(depth, tuple(atoms), frozenset(fixed))
        self.lines.append(line)
        return line

    def blank(self, count: int, scrambled: tuple[int, int] | None = None) -> None:
        self.lines.append(GenLine(None, blanks=count, blank_range=scrambled or (count, count)))

    def pick(self, options: tuple[str, ...]) -> str:
        return self.draw(st.sampled_from(options))

    # -- expressions --------------------------------------------------------

    def number(self) -> Atom:
        return NUM(self.pick(tuple(NUMBER_SPELLINGS)))

    def operand(self, callables: list[str]) -> list[Atom]:
        choice = self.draw(st.integers(0, 5))
        if choice == 0:
            return [self.number()]
        if choice == 1:
            return [W("LIMIT")]
        if choice == 2:
            return [W(self.pick(("Light", "Dark")))]
        if choice == 3 and callables:
            return [FN(self.pick(tuple(callables))), OPEN, W(self.pick(LONGS)), CLOSE]
        if choice == 4:
            return [FN("Len"), OPEN, W(self.pick(STRINGS)), CLOSE]
        return [W(self.pick(LONGS))]

    def expression(self, callables: list[str]) -> list[Atom]:
        atoms = self.operand(callables)
        for _ in range(self.draw(st.integers(0, 2))):
            op = self.pick(("+", "*", "-", "Mod", "\\"))
            if op == "Mod":
                atoms.append(W("Mod"))
            elif op in ("-", "\\"):
                atoms.append(SEP(op))
            else:
                atoms.append(OP(op))
            atoms.extend(self.operand(callables))
        return atoms

    def condition(self, callables: list[str]) -> list[Atom]:
        atoms = [W(self.pick(LONGS)), OP(self.pick(("<", ">", "<=", ">=", "=", "<>")))]
        atoms.extend(self.operand(callables))
        if self.draw(st.booleans()):
            atoms.append(W(self.pick(("And", "Or"))))
            atoms.extend([W(self.pick(LONGS)), OP(">"), self.number()])
        return atoms

    # -- statements ---------------------------------------------------------

    def simple(self, depth: int, proc: _Procedure, in_with: bool) -> None:
        callables = [p.name for p in self.procedures if p.kind == "Function" and p.params]
        variable = W(self.pick(LONGS))
        choice = self.draw(st.integers(0, 13))
        if choice == 0:
            self.add(depth, variable, OP("="), *self.expression(callables))
        elif choice == 1:
            text = W(self.pick(STRINGS))
            self.add(depth, text, OP("="), W(text.text), SEP("&"), LIT('"x"'))
        elif choice == 2:
            self.add(depth, W("Debug"), DOT, W("Print"), *self.expression(callables))
        elif choice == 3:
            self.add(depth, W("MsgBox"), LIT('"Done"'), COMMA, W("vbInformation"))
        elif choice == 4:
            self.add(depth, W("Call"), FN("Helper"), OPEN, *self.expression(callables), COMMA, self.number(), CLOSE)
        elif choice == 5:
            self.add(depth, W("Helper"), *self.expression(callables), COMMA, self.number())
        elif choice == 6:
            self.add(
                depth, W("If"), *self.condition(callables), W("Then"), variable, OP("="), *self.operand(callables)
            )
        elif choice == 7:
            self.add(depth, W("Set"), W("items"), OP("="), W("New"), W("Collection"))
        elif choice == 8:
            self.add(depth, W("items"), DOT, D("Add"), *self.operand(callables))
        elif choice == 9:
            self.add(depth, W("counter"), OP("="), W("counter"), OP("+"), NUM("1"))
        elif choice == 10:
            self.add(depth, variable, OP("="), Atom("-", "unary"), W(self.pick(LONGS)))
        elif choice == 11 and in_with:
            self.add(depth, WDOT, D("Add"), *self.operand(callables))
        elif choice == 12 and in_with:
            self.add(depth, variable, OP("="), WDOT, D("Count"))
        elif choice == 13 and proc.kind == "Function":
            self.add(depth, W(proc.name), OP("="), *self.expression(callables))
        else:
            self.add(depth, variable, OP("="), variable, OP("+"), NUM("1"))

    def block(self, depth: int, proc: _Procedure, in_with: bool) -> None:
        callables = [p.name for p in self.procedures if p.kind == "Function" and p.params]
        choice = self.draw(st.integers(0, 8))
        if choice == 0:
            self.add(depth, W("If"), *self.condition(callables), W("Then"))
            self.statements(depth + 1, proc, in_with)
            if self.draw(st.booleans()):
                self.add(depth, W("ElseIf"), *self.condition(callables), W("Then"))
                self.statements(depth + 1, proc, in_with)
            if self.draw(st.booleans()):
                self.add(depth, W("Else"))
                self.statements(depth + 1, proc, in_with)
            self.add(depth, W("End"), W("If"))
        elif choice == 1:
            counter = self.pick(LONGS)
            step = [W("Step"), Atom("-", "unary"), NUM("1")] if self.draw(st.booleans()) else []
            self.add(depth, W("For"), W(counter), OP("="), NUM("1"), W("To"), NUM("10"), *step)
            self.statements(depth + 1, proc, in_with)
            self.add(depth, W("Next"), *([W(counter)] if self.draw(st.booleans()) else []))
        elif choice == 2:
            self.add(depth, W("For"), W("Each"), W("entry"), W("In"), W("items"))
            self.statements(depth + 1, proc, in_with)
            self.add(depth, W("Next"), W("entry"))
        elif choice == 3:
            self.add(depth, W("Do"), W(self.pick(("While", "Until"))), *self.condition(callables))
            self.statements(depth + 1, proc, in_with)
            self.add(depth, W("Loop"))
        elif choice == 4:
            self.add(depth, W("Do"))
            self.statements(depth + 1, proc, in_with)
            self.add(depth, W("Loop"), W(self.pick(("While", "Until"))), *self.condition(callables))
        elif choice == 5:
            self.add(depth, W("While"), *self.condition(callables))
            self.statements(depth + 1, proc, in_with)
            self.add(depth, W("Wend"))
        elif choice == 6:
            self.add(depth, W("Select"), W("Case"), W(self.pick(LONGS)))
            self.add(depth + 1, W("Case"), NUM("1"))
            self.statements(depth + 2, proc, in_with)
            if self.draw(st.booleans()):
                self.add(depth + 1, W("Case"), NUM("2"), COMMA, NUM("3"))
                self.statements(depth + 2, proc, in_with)
            if self.draw(st.booleans()):
                self.add(depth + 1, W("Case"), W("Else"))
                self.statements(depth + 2, proc, in_with)
            self.add(depth, W("End"), W("Select"))
        elif choice == 7:
            self.add(depth, W("With"), W("items"))
            self.statements(depth + 1, proc, True)
            self.add(depth, W("End"), W("With"))
        else:
            # A balanced #If block: its arms leave the block structure as
            # they found it, so their content is indented one level.
            self.add(depth, W("#If"), W("VBA7"), W("Then"))
            self.statements(depth + 1, proc, in_with)
            if self.draw(st.booleans()):
                self.add(depth, W("#Else"))
                self.statements(depth + 1, proc, in_with)
            self.add(depth, W("#End"), W("If"))

    def statements(self, depth: int, proc: _Procedure, in_with: bool) -> None:
        count = self.draw(st.integers(1, 3))
        for i in range(count):
            if i > 0 and self.draw(st.integers(0, 7)) == 0:
                wanted = self.draw(st.integers(1, 2))
                # More than two blank lines in a row come down to two.
                self.blank(wanted, (2, 4) if wanted == 2 else (1, 1))
            if self.draw(st.integers(0, 7)) == 0:
                self.lines.append(GenLine(depth, comment="' " + self.pick(("Note.", "Step two.", "TODO"))))
            if depth < MAX_DEPTH and self.draw(st.integers(0, 2)) == 0:
                self.block(depth, proc, in_with)
            else:
                self.simple(depth, proc, in_with)

    def procedure(self, proc: _Procedure) -> None:
        scope = self.pick(("Public ", "Private ", ""))
        head: list[Atom] = [W(scope.strip())] if scope else []
        head += [W(proc.kind), D(proc.name, "fn"), OPEN]
        for i, param in enumerate(proc.params):
            if i:
                head.append(COMMA)
            head += [W("ByVal"), D(param), W("As"), W("Long")]
        head.append(CLOSE)
        if proc.kind == "Function":
            head += [W("As"), W("Long")]
        if self.draw(st.integers(0, 3)) == 0:
            self.lines.append(GenLine(0, comment="' " + proc.name + " does its part."))
        self.add(0, *head).pad_after = True
        self.add(1, W("Dim"), D("alpha"), W("As"), W("Long"), COMMA, D("beta"), W("As"), W("Long"), fixed=(2, 6))
        self.add(1, W("Dim"), D("rowIndex"), W("As"), W("Long"), fixed=(2,))
        self.add(1, W("Dim"), D("text"), W("As"), W("String"), COMMA, D("message"), W("As"), W("String"),
                 fixed=(2, 6))
        self.add(1, W("Dim"), D("items"), W("As"), W("Collection"), fixed=(2,))
        self.add(1, W("Dim"), D("entry"), W("As"), W("Variant"), fixed=(2,))
        if proc.handler:
            self.add(1, W("On"), W("Error"), W("GoTo"), W("fail"))
        self.statements(1, proc, False)
        if proc.handler:
            self.add(1, W("Exit"), W(proc.kind))
            self.add(-1, D("fail:"))
            self.add(1, W("MsgBox"), W("Err"), DOT, W("Description"))
        self.add(0, W("End"), W(proc.kind)).pad_before = True


@dataclass
class GeneratedModule:
    canonical: str
    scrambled: str


@st.composite
def canonical_modules(draw: st.DrawFn) -> GeneratedModule:
    b = _Builder(draw)
    header = 'Attribute VB_Name = "Generated"\r\n' if draw(st.booleans()) else ""
    b.add(0, W("Option"), W("Explicit"))
    b.blank(1)
    b.add(0, W("Private"), D("counter"), W("As"), W("Long"), fixed=(2,))
    b.add(0, W("Public"), W("Const"), D("LIMIT"), W("As"), W("Long"), OP("="), NUM("10"), fixed=(3,))
    b.blank(1)
    b.add(0, W("Public"), W("Enum"), D("Shade"))
    b.add(1, D("Light"), OP("="), NUM("1"))
    b.add(1, D("Dark"))
    b.add(0, W("End"), W("Enum"))
    b.blank(1)
    b.add(0, W("Private"), W("Type"), D("Pair"))
    b.add(1, D("lower"), W("As"), W("Long"), fixed=(1,))
    b.add(1, D("upper"), W("As"), W("Long"), fixed=(1,))
    b.add(0, W("End"), W("Type"))
    helper = _Procedure("Sub", "Helper", ("lhs", "rhs"))
    b.procedures.append(helper)
    count = draw(st.integers(1, 3))
    for i in range(count):
        kind = draw(st.sampled_from(("Sub", "Function")))
        params = ("lhs",) if kind == "Function" else ()
        b.procedures.append(_Procedure(kind, f"{kind[0]}{'ub' if kind == 'Sub' else 'unc'}Part{i}", params,
                                       draw(st.booleans())))
    for proc in b.procedures:
        # One blank line before each procedure, however many were written.
        b.blank(1, (0, 3))
        b.procedure(proc)
    rnd = draw(st.randoms(use_true_random=False))
    canonical = [text for line in b.lines for text in line.canonical()]
    scrambled = [text for line in b.lines for text in line.scrambled(rnd)]
    ending = rnd.choice(("", "\r\n", "\r\n\r\n", "\r\n \r\n"))
    return GeneratedModule(
        header + "\r\n".join(canonical) + "\r\n",
        header + "\r\n".join(scrambled) + ending,
    )
