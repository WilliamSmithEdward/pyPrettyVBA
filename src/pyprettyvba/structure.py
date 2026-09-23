"""Block structure: how deep each line sits, and where the structure breaks.

``analyze_structure`` walks a module's logical lines with a stack of open
blocks and gives every line a nesting level. Blocks are procedures, Type and
Enum declarations, block If (with its ElseIf and Else arms), Select Case
(with its Case arms), For, Do, While and With. A line's level is the sum of
the weights of the blocks open around it; the weights come from the indent
rule's options, so `case = false` makes Select Case weigh nothing.

Conditional compilation is read the way the VBE reads it: only one arm of
an `#If` is ever compiled, so each arm starts from the stack the `#If` saw,
and after `#End If` the stack is what the first arm left. That keeps

    #If VBA7 Then
    Private Sub Tick(ByVal ms As LongPtr)
    #Else
    Private Sub Tick(ByVal ms As Long)
    #End If
        ...
    End Sub

one procedure rather than two. An `#If` block whose arms all leave the stack
as they found it is balanced, and only a balanced block's content takes the
extra level the directive arms get.

A closer with no opener, an opener never closed, or a closer of the wrong
kind leaves the structure of the region around it unknown: its procedure,
or the module-level lines between two procedures. Such regions are listed
in ``broken`` with the reason in ``problems``, and the indent rule leaves
them exactly as written.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .document import Document, LineKind, StatementKind

__all__ = ["Structure", "analyze_structure"]


@dataclass
class _Block:
    kind: str
    weight: int
    line: int


@dataclass
class _CcFrame:
    line: int
    snapshot: list[_Block]
    # The level of the #If line; its #Else and #End If lines share it.
    level: int = 0
    first_arm_end: list[_Block] | None = None
    balanced: bool = True


@dataclass
class Structure:
    """Per-line nesting levels and the regions whose structure is broken."""

    # Level of each logical line. A line holding only a comment gets the
    # level a statement would get in its place.
    levels: list[int] = field(default_factory=list)
    # (first line, last line) of regions whose structure could not be read.
    broken: list[tuple[int, int]] = field(default_factory=list)
    # (line index, message) for each structural problem found.
    problems: list[tuple[int, str]] = field(default_factory=list)
    # (first line, last line) of each procedure, header to End line.
    procedures: list[tuple[int, int]] = field(default_factory=list)

    def is_broken(self, line_index: int) -> bool:
        return any(first <= line_index <= last for first, last in self.broken)


_OPENER_KIND = {
    StatementKind.IF_BLOCK: "if",
    StatementKind.SELECT: "select",
    StatementKind.FOR: "for",
    StatementKind.DO: "do",
    StatementKind.WHILE: "while",
    StatementKind.WITH: "with",
    StatementKind.TYPE_START: "type",
    StatementKind.ENUM_START: "enum",
}
_CLOSER_KIND = {
    StatementKind.END_IF: "if",
    StatementKind.END_SELECT: "select",
    StatementKind.NEXT: "for",
    StatementKind.LOOP: "do",
    StatementKind.WEND: "while",
    StatementKind.END_WITH: "with",
    StatementKind.TYPE_END: "type",
    StatementKind.ENUM_END: "enum",
}
_CLOSER_TEXT = {
    "if": "End If",
    "select": "End Select",
    "case": "End Select",
    "for": "Next",
    "do": "Loop",
    "while": "Wend",
    "with": "End With",
    "type": "End Type",
    "enum": "End Enum",
    "proc": "End Sub, End Function or End Property",
}


def analyze_structure(
    doc: Document,
    *,
    procedure_body: bool = True,
    case_arms: bool = True,
    type_members: bool = True,
    directive_arms: bool = True,
) -> Structure:
    """Nesting levels for every logical line of ``doc``."""
    weights = {
        "proc": 1 if procedure_body else 0,
        "select": 1 if case_arms else 0,
        "type": 1 if type_members else 0,
        "enum": 1 if type_members else 0,
    }
    if not directive_arms:
        return _Walker(doc, weights, set()).walk()
    # Which #If blocks are balanced is only known at their #End If, so a
    # first walk finds them and a second one gives their content its level.
    first = _Walker(doc, weights, set())
    first.walk()
    return _Walker(doc, weights, first.balanced_cc).walk()


class _Walker:
    def __init__(self, doc: Document, weights: dict[str, int], weighted_cc: set[int]) -> None:
        self.doc = doc
        self.weights = weights
        self.weighted_cc = weighted_cc
        self.stack: list[_Block] = []
        self.cc: list[_CcFrame] = []
        self.cc_weight = 0
        self.cc_weights: list[int] = []
        self.balanced_cc: set[int] = set()
        self.result = Structure()
        self.region_start = 0
        self.region_broken = False
        self.proc_start: int | None = None

    # -- helpers ---------------------------------------------------------------

    def level(self) -> int:
        return sum(block.weight for block in self.stack) + self.cc_weight

    def push(self, kind: str, line: int) -> None:
        self.stack.append(_Block(kind, self.weights.get(kind, 1), line))

    def problem(self, line: int, message: str) -> None:
        self.result.problems.append((line, message))
        self.region_broken = True

    def unclosed(self, block: _Block) -> None:
        self.problem(block.line, f"This block is never closed: expected {_CLOSER_TEXT[block.kind]}.")

    def close_region(self, last_line: int) -> None:
        if self.region_broken and last_line >= self.region_start:
            self.result.broken.append((self.region_start, last_line))
        self.region_broken = False

    def proc_floor(self) -> int:
        """Index of the innermost open procedure, or -1."""
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i].kind == "proc":
                return i
        return -1

    def find(self, kind: str) -> int:
        """Index of the innermost open ``kind`` inside the current procedure."""
        lowest = 0 if kind == "proc" else self.proc_floor() + 1
        for i in range(len(self.stack) - 1, lowest - 1, -1):
            if self.stack[i].kind == kind:
                return i
        return -1

    # -- walking ---------------------------------------------------------------

    def walk(self) -> Structure:
        levels = self.result.levels
        for line in self.doc.lines:
            if line.kind is LineKind.DIRECTIVE:
                levels.append(self.directive(line.index, line.statements[0].kind))
            elif line.kind is LineKind.CODE:
                levels.append(self.code(line.index))
            else:
                levels.append(self.level())
        last = len(self.doc.lines) - 1
        for block in reversed(self.stack):
            self.unclosed(block)
            break
        if self.proc_start is not None:
            self.result.procedures.append((self.proc_start, last))
        self.close_region(last)
        return self.result

    def directive(self, index: int, kind: StatementKind) -> int:
        if kind is StatementKind.CC_IF:
            level = self.level()
            self.cc.append(_CcFrame(index, [_copy(b) for b in self.stack], level))
            self.cc_weights.append(self.cc_weight)
            if index in self.weighted_cc:
                self.cc_weight += 1
            return level
        if kind in (StatementKind.CC_ELSEIF, StatementKind.CC_ELSE):
            if not self.cc:
                return self.level()
            frame = self.cc[-1]
            if not _same(self.stack, frame.snapshot):
                frame.balanced = False
            if frame.first_arm_end is None:
                frame.first_arm_end = [_copy(b) for b in self.stack]
            self.stack = [_copy(b) for b in frame.snapshot]
            return frame.level
        if kind is StatementKind.CC_END_IF:
            if not self.cc:
                return self.level()
            frame = self.cc.pop()
            if not _same(self.stack, frame.snapshot):
                frame.balanced = False
            if frame.first_arm_end is not None:
                self.stack = frame.first_arm_end
            self.cc_weight = self.cc_weights.pop()
            if frame.balanced:
                self.balanced_cc.add(frame.line)
            return frame.level
        return self.level()

    def code(self, index: int) -> int:
        line = self.doc.lines[index]
        line_level: int | None = None
        for position, statement in enumerate(line.statements):
            kind = statement.kind
            first = position == 0
            before = self.level()
            if kind is StatementKind.PROC_START:
                self.start_procedure(index)
                before = self.level()
                self.push("proc", index)
            elif kind is StatementKind.PROC_END:
                self.end_procedure(index)
                before = self.level()
            elif kind in _CLOSER_KIND:
                count = statement.count if kind is StatementKind.NEXT else 1
                for _ in range(count):
                    self.close(_CLOSER_KIND[kind], index)
                before = self.level()
            elif kind in (StatementKind.ELSEIF, StatementKind.ELSE):
                at = self.find("if")
                if at < 0:
                    self.problem(index, "This Else or ElseIf has no block If.")
                else:
                    if at != len(self.stack) - 1:
                        self.unclosed(self.stack[-1])
                        self.stack = self.stack[: at + 1]
                    before = self.level() - self.stack[at].weight
            elif kind is StatementKind.CASE:
                at = self.find("select")
                if at < 0:
                    self.problem(index, "This Case has no Select Case.")
                else:
                    above = self.stack[at + 1 :]
                    if above and not (len(above) == 1 and above[0].kind == "case"):
                        self.unclosed(self.stack[-1])
                    self.stack = self.stack[: at + 1]
                    before = self.level()
                    self.push("case", index)
            elif kind in _OPENER_KIND:
                self.push(_OPENER_KIND[kind], index)
            if first:
                line_level = before
        return self.level() if line_level is None else line_level

    def start_procedure(self, index: int) -> None:
        floor = self.proc_floor()
        if floor >= 0:
            # The previous procedure never ended.
            self.unclosed(self.stack[-1] if self.stack[-1].kind != "proc" else self.stack[floor])
            if self.proc_start is not None:
                self.result.procedures.append((self.proc_start, index - 1))
            self.stack = self.stack[:floor]
        elif self.stack and self.stack[-1].kind in ("type", "enum"):
            self.unclosed(self.stack[-1])
            self.stack = [b for b in self.stack if b.kind not in ("type", "enum")]
        self.close_region(index - 1)
        self.region_start = index
        self.proc_start = index

    def end_procedure(self, index: int) -> None:
        at = self.find("proc")
        if at < 0:
            self.problem(index, "This End has no procedure to close.")
            self.close_region(index)
            self.region_start = index + 1
            return
        if at != len(self.stack) - 1:
            self.unclosed(self.stack[-1])
        self.stack = self.stack[:at]
        if self.proc_start is not None:
            self.result.procedures.append((self.proc_start, index))
            self.proc_start = None
        self.close_region(index)
        self.region_start = index + 1

    def close(self, kind: str, index: int) -> None:
        at = self.find(kind)
        if at < 0:
            self.problem(index, f"This {_CLOSER_TEXT[kind]} has nothing to close.")
            return
        above = [b for b in self.stack[at + 1 :] if not (kind == "select" and b.kind == "case")]
        if above:
            self.unclosed(above[-1])
        self.stack = self.stack[:at]


def _copy(block: _Block) -> _Block:
    return _Block(block.kind, block.weight, block.line)


def _same(a: list[_Block], b: list[_Block]) -> bool:
    return [(x.kind, x.line) for x in a] == [(y.kind, y.line) for y in b]
