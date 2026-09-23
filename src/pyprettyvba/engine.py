"""The formatting pipeline.

Rules run one after another over the module body, each on the text the
rules before it produced. A rule's findings are filtered through the
suppression directives of that text, its fixes applied, and every finding
mapped back to the line and column the user wrote, so a report names the
original source even when earlier rules moved things.

The pipeline repeats until a pass changes nothing, so the output is a fixed
point: formatting it again changes nothing. The rules are ordered so that
one pass is enough, and the second pass is the proof. Before returning, the
output is checked against the input with ``safety.first_difference``; output
that means something different is never returned.
"""

from __future__ import annotations

import bisect
from collections.abc import Sequence
from dataclasses import dataclass, field

from .document import Document, split_header
from .rules.base import Finding, FormatContext, Origin, Rule
from .safety import SafetyError, first_difference
from .suppression import DIRECTIVE_CODE, scan_suppressions

__all__ = ["FormatResult", "UnstableFormattingError", "Violation", "run_pipeline"]

MAX_PASSES = 4


class UnstableFormattingError(Exception):
    """The rules kept changing the text; a rule disagrees with another."""


@dataclass(frozen=True, order=True)
class Violation:
    """One thing a rule found, placed in the original file (1-based)."""

    line: int
    column: int
    end_line: int
    end_column: int
    rule: str
    message: str
    # True when formatting fixes it (the rule offered a fix and it applied).
    fixable: bool


@dataclass
class FormatResult:
    """The outcome of formatting one module's text."""

    source: str
    output: str
    violations: list[Violation] = field(default_factory=list)

    @property
    def changed(self) -> bool:
        return self.output != self.source


class _Stage:
    """The edits one rule made: maps offsets in its output back to its input."""

    def __init__(self, fixes: Sequence[Finding]) -> None:
        self.old_starts: list[int] = []
        self.old_ends: list[int] = []
        self.new_starts: list[int] = []
        self.new_ends: list[int] = []
        delta = 0
        for fix in fixes:
            assert fix.replacement is not None
            new_start = fix.start + delta
            self.old_starts.append(fix.start)
            self.old_ends.append(fix.end)
            self.new_starts.append(new_start)
            self.new_ends.append(new_start + len(fix.replacement))
            delta += len(fix.replacement) - (fix.end - fix.start)

    def back(self, pos: int) -> int:
        i = bisect.bisect_right(self.new_starts, pos) - 1
        if i < 0:
            return pos
        if pos < self.new_ends[i] or (pos == self.new_starts[i]):
            inside = pos - self.new_starts[i]
            return self.old_starts[i] + min(inside, self.old_ends[i] - self.old_starts[i])
        return self.old_ends[i] + (pos - self.new_ends[i])


def _apply(text: str, fixes: Sequence[Finding]) -> str:
    parts: list[str] = []
    cursor = 0
    for fix in fixes:
        assert fix.replacement is not None
        parts.append(text[cursor : fix.start])
        parts.append(fix.replacement)
        cursor = fix.end
    parts.append(text[cursor:])
    return "".join(parts)


def _non_overlapping(fixes: list[Finding]) -> list[Finding]:
    """Sort fixes and drop any that overlaps one kept before it."""
    fixes.sort(key=lambda f: (f.start, f.end))
    kept: list[Finding] = []
    last_end = -1
    for fix in fixes:
        if kept and fix.start < last_end:
            continue
        if kept and fix.start == last_end == fix.end and kept[-1].start == kept[-1].end:
            # Two insertions at one point: keep the first.
            continue
        kept.append(fix)
        last_end = fix.end
    return kept


@dataclass
class _Pending:
    rule: str
    message: str
    start: int
    end: int
    stage: int  # number of stages applied before the text this was found in
    fixable: bool
    group: str | None = None


def run_pipeline(
    text: str,
    rules: Sequence[Rule],
    context: FormatContext,
    known_codes: frozenset[str],
    *,
    check_safety: bool = True,
    header_newline: str | None = None,
) -> FormatResult:
    """Run ``rules`` over ``text`` and return the result.

    The export header (VERSION/BEGIN block, Attribute lines) is never
    formatted. ``header_newline``, when given, is the only change it gets:
    its line endings are rewritten to match the body's.
    """
    header, body = split_header(text)
    header_changed = False
    if header_newline is not None and header:
        converted = _convert_newlines(header, header_newline)
        header_changed = converted != header
        new_header = converted
    else:
        new_header = header
    stages: list[_Stage] = []
    pending: list[_Pending] = []
    current = body
    doc = Document(current)
    suppressions = scan_suppressions(doc, known_codes)

    def back(pos: int) -> int:
        return _to_original(pos, stages)

    for pass_number in range(MAX_PASSES):
        changed = False
        for rule in rules:
            context.origin = Origin(body, back if stages else None)
            findings = list(rule.run(doc))
            fixes: list[Finding] = []
            for finding in findings:
                if rule.code != DIRECTIVE_CODE and suppressions.suppresses(
                    rule.code, finding.start, finding.end
                ):
                    continue
                if finding.replacement is None:
                    if pass_number == 0:
                        pending.append(
                            _Pending(
                                rule.code, finding.message, finding.start, finding.end,
                                len(stages), False, finding.group,
                            )
                        )
                    continue
                if current[finding.start : finding.end] == finding.replacement:
                    continue
                fixes.append(finding)
            fixes = _non_overlapping(fixes)
            if not fixes:
                continue
            for fix in fixes:
                pending.append(
                    _Pending(rule.code, fix.message, fix.start, fix.end, len(stages), True, fix.group)
                )
            current = _apply(current, fixes)
            stages.append(_Stage(fixes))
            doc = Document(current)
            suppressions = scan_suppressions(doc, known_codes)
            changed = True
        if not changed:
            break
    else:
        raise UnstableFormattingError(
            "the rules did not settle after "
            f"{MAX_PASSES} passes; the last pass still changed the text"
        )

    if check_safety and current != body:
        difference = first_difference(body, current)
        if difference is not None:
            raise SafetyError(f"formatting would change the code's meaning ({difference})")

    line_starts = _line_starts(text)
    offset = len(header)
    violations = []
    reported_groups: set[tuple[str, str]] = set()
    if header_changed:
        names = {"\r\n": "CRLF", "\n": "LF", "\r": "CR"}
        assert header_newline is not None
        pending.insert(
            0,
            _Pending(
                "line-endings",
                f"Use {names[header_newline]} line endings throughout.",
                -len(header),
                -len(header),
                0,
                True,
                "line-endings",
            ),
        )
    for item in pending:
        if item.group is not None:
            key = (item.rule, item.group)
            if key in reported_groups:
                continue
            reported_groups.add(key)
        start = _to_original(item.start, stages[: item.stage])
        end = _to_original(item.end, stages[: item.stage])
        if end < start:
            end = start
        line, column = _line_col(line_starts, offset + start)
        end_line, end_column = _line_col(line_starts, offset + end)
        violations.append(
            Violation(line, column, end_line, end_column, item.rule, item.message, item.fixable)
        )
    violations.sort()
    return FormatResult(source=text, output=new_header + current, violations=violations)


def _convert_newlines(text: str, newline: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", newline)


def _to_original(pos: int, stages: Sequence[_Stage]) -> int:
    for stage in reversed(stages):
        pos = stage.back(pos)
    return pos


def _line_starts(text: str) -> list[int]:
    starts = [0]
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == "\r":
            i += 2 if i + 1 < n and text[i + 1] == "\n" else 1
            starts.append(i)
            continue
        if c == "\n":
            i += 1
            starts.append(i)
            continue
        i += 1
    return starts


def _line_col(starts: list[int], pos: int) -> tuple[int, int]:
    index = bisect.bisect_right(starts, pos) - 1
    return index + 1, pos - starts[index] + 1
