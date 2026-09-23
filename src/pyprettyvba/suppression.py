"""Suppression directives: comments that switch rules off for part of a module.

The syntax follows pyVBAanalysis's `'@pyvba-ignore`, in a namespace of its
own so each tool leaves the other's directives alone:

    x   =  1   '@prettyvba-ignore: spacing          this line
    '@prettyvba-ignore-next-line: indent            the next non-blank line
    '@prettyvba-ignore-file: identifier-case        the whole module
    '@prettyvba-ignore-start: indent, spacing       from here ...
    '@prettyvba-ignore-end                          ... to here

The rule list after the colon is optional; leaving it out, or writing
`all`, means every rule. A `-- reason` after the list is free text. A line
is the logical line: continuation lines belong to the line they continue.
`-next-line` skips blank lines and other directive comments, so directives
for different rules can be stacked above one line. `-file` must come before
the module's first line of code. Regions nest.

Only apostrophe comments are directives. `Rem` comments and `'''`
documentation comments never are.

A directive that cannot be read (an unknown verb or rule, a misplaced
`-file`, an unmatched `-end`) suppresses nothing, and a `-start` never
closed suppresses to the end of the module. Each is reported under the
`suppression-directive` code, which no directive can suppress.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .document import Document, LineKind
from .lexer import TokenKind

__all__ = ["DIRECTIVE_CODE", "Directive", "SuppressionIndex", "scan_suppressions"]

DIRECTIVE_CODE = "suppression-directive"

_NAMESPACE = "@prettyvba-ignore"
_CANDIDATE_RE = re.compile(r"^@prettyvba-ignore(?=$|[-\s:])", re.IGNORECASE)
_DIRECTIVE_RE = re.compile(
    r"^@prettyvba-ignore(?P<verb>-next-line|-file|-start|-end)?(?=$|[\s:])(?P<rest>.*)$",
    re.IGNORECASE | re.DOTALL,
)

ALL = "all"


@dataclass(frozen=True)
class Directive:
    """One parsed directive comment."""

    verb: str  # "", "next-line", "file", "start" or "end"
    codes: frozenset[str] | None  # None: every rule
    line: int  # index of the logical line the comment is on
    start: int  # the comment's character span
    end: int


@dataclass
class SuppressionIndex:
    """Which rules are suppressed where, for one version of a module's text."""

    file_codes: set[str] = field(default_factory=set)
    file_all: bool = False
    # (start offset, end offset, codes or None for every rule)
    ranges: list[tuple[int, int, frozenset[str] | None]] = field(default_factory=list)
    # (start, end, message) for each directive that could not be read.
    issues: list[tuple[int, int, str]] = field(default_factory=list)
    directives: list[Directive] = field(default_factory=list)

    def suppresses(self, code: str, start: int, end: int) -> bool:
        """True when ``code`` is off anywhere in the span [start, end)."""
        if code == DIRECTIVE_CODE:
            return False
        if self.file_all or code in self.file_codes:
            return True
        # An empty span (an insertion) still sits at a position; treat it as
        # covering one character so it meets the range it lands in.
        stop = end if end > start else start + 1
        for r_start, r_end, codes in self.ranges:
            if r_start < stop and start < r_end and (codes is None or code in codes):
                return True
        return False

    @property
    def empty(self) -> bool:
        return not (self.file_all or self.file_codes or self.ranges)


def scan_suppressions(doc: Document, known_codes: frozenset[str]) -> SuppressionIndex:
    """Read every directive in ``doc``."""
    index = SuppressionIndex()
    first_code_line = next(
        (line.index for line in doc.lines if line.kind in (LineKind.CODE, LineKind.DIRECTIVE)),
        None,
    )
    open_regions: list[Directive] = []
    directive_lines: set[int] = set()
    pending_next: list[Directive] = []

    for line in doc.lines:
        comment_at = line.comment
        if comment_at is None:
            continue
        token = doc.tokens[comment_at]
        if token.kind is not TokenKind.COMMENT:
            continue
        text = token.text
        if not text.startswith("'") or text.startswith("'''"):
            continue
        body = text[1:].lstrip()
        if _CANDIDATE_RE.match(body) is None:
            continue
        match = _DIRECTIVE_RE.match(body)
        if match is None:
            word = body.split()[0] if body.split() else body
            index.issues.append((token.start, token.end, f"Unknown suppression directive {word!r}."))
            continue
        verb = (match.group("verb") or "").lstrip("-").lower()
        codes, problems = _parse_codes(match.group("rest"), known_codes)
        for problem in problems:
            index.issues.append((token.start, token.end, problem))
        directive = Directive(verb, codes, line.index, token.start, token.end)
        index.directives.append(directive)
        if line.kind is LineKind.COMMENT:
            directive_lines.add(line.index)
        if verb == "":
            index.ranges.append((line.start, line.end, codes))
        elif verb == "next-line":
            pending_next.append(directive)
        elif verb == "file":
            if first_code_line is not None and line.index > first_code_line:
                index.issues.append(
                    (token.start, token.end,
                     "'@prettyvba-ignore-file must come before the first line of code.")
                )
                continue
            if codes is None:
                index.file_all = True
            else:
                index.file_codes |= codes
        elif verb == "start":
            open_regions.append(directive)
        else:  # end
            if not open_regions:
                index.issues.append(
                    (token.start, token.end, "'@prettyvba-ignore-end has no matching -start.")
                )
                continue
            opened = open_regions.pop()
            index.ranges.append((doc.lines[opened.line].start, line.end, opened.codes))

    for directive in pending_next:
        target = _next_line(doc, directive.line, directive_lines)
        if target is None:
            index.issues.append(
                (directive.start, directive.end, "'@prettyvba-ignore-next-line has no line after it.")
            )
            continue
        index.ranges.append((doc.lines[target].start, doc.lines[target].end, directive.codes))

    for opened in open_regions:
        index.issues.append(
            (opened.start, opened.end, "'@prettyvba-ignore-start is never closed with -end.")
        )
        index.ranges.append((doc.lines[opened.line].start, len(doc.text), opened.codes))
    return index


def _next_line(doc: Document, after: int, directive_lines: set[int]) -> int | None:
    for line in doc.lines[after + 1 :]:
        if line.kind is LineKind.BLANK or line.index in directive_lines:
            continue
        return line.index
    return None


def _parse_codes(rest: str, known: frozenset[str]) -> tuple[frozenset[str] | None, list[str]]:
    """Read an optional `: code, code -- reason` tail. None means every rule."""
    problems: list[str] = []
    rest = rest.strip()
    if rest.startswith(":"):
        rest = rest[1:].strip()
    if "--" in rest:
        rest = rest.split("--", 1)[0].strip()
    if rest == "" or rest.lower() == ALL:
        return None, problems
    codes: set[str] = set()
    for part in rest.split(","):
        code = part.strip().lower()
        if not code:
            problems.append("Empty rule name in suppression directive.")
        elif code == ALL:
            problems.append("'all' cannot be combined with rule names.")
        elif code not in known:
            problems.append(f"Unknown rule {code!r} in suppression directive.")
        else:
            codes.add(code)
    if problems and not codes:
        # Nothing readable: suppress nothing rather than everything.
        return frozenset(), problems
    return frozenset(codes), problems
