"""Layout: indentation, continuation lines, labels, blank lines, line ends.

The VBE never indents code; it keeps whatever indentation a line has. It
does move a line label or line number to column 1, keeping the statement
after it where it was (tests/oracle/vbe_rendering.json), and it trims the
whitespace at the end of a line. Everything else here is layout the VBE
leaves to the author, which these rules make consistent.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator

from ..document import Document, LineKind, LogicalLine, StatementKind
from ..lexer import Token, TokenKind, tokenize
from ..structure import analyze_structure
from .base import Finding, Option, Rule
from .spacing import display_width

__all__ = [
    "BlankLinesRule",
    "ContinuationIndentRule",
    "EndOfFileRule",
    "IndentRule",
    "LineEndingsRule",
    "TrailingWhitespaceRule",
    "detect_newline",
]


def _leading(doc: Document, line: LogicalLine) -> tuple[int, int, str]:
    """(start, end, text) of the whitespace a logical line starts with."""
    tokens = doc.tokens
    if line.first < line.stop and tokens[line.first].kind is TokenKind.WHITESPACE:
        tok = tokens[line.first]
        return tok.start, tok.end, tok.text
    return line.start, line.start, ""


def _opens_with_continuation(doc: Document, line: LogicalLine) -> bool:
    """True for a line whose first physical line is only ` _`.

    The space there belongs to the continuation (`_` alone is not one), so
    it is not indentation to change.
    """
    tokens = doc.tokens
    k = line.first
    if k < line.stop and tokens[k].kind is TokenKind.WHITESPACE:
        k += 1
    return k < line.stop and tokens[k].kind is TokenKind.CONTINUATION


def _statement_head(doc: Document, line: LogicalLine) -> Token | None:
    """The first token after a line's label: its first statement or comment."""
    if line.statements:
        return doc.tokens[line.statements[0].tokens[0]]
    # A label followed by a separator and nothing else, or by a comment.
    after = line.label_colon if line.label_colon is not None else line.label
    if after is not None:
        k = doc.next_code(after)
        while k is not None and doc.tokens[k].kind is TokenKind.COLON:
            k = doc.next_code(k)
        return doc.tokens[k] if k is not None else None
    return None


def _continuation_gaps(doc: Document, line: LogicalLine) -> Iterator[tuple[int, int, str, Token | None]]:
    """For each continuation line: (start, end, text) of its indentation and its first token."""
    tokens = doc.tokens
    for j in range(line.first, line.stop):
        if tokens[j].kind is not TokenKind.CONTINUATION:
            continue
        k = j + 1
        if k < line.stop and tokens[k].kind is TokenKind.WHITESPACE:
            ws = tokens[k]
            nxt = tokens[k + 1] if k + 1 < line.stop else None
            if nxt is not None and nxt.kind is TokenKind.CONTINUATION:
                # ` _` alone on a line: the space belongs to the continuation.
                continue
            yield ws.start, ws.end, ws.text, nxt
        else:
            nxt = tokens[k] if k < line.stop else None
            yield tokens[j].end, tokens[j].end, "", nxt


class IndentRule(Rule):
    """Indent each line by the blocks around it.

    Procedure bodies, block If arms, loops, With, Select Case arms and their
    statements, and Type and Enum members each sit one level in. Labels and
    line numbers go to column 1, as the VBE puts them, with the statement
    after them at its level. Continuation lines move with the line they
    continue. Where the block structure cannot be matched (an End If with no
    If, a For never closed), the region is left as written and reported.
    """

    code = "indent"
    summary = "Indent lines by block structure."
    category = "layout"
    options = (
        Option(
            "reindent",
            True,
            "Indent by block structure. When false, each line keeps its indentation "
            "(rewritten in the configured style, tabs expanded) and only line labels "
            "move to column 1, with the statement after them kept in its column: "
            "exactly what the VBE does.",
        ),
        Option("procedure-body", True, "Indent procedure bodies one level."),
        Option("case", True, "Indent Case arms one level inside Select Case (their statements one more)."),
        Option("type-members", True, "Indent the members of Type and Enum blocks."),
        Option(
            "directives",
            "indent",
            "Place `#If` lines: `indent` keeps them at the level of the code around them "
            "and indents a balanced block's content one level; `flat` keeps them at that "
            "level without indenting the content; `column-zero` puts them at column 1; "
            "`preserve` leaves them where they are.",
            choices=("indent", "flat", "column-zero", "preserve"),
        ),
        Option(
            "comments",
            "code",
            "Place comment-only lines: `code` at the level of the code around them, "
            "`next` at the level of the next line of code, `preserve` where they are.",
            choices=("code", "next", "preserve"),
        ),
        Option("debug-column-zero", False, "Put `Debug.Print` and `Debug.Assert` lines at column 1."),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        s = self.settings
        if not s["reindent"]:
            yield from self._keep(doc)
            return
        structure = analyze_structure(
            doc,
            procedure_body=bool(s["procedure-body"]),
            case_arms=bool(s["case"]),
            type_members=bool(s["type-members"]),
            directive_arms=s["directives"] == "indent",
        )
        width = self.context.indent_width
        levels = structure.levels
        # For `comments = "next"`: the level of the next line of code.
        next_level = list(levels)
        upcoming: int | None = None
        for index in range(len(doc.lines) - 1, -1, -1):
            kind = doc.lines[index].kind
            if kind in (LineKind.CODE, LineKind.DIRECTIVE):
                upcoming = levels[index]
            elif kind is LineKind.COMMENT and upcoming is not None:
                next_level[index] = upcoming

        for line_index, message in structure.problems:
            line = doc.lines[line_index]
            head = line.head if line.head is not None else line.first
            tok = doc.tokens[head] if head < len(doc.tokens) else None
            start = tok.start if tok is not None else line.start
            yield Finding(start, start, None, f"{message} Indentation here is left as written.")

        for line in doc.lines:
            if line.kind in (LineKind.BLANK, LineKind.ATTRIBUTE) or line.head is None:
                continue
            if structure.is_broken(line.index) or _opens_with_continuation(doc, line):
                continue
            level = levels[line.index]
            if line.kind is LineKind.DIRECTIVE:
                mode = s["directives"]
                if mode == "preserve":
                    continue
                columns = 0 if mode == "column-zero" else level * width
            elif line.kind is LineKind.COMMENT:
                mode = s["comments"]
                if mode == "preserve":
                    continue
                columns = (next_level[line.index] if mode == "next" else level) * width
            else:
                columns = level * width
                if s["debug-column-zero"] and _is_debug_statement(doc, line):
                    columns = 0
            yield from self._place(doc, line, columns)

    def _keep(self, doc: Document) -> Iterator[Finding]:
        """What the VBE does to indentation: labels to column 1, tabs to spaces."""
        ctx = self.context
        tab = ctx.tab_width
        for line in doc.lines:
            if line.kind in (LineKind.BLANK, LineKind.ATTRIBUTE) or line.head is None:
                continue
            if _opens_with_continuation(doc, line):
                continue
            start, end, text = _leading(doc, line)
            if line.label is None:
                wanted = ctx.indent_text(ctx.width(text))
                if wanted != text:
                    yield Finding(start, end, wanted, "Write the indentation in the configured style.")
                for c_start, c_end, c_text, _first in _continuation_gaps(doc, line):
                    c_wanted = ctx.indent_text(ctx.width(c_text))
                    if c_wanted != c_text:
                        yield Finding(c_start, c_end, c_wanted, "Write the indentation in the configured style.")
                continue
            if not text:
                continue
            yield Finding(start, end, "", "Put the line label in column 1, as the VBE does.")
            head = _statement_head(doc, line)
            if head is None or head.kind in (TokenKind.COMMENT, TokenKind.UNKNOWN):
                continue
            label_end = line.label_colon if line.label_colon is not None else line.label
            gap_start = doc.tokens[label_end].end
            gap = doc.text[gap_start : head.start]
            if "\n" in gap or "\r" in gap:
                continue
            old_head = display_width(doc.text[start : head.start], 0, tab)
            label_width = display_width(doc.text[end:gap_start], 0, tab)
            wanted = " " * max(1, old_head - label_width)
            if wanted != gap:
                yield Finding(gap_start, head.start, wanted, "Keep the statement after the label in its column.")

    def _place(self, doc: Document, line: LogicalLine, columns: int) -> Iterator[Finding]:
        ctx = self.context
        tab = ctx.tab_width
        start, end, text = _leading(doc, line)
        old_col = ctx.width(text)
        if line.label is None:
            wanted = ctx.indent_text(columns)
            if wanted != text:
                yield Finding(start, end, wanted, _indent_message(columns, old_col))
            delta = columns - old_col
        else:
            # The label goes to column 1; the statement after it to its level.
            label = doc.tokens[line.label]
            if text:
                yield Finding(start, end, "", "Put the line label in column 1, as the VBE does.")
            label_end = line.label_colon if line.label_colon is not None else line.label
            head = _statement_head(doc, line)
            if head is None or head.kind in (TokenKind.COMMENT, TokenKind.UNKNOWN):
                return
            gap_start = doc.tokens[label_end].end
            gap = doc.text[gap_start : head.start]
            if "\n" in gap or "\r" in gap:
                return
            label_width = display_width(doc.text[label.start : gap_start], 0, tab)
            old_head = old_col + label_width + display_width(gap, old_col + label_width, tab)
            target = max(columns, label_width + 1)
            wanted = " " * (target - label_width)
            if wanted != gap:
                yield Finding(gap_start, head.start, wanted, _indent_message(target, old_head))
            delta = target - old_head
        if delta:
            for c_start, c_end, c_text, _first in _continuation_gaps(doc, line):
                old = ctx.width(c_text)
                new = max(0, old + delta)
                wanted = ctx.indent_text(new)
                if wanted != c_text:
                    yield Finding(c_start, c_end, wanted, "Move the continuation line with the line it continues.")


def _indent_message(columns: int, old: int) -> str:
    return f"Indent to column {columns + 1} (found column {old + 1})."


def _is_debug_statement(doc: Document, line: LogicalLine) -> bool:
    if not line.statements:
        return False
    indices = line.statements[0].tokens
    if len(indices) < 3:
        return False
    a, b, c = (doc.tokens[j] for j in indices[:3])
    return a.lower == "debug" and b.text == "." and c.lower in ("print", "assert")


class ContinuationIndentRule(Rule):
    """Indent continuation lines relative to the line they continue.

    `relative` keeps a continuation line's extra indentation and gives one
    level to a line that has none (or sits left of its first line), which is
    what XLIDE does. `hanging` puts every continuation line one level (the
    `hanging` option) in from the first line. `preserve` leaves them.
    """

    code = "continuation-indent"
    summary = "Indent continuation lines consistently."
    category = "layout"
    options = (
        Option(
            "style",
            "relative",
            "`relative` keeps extra indentation and fixes lines left of their first "
            "line; `hanging` indents every continuation line by a fixed amount; "
            "`preserve` leaves them.",
            choices=("relative", "hanging", "preserve"),
        ),
        Option("hanging", 1, "Levels a hanging continuation line sits in from its first line.", minimum=0, maximum=8),
        Option(
            "closing-paren",
            "align",
            "A continuation line that starts with `)`: `align` lets it sit in the "
            "column of the line it continues (`) As Long` under its declaration); "
            "`indent` treats it like any other continuation line, as XLIDE does.",
            choices=("align", "indent"),
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        style = self.settings["style"]
        align_closers = self.settings["closing-paren"] == "align"
        if style == "preserve":
            return
        ctx = self.context
        step = ctx.indent_width * int(self.settings["hanging"])
        for line in doc.lines:
            if line.kind not in (LineKind.CODE, LineKind.DIRECTIVE) or line.head is None:
                continue
            if _opens_with_continuation(doc, line):
                continue
            gaps = list(_continuation_gaps(doc, line))
            if not gaps:
                continue
            head = _statement_head(doc, line) if line.label is not None else doc.tokens[line.head]
            if head is None or any(
                doc.tokens[k].kind is TokenKind.CONTINUATION and doc.tokens[k].end <= head.start
                for k in range(line.first, line.stop)
            ):
                # The statement itself starts on a continuation line
                # (`10 _` above it): there is no first line to indent from.
                continue
            first_col = self._column(doc, head)
            for c_start, c_end, c_text, first in gaps:
                if first is None or first.kind is TokenKind.NEWLINE:
                    continue
                col = ctx.width(c_text)
                closer = align_closers and first.kind is TokenKind.PUNCTUATION and first.text == ")"
                if closer and col == first_col:
                    # `) As Long` under its declaration: a closing parenthesis
                    # lined up with the statement it closes is deliberate.
                    continue
                if style == "hanging":
                    target = first_col + step
                elif col <= first_col:
                    target = first_col + max(step, ctx.indent_width)
                else:
                    continue
                wanted = ctx.indent_text(target)
                if wanted != c_text:
                    yield Finding(c_start, c_end, wanted, f"Indent the continuation line to column {target + 1}.")

    def _column(self, doc: Document, tok: Token) -> int:
        start = doc.physical_starts[doc.physical_line_of(tok.start)]
        return display_width(doc.text[start : tok.start], 0, self.context.tab_width)


class BlankLinesRule(Rule):
    """Limit blank lines, and separate procedures by a set number of them.

    Runs of blank lines are cut to `max-consecutive`. Each procedure gets
    `between-procedures` blank lines above it, counted above the comment
    block that documents it, so the comment stays attached to its
    procedure. Blank lines right after a procedure's header and right before
    its End go with `trim-procedures`, and blank lines at the top of the
    module with `leading`. The VBE leaves blank lines alone; blank lines at
    the end of the file belong to end-of-file.
    """

    code = "blank-lines"
    summary = "Keep blank lines consistent."
    category = "layout"
    options = (
        Option("max-consecutive", 2, "The most blank lines in a row.", minimum=0),
        Option(
            "between-procedures",
            1,
            "Blank lines before each procedure (and its comment block), counted from "
            "the code above it. -1 leaves them as they are.",
            minimum=-1,
        ),
        Option("trim-procedures", True, "Remove blank lines right after a procedure header and right before its End."),
        Option("leading", 0, "Blank lines allowed at the start of the module body.", minimum=0),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        s = self.settings
        max_run = int(s["max-consecutive"])
        between = int(s["between-procedures"])
        trim = bool(s["trim-procedures"])
        lines = doc.lines
        structure = analyze_structure(doc)
        # Headers and End lines of procedures that span more than one line
        # and are closed by an End line.
        closed = [
            (first, last)
            for first, last in structure.procedures
            if last > first and any(st.kind is StatementKind.PROC_END for st in lines[last].statements)
        ]
        headers = {first for first, _last in closed}
        closers = {last for _first, last in closed}
        # A procedure's leading block: the comment lines right above its header.
        block_start: dict[int, int] = {}
        for first, _last in structure.procedures:
            top = first
            while top - 1 >= 0 and lines[top - 1].kind is LineKind.COMMENT:
                top -= 1
            block_start[top] = first
        plural = "s" if max_run != 1 else ""

        index = 0
        count = len(lines)
        while index < count:
            if lines[index].kind is not LineKind.BLANK:
                index += 1
                continue
            run_start = index
            while index < count and lines[index].kind is LineKind.BLANK:
                index += 1
            run_end = index  # exclusive
            if run_end >= count:
                continue  # trailing blank lines belong to end-of-file
            prev_index = run_start - 1
            have = run_end - run_start
            wanted = min(have, max_run)
            reason = f"Use at most {max_run} blank line{plural} in a row."
            if prev_index < 0:
                wanted = min(have, int(s["leading"]))
                reason = "Remove blank lines at the start of the module."
            elif trim and prev_index in headers:
                wanted = 0
                reason = "Remove blank lines after a procedure header."
            elif trim and run_end in closers:
                wanted = 0
                reason = "Remove blank lines before the end of a procedure."
            elif between >= 0 and run_end in block_start and self._separates(doc, prev_index):
                wanted = between
                reason = f"Separate procedures with {between} blank line{'s' if between != 1 else ''}."
            if have > wanted:
                first_removed = lines[run_start + wanted]
                last_removed = lines[run_end - 1]
                yield Finding(first_removed.start, last_removed.end, "", reason)
            elif have < wanted:
                at = lines[run_end].start
                yield Finding(at, at, self._newline(doc, lines[run_end - 1]) * (wanted - have), reason)

        if between > 0:
            # Procedures with no blank line at all above their leading block.
            for top in sorted(block_start):
                prev_index = top - 1
                if prev_index < 0 or lines[prev_index].kind is LineKind.BLANK:
                    continue
                if not self._separates(doc, prev_index) or (trim and prev_index in headers):
                    # Right after another procedure's header, trimming wins,
                    # as it does above.
                    continue
                at = lines[top].start
                yield Finding(
                    at,
                    at,
                    self._newline(doc, lines[prev_index]) * between,
                    f"Separate procedures with {between} blank line{'s' if between != 1 else ''}.",
                )

    @staticmethod
    def _separates(doc: Document, prev_index: int) -> bool:
        """True when blank lines above a procedure follow code it should be set apart from.

        A directive line above it (the procedure opens inside an #If arm) is
        left alone.
        """
        prev = doc.lines[prev_index]
        return prev.kind in (LineKind.CODE, LineKind.COMMENT, LineKind.ATTRIBUTE) or (
            prev.kind is LineKind.DIRECTIVE
            and bool(prev.statements)
            and prev.statements[0].kind is StatementKind.CC_END_IF
        )

    def _newline(self, doc: Document, line: LogicalLine) -> str:
        tokens = doc.tokens
        if line.stop > line.first and tokens[line.stop - 1].kind is TokenKind.NEWLINE:
            return tokens[line.stop - 1].text
        return self.context.newline


class TrailingWhitespaceRule(Rule):
    """Remove whitespace at the end of lines, as the VBE does.

    The VBE trims every line it stores, but keeps a line that holds nothing
    but whitespace (measured); `blank-lines` chooses what happens to those.
    Whitespace after a line continuation's `_` goes too: the continuation
    still continues without it. The lines of a comment continued with ` _`
    are trimmed like any other.
    """

    code = "trailing-whitespace"
    summary = "Remove whitespace at the end of lines."
    category = "layout"
    vbe_canonical = True
    options = (
        Option(
            "blank-lines",
            "remove",
            "Whitespace on otherwise blank lines: `remove` it, `keep` it, or "
            "`indent` the line to the level of the code around it, the way the VBE's "
            "editor leaves a blank line (and XLIDE's Format Document writes it).",
            choices=("remove", "keep", "indent"),
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        mode = self.settings["blank-lines"]
        tokens = doc.tokens
        count = len(tokens)
        levels: list[int] | None = None
        if mode == "indent":
            indent = self.context.rule_settings.get("indent", {})
            levels = analyze_structure(
                doc,
                procedure_body=bool(indent.get("procedure-body", True)),
                case_arms=bool(indent.get("case", True)),
                type_members=bool(indent.get("type-members", True)),
                directive_arms=indent.get("directives", "indent") == "indent",
            ).levels
        for line in doc.lines:
            continued = any(tokens[j].kind is TokenKind.CONTINUATION for j in range(line.first, line.stop))
            # ` _` on a line of its own continues onto the next: its space is
            # part of the continuation, not indentation.
            if line.kind is LineKind.BLANK and levels is not None and not continued:
                wanted = self.context.indent_text(levels[line.index] * self.context.indent_width)
                start, end, text = _leading(doc, line)
                at_end = line.stop == line.first or tokens[line.stop - 1].kind is not TokenKind.NEWLINE
                if text != wanted and not (at_end and not text):
                    yield Finding(start, end, wanted, "Indent the blank line to the code around it.")
                continue
            for j in range(line.first, line.stop):
                tok = tokens[j]
                if tok.kind is TokenKind.WHITESPACE:
                    nxt = tokens[j + 1] if j + 1 < count else None
                    if nxt is None or nxt.kind is TokenKind.NEWLINE:
                        if mode == "keep" and line.kind is LineKind.BLANK:
                            continue
                        yield Finding(tok.start, tok.end, "", "Remove trailing whitespace.")
                elif tok.kind is TokenKind.CONTINUATION:
                    body = tok.text.rstrip("\r\n")
                    if body != "_":
                        yield Finding(tok.start + 1, tok.start + len(body), "", "Remove trailing whitespace.")
                elif tok.kind is TokenKind.COMMENT:
                    yield from _comment_trailing(tok)


def _comment_trailing(tok: Token) -> Iterator[Finding]:
    """Trailing whitespace on each physical line a comment spans."""
    text = tok.text
    pos = 0
    while pos <= len(text):
        brk = len(text)
        for sep in ("\r", "\n"):
            found = text.find(sep, pos)
            if found != -1:
                brk = min(brk, found)
        segment = text[pos:brk]
        stripped = segment.rstrip(" \t")
        if len(stripped) != len(segment):
            yield Finding(tok.start + pos + len(stripped), tok.start + brk, "", "Remove trailing whitespace.")
        if brk >= len(text):
            break
        pos = brk + (2 if text.startswith("\r\n", brk) else 1)


class EndOfFileRule(Rule):
    """End the module with exactly one line break and no blank lines.

    The VBE writes every line of an exported module with a line break after
    it, and keeps the blank lines at the end of a module it imports
    (measured), so the `vbe` preset keeps them too.
    """

    code = "end-of-file"
    summary = "End the file with one line break."
    category = "layout"
    options = (
        Option("final-newline", True, "Require a line break after the last line."),
        Option(
            "trailing-blank-lines",
            "remove",
            "Blank lines after the last line of code: `remove` them, or `keep` them "
            "as the VBE does.",
            choices=("remove", "keep"),
        ),
    )

    def run(self, doc: Document) -> Iterable[Finding]:
        text = doc.text
        if not text.strip():
            return
        lines = doc.lines
        last = len(lines) - 1
        while last >= 0 and lines[last].kind is LineKind.BLANK:
            last -= 1
        if last < 0:
            return
        if self.settings["trailing-blank-lines"] == "keep":
            # The file's own last line, blank or not, is the one to end.
            last = len(lines) - 1
        final = lines[last]
        tokens = doc.tokens
        has_break = final.stop > final.first and tokens[final.stop - 1].kind is TokenKind.NEWLINE
        if final.end < len(text):
            yield Finding(final.end, len(text), "", "Remove blank lines at the end of the file.")
        if not has_break and self.settings["final-newline"]:
            newline = self.context.newline
            if _ends_the_same(text[final.start : final.end], newline):
                yield Finding(final.end, final.end, newline, "End the file with a line break.")
            else:
                # `Foo _` with a line break after it is a line continuation.
                yield Finding(final.end, final.end, None, "End the file with a line break.")


def _ends_the_same(line: str, newline: str) -> bool:
    """True when adding a line break after the last line changes no token."""
    before = [(tok.kind, tok.text) for tok in tokenize(line)]
    after = [(tok.kind, tok.text) for tok in tokenize(line + newline)]
    return after[:-1] == before and after[-1][0] is TokenKind.NEWLINE


def detect_newline(text: str, setting: str) -> str:
    """The line terminator a file should use under a line-ending setting."""
    if setting == "crlf":
        return "\r\n"
    if setting == "lf":
        return "\n"
    crlf = text.count("\r\n")
    lf = text.count("\n") - crlf
    cr = text.count("\r") - crlf
    if lf > crlf and lf >= cr:
        return "\n"
    if cr > crlf and cr > lf:
        return "\r"
    return "\r\n"


class LineEndingsRule(Rule):
    """Use one line ending throughout: the file's own, or the configured one.

    The VBE writes CRLF and reads a module saved with bare LF line ends badly
    on import, so `crlf` is the safe choice for code that goes back into
    Office. The default, `auto`, keeps whichever ending the file mostly uses.
    """

    code = "line-endings"
    summary = "Use one line ending throughout."
    category = "layout"

    def run(self, doc: Document) -> Iterable[Finding]:
        target = self.context.newline
        message = f"Use {_ending_name(target)} line endings throughout."
        for tok in doc.tokens:
            if tok.kind is TokenKind.NEWLINE:
                if tok.text != target:
                    yield Finding(tok.start, tok.end, target, message, group="line-endings")
            elif tok.kind is TokenKind.CONTINUATION:
                body = tok.text.rstrip("\r\n")
                if tok.text[len(body) :] != target:
                    yield Finding(tok.start + len(body), tok.end, target, message, group="line-endings")
            elif tok.kind is TokenKind.COMMENT and ("\r" in tok.text or "\n" in tok.text):
                pos = 0
                text = tok.text
                while pos < len(text):
                    c = text[pos]
                    if c == "\r" or c == "\n":
                        size = 2 if text.startswith("\r\n", pos) else 1
                        if text[pos : pos + size] != target:
                            yield Finding(tok.start + pos, tok.start + pos + size, target, message, group="line-endings")
                        pos += size
                        continue
                    pos += 1


def _ending_name(newline: str) -> str:
    return {"\r\n": "CRLF", "\n": "LF", "\r": "CR"}[newline]
