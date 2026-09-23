"""Render every fixture case as one self-contained page: the fixture gallery.

    python tools/build_gallery.py [--out artifacts/gallery/index.html]

For each case under tests/fixtures the page shows its description, its
configuration as a pyprettyvba.toml, the module as written beside the module
as formatted (with the changed characters marked), and what `check`
reports; for a command-line case, the command and what it prints. The code
is colored the way the VBE colors its code window, with pyPrettyVBA's own
lexer. The page loads nothing from anywhere: fonts are the system's, and
the few lines of script only filter, toggle and scroll.

The output is a page body for publishing as a claude.ai artifact (no
<html> or <body> tags; the host adds them).
"""

from __future__ import annotations

import argparse
import difflib
import html
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

from pyprettyvba import PRESETS, RULES, Config, __version__  # noqa: E402
from pyprettyvba.document import split_header  # noqa: E402
from pyprettyvba.lexer import TokenKind, tokenize  # noqa: E402
from pyprettyvba.textio import decode  # noqa: E402

from fixture_support import FIXTURES, Case, load_cases  # noqa: E402

DEFAULT_OUT = REPO / "artifacts" / "gallery" / "index.html"

SURFACES: list[tuple[str, str, str]] = [
    ("showcase", "Showcase", "One untidy module, formatted with the defaults."),
    ("rules", "Rules", "Each rule on its own, with the `none` preset and only that rule on, and each of its options."),
    ("presets", "Presets", "One module through each of the six presets."),
    ("suppression", "Suppression", "`'@prettyvba-ignore` comments switch rules off for a line, the next line, a region or a module."),
    ("config", "Configuration", "Configuration files, per-file overrides, `extend`, and the errors a bad file gets."),
    ("files", "Files", "Module kinds, headers, encodings and line endings: the bytes a file holds come back as they were."),
    ("safety", "Safety", "Places where whitespace or letter case changes what the code does, and the formatter holds back."),
    ("project", "Projects", "Modules formatted together: a name one module declares is spelled its way in the others."),
    ("cli", "Command line", "`format`, `check`, `rules` and `config`, as a terminal shows them."),
]

# Cases where whitespace is the point: their markers are always shown.
WHITESPACE_CASES = (
    "rules/trailing-whitespace/", "rules/line-endings/", "rules/end-of-file/", "rules/indent/tabs",
    "files/lf-endings", "safety/continuation-trailing-space", "rules/blank-lines/",
)

TOKEN_CLASS = {
    TokenKind.KEYWORD: "k",
    TokenKind.DIRECTIVE: "k",
    TokenKind.COMMENT: "c",
    TokenKind.STRING: "s",
    TokenKind.INTEGER: "n",
    TokenKind.FLOAT: "n",
    TokenKind.DATE: "n",
    TokenKind.UNKNOWN: "x",
}

# What counts as trailing whitespace on a line: space, tab, ideographic and
# no-break space.
TRAILING = " \t" + chr(0x3000) + chr(0xA0)

CHECK_RE = re.compile(r"^(?:(?P<file>[^:\s]+\.\w+):)?(?P<line>\d+):(?P<col>\d+) (?P<rule>[\w-]+) (?P<msg>.*?)(?P<fix> \[\*\])?$")


# --------------------------------------------------------------------- code


@dataclass
class Line:
    start: int
    end: int  # before the line break
    eol: str  # "crlf", "lf", "cr", or "" for none


def split_lines(text: str) -> list[Line]:
    lines: list[Line] = []
    i = start = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == "\r":
            crlf = i + 1 < n and text[i + 1] == "\n"
            lines.append(Line(start, i, "crlf" if crlf else "cr"))
            i += 2 if crlf else 1
            start = i
            continue
        if c == "\n":
            lines.append(Line(start, i, "lf"))
            i += 1
            start = i
            continue
        i += 1
    if start < n or not lines:
        lines.append(Line(start, n, ""))
    return lines


def syntax_classes(text: str) -> list[str]:
    classes = [""] * len(text)
    for tok in tokenize(text):
        cls = TOKEN_CLASS.get(tok.kind, "")
        if cls:
            for i in range(tok.start, tok.end):
                classes[i] = cls
    return classes


@dataclass
class Marks:
    """What changed, per line: whole-line flag, changed character offsets, changed line break."""

    lines: dict[int, str] = field(default_factory=dict)  # line index -> "chg" | "gone" | "new"
    chars: set[int] = field(default_factory=set)
    eols: set[int] = field(default_factory=set)


def diff_marks(a_text: str, b_text: str) -> tuple[Marks, Marks]:
    a_lines, b_lines = split_lines(a_text), split_lines(b_text)
    a_keys = [(a_text[x.start : x.end], x.eol) for x in a_lines]
    b_keys = [(b_text[x.start : x.end], x.eol) for x in b_lines]
    a_marks, b_marks = Marks(), Marks()
    matcher = difflib.SequenceMatcher(None, a_keys, b_keys, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        pairs = min(i2 - i1, j2 - j1) if tag == "replace" else 0
        for k in range(pairs):
            ai, bi = i1 + k, j1 + k
            a_marks.lines[ai] = b_marks.lines[bi] = "chg"
            a_line, b_line = a_lines[ai], b_lines[bi]
            a_str, b_str = a_keys[ai][0], b_keys[bi][0]
            chars = difflib.SequenceMatcher(None, a_str, b_str, autojunk=False)
            for op, x1, x2, y1, y2 in chars.get_opcodes():
                if op == "equal":
                    continue
                a_marks.chars.update(range(a_line.start + x1, a_line.start + x2))
                b_marks.chars.update(range(b_line.start + y1, b_line.start + y2))
                if op == "insert" and x1 == x2:
                    # Show where something was inserted: mark the neighbour.
                    if x1 < len(a_str):
                        a_marks.chars.add(a_line.start + x1)
                if op == "delete" and y1 == y2 and y1 < len(b_str):
                    b_marks.chars.add(b_line.start + y1)
            if a_keys[ai][1] != b_keys[bi][1]:
                a_marks.eols.add(ai)
                b_marks.eols.add(bi)
        for ai in range(i1 + pairs, i2):
            a_marks.lines[ai] = "gone"
        for bi in range(j1 + pairs, j2):
            b_marks.lines[bi] = "new"
    return a_marks, b_marks


def _run_html(cls: str, text: str) -> str:
    """One run of characters. A byte the code page does not define (read with
    surrogateescape) is shown as its value, since no character stands for it."""
    parts = []
    for piece in re.split(r"([\udc80-\udcff])", text):
        if len(piece) == 1 and 0xDC80 <= ord(piece) <= 0xDCFF:
            parts.append(f'<span class="raw" title="a byte the code page does not define">0x{ord(piece) - 0xDC00:02X}</span>')
        elif piece:
            parts.append(html.escape(piece, quote=False))
    body = "".join(parts)
    return f'<span class="{cls}">{body}</span>' if cls else body


def render_code(text: str, *, marks: Marks | None, side: str, anchor: str, lines_hit: set[int] | None = None) -> str:
    """HTML for a code pane: one row per line, colored, with changes marked."""
    classes = syntax_classes(text)
    header_end = len(split_header(text)[0])
    rows = []
    lines = split_lines(text)
    for index, line in enumerate(lines):
        content = text[line.start : line.end]
        last_solid = len(content.rstrip(TRAILING))
        # Runs of characters that share their classes, as (classes, characters).
        runs: list[tuple[str, list[str]]] = []
        for offset in range(line.start, line.end):
            ch = text[offset]
            parts = [classes[offset]] if classes[offset] else []
            if ch == " ":
                parts.append("w")
            elif ch == "\t":
                parts.append("t")
            if ch in " \t" and offset - line.start >= last_solid:
                parts.append("tr")
            if marks is not None and offset in marks.chars:
                parts.append("d")
            cls = " ".join(parts)
            if runs and runs[-1][0] == cls:
                runs[-1][1].append(ch)
            else:
                runs.append((cls, [ch]))
        pieces = [_run_html(cls, "".join(chars)) for cls, chars in runs]
        eol_cls = f"eol {line.eol or 'none'}"
        if marks is not None and index in marks.eols:
            eol_cls += " d"
        last = index == len(lines) - 1
        if line.eol or (last and text):
            pieces.append(f'<span class="{eol_cls}" aria-hidden="true"></span>')
        row_cls = "ln"
        if marks is not None and index in marks.lines:
            row_cls += f" {marks.lines[index]}"
        if line.end <= header_end and header_end:
            row_cls += " hdr"
        if lines_hit and index + 1 in lines_hit:
            row_cls += " hit"
        rows.append(
            f'<div class="{row_cls}" id="{anchor}-{side}-{index + 1}"><span class="no">{index + 1}</span>'
            f'<span class="tx">{"".join(pieces)}</span></div>'
        )
    if not text:
        rows = ['<div class="ln empty"><span class="no">1</span><span class="tx">(an empty file)</span></div>']
    return f'<div class="code" tabindex="0"><div class="rows">{"".join(rows)}</div></div>'


def file_facts(data: bytes, encoding: str) -> list[str]:
    decoded = decode(data, encoding)
    facts = []
    if decoded.encoding.replace("-", "").lower() in ("utf8",):
        facts.append("UTF-8 with BOM" if decoded.bom else "UTF-8")
    else:
        facts.append(decoded.encoding)
    text = decoded.text
    crlf = text.count("\r\n")
    lf = text.count("\n") - crlf
    cr = text.count("\r") - crlf
    kinds = [name for name, count in (("CRLF", crlf), ("LF", lf), ("CR", cr)) if count]
    facts.append(" + ".join(kinds) if kinds else "one line")
    return facts


# --------------------------------------------------------------------- text


def inline(text: str) -> str:
    """Escape text and turn `code` spans into <code>."""
    parts = re.split(r"(`[^`]+`)", text)
    out = []
    for part in parts:
        if part.startswith("`") and part.endswith("`") and len(part) > 1:
            out.append(f"<code>{html.escape(part[1:-1], quote=False)}</code>")
        else:
            out.append(html.escape(part, quote=False))
    return "".join(out)


def paragraphs(text: str) -> str:
    blocks = [" ".join(block.split()) for block in re.split(r"\n\s*\n", text.strip()) if block.strip()]
    return "".join(f"<p>{inline(block)}</p>" for block in blocks)


def _toml_value(value: object) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return json.dumps(value)
    if isinstance(value, list):
        return "[" + ", ".join(_toml_value(v) for v in value) + "]"
    if isinstance(value, dict):
        return "{ " + ", ".join(f"{k} = {_toml_value(v)}" for k, v in value.items()) + " }"
    raise TypeError(value)


def config_toml(case: Case) -> str:
    """The case's [config] table, written as the pyprettyvba.toml it stands for."""
    config = case.config
    lines = [f"{key} = {_toml_value(value)}" for key, value in config.items() if key not in ("rules", "overrides")]
    if config.get("rules"):
        lines += ["", "[rules]"]
        lines += [f"{code} = {_toml_value(value)}" for code, value in config["rules"].items()]
    for override in config.get("overrides", []):
        lines += ["", "[[overrides]]"]
        lines += [f"{key} = {_toml_value(value)}" for key, value in override.items()]
    return "\n".join(lines).strip("\n")


def toml_html(text: str) -> str:
    rows = []
    for line in text.splitlines():
        escaped = html.escape(line, quote=False)
        if line.lstrip().startswith("#"):
            rows.append(f'<span class="tc">{escaped}</span>')
        elif re.match(r"^\s*\[", line):
            rows.append(f'<span class="th">{escaped}</span>')
        else:
            match = re.match(r"^(\s*)([\w.-]+)(\s*=)(.*)$", line)
            if match:
                rows.append(
                    f"{html.escape(match.group(1))}<span class=\"tk\">{html.escape(match.group(2))}</span>"
                    f"{html.escape(match.group(3))}{html.escape(match.group(4), quote=False)}"
                )
            else:
                rows.append(escaped)
    return "\n".join(rows)


def terminal_html(text: str) -> str:
    rows = []
    for line in text.splitlines():
        escaped = html.escape(line, quote=False)
        if line.startswith(("+++", "---")):
            rows.append(f'<span class="tm">{escaped}</span>')
        elif line.startswith("+"):
            rows.append(f'<span class="ta">{escaped}</span>')
        elif line.startswith("-"):
            rows.append(f'<span class="tg">{escaped}</span>')
        elif line.startswith("@@"):
            rows.append(f'<span class="tm">{escaped}</span>')
        else:
            rows.append(escaped)
    return "\n".join(rows)


# --------------------------------------------------------------------- cases


def anchor_of(case: Case) -> str:
    return case.id.replace("/", "--")


def parse_check(text: str) -> list[dict[str, str]]:
    found = []
    for line in text.splitlines():
        match = CHECK_RE.match(line)
        if match:
            found.append({k: (v or "") for k, v in match.groupdict().items()})
    return found


def report_html(entries: list[dict[str, str]], anchor: str, *, project: bool) -> str:
    if not entries:
        return '<p class="quiet">Nothing to report: the module is already formatted.</p>'
    rows = []
    for e in entries:
        where = f"{e['line']}:{e['col']}"
        if project and e["file"]:
            where = f"{e['file']}:{where}"
            target = ""
        else:
            target = f' data-target="{anchor}-a-{e["line"]}"'
        badge = '<span class="badge fix">fixed</span>' if e["fix"] else '<span class="badge rep">reported</span>'
        rows.append(
            f"<tr{target}><td class=\"pos\">{html.escape(where)}</td>"
            f"<td><span class=\"rule\">{html.escape(e['rule'])}</span></td>"
            f"<td>{inline(e['msg'])}</td><td>{badge}</td></tr>"
        )
    return (
        '<div class="table-wrap"><table class="report"><thead><tr><th scope="col">Where</th>'
        '<th scope="col">Rule</th><th scope="col">Finding</th><th scope="col"><span class="sr">Status</span>'
        f"</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>"
    )


def pane(title: str, facts: list[str], body: str, cls: str) -> str:
    chips = "".join(f'<span class="fact">{html.escape(f)}</span>' for f in facts)
    return (
        f'<figure class="pane {cls}"><figcaption><span class="pane-title">{html.escape(title)}</span>'
        f"<span class=\"facts\">{chips}</span></figcaption>{body}</figure>"
    )


def module_pair(name: str, source: bytes, output: bytes | None, encoding: str, anchor: str,
                hits: set[int] | None = None) -> str:
    src = decode(source, encoding).text
    out = decode(output, encoding).text if output is not None else src
    changed = output is not None and output != source
    a_marks, b_marks = diff_marks(src, out) if changed else (None, None)
    left = pane(
        f"{name}, as written", file_facts(source, encoding),
        render_code(src, marks=a_marks, side="a", anchor=anchor, lines_hit=hits), "before",
    )
    right_title = f"{name}, formatted" if changed else f"{name}, formatted: unchanged"
    right = pane(
        right_title, file_facts(output if output is not None else source, encoding),
        render_code(out, marks=b_marks, side="b", anchor=anchor), "after" if changed else "after same",
    )
    return f'<div class="pair">{left}{right}</div>'


def rule_counts(entries: list[dict[str, str]]) -> str:
    counts = Counter(e["rule"] for e in entries)
    return "".join(
        f'<span class="chip">{html.escape(rule)} <b>{count}</b></span>' for rule, count in counts.most_common()
    )


def case_html(case: Case) -> str:
    anchor = anchor_of(case)
    config = Config(case.config) if case.cli is None else Config()
    encoding = config.settings_for(case.file).encoding if case.cli is None else "auto"
    parts: list[str] = []
    ws = any(case.id.startswith(prefix) for prefix in WHITESPACE_CASES)
    search = " ".join([case.id, case.title, case.description]).lower()
    toml = config_toml(case)
    blocks: list[str] = []
    entries: list[dict[str, str]] = []
    if case.cli is not None:
        args = [str(a).replace("{file}", case.file) for a in case.cli["args"]]
        command = "pyprettyvba " + " ".join(_shell_quote(a) for a in args)
        stdin = case.cli.get("stdin")
        stdout = (case.path / "stdout.txt").read_bytes().decode("utf-8")
        exit_code = case.cli.get("exit", 0)
        term = [f'<span class="tp">$</span> {html.escape(command)}']
        if stdin:
            term.insert(0, f'<span class="tm"># standard input:</span>\n{html.escape(stdin.rstrip())}')
        term.append(terminal_html(stdout.rstrip("\n")) if stdout.strip() else '<span class="tm">(prints nothing)</span>')
        term_text = "\n".join(term)
        exit_cls = "fact exit" if exit_code else "fact exit ok"
        blocks.append(
            f'<figure class="terminal"><figcaption><span class="pane-title">Terminal</span>'
            f'<span class="facts"><span class="{exit_cls}">exit {exit_code}</span></span>'
            f"</figcaption><pre>{term_text}</pre></figure>"
        )
        extra = case.path / "files"
        if extra.is_dir():
            for path in sorted(p for p in extra.rglob("*") if p.is_file()):
                rel = path.relative_to(extra).as_posix()
                blocks.append(
                    f'<figure class="toml"><figcaption><span class="pane-title">{html.escape(rel)}</span>'
                    f'</figcaption><pre>{toml_html(path.read_text(encoding="utf-8").rstrip())}</pre></figure>'
                )
        if case.input_path().exists():
            output = case.output_path().read_bytes() if case.output_path().exists() else None
            blocks.append(module_pair(case.file, case.input_path().read_bytes(), output, "auto", anchor))
        search += " " + command.lower()
    elif case.project:
        check_text = (case.path / "check.txt").read_bytes().decode("utf-8")
        entries = parse_check(check_text)
        for path in sorted((case.path / "input").iterdir()):
            out_path = case.path / "output" / path.name
            output = out_path.read_bytes() if out_path.exists() else None
            blocks.append(module_pair(path.name, path.read_bytes(), output, encoding, f"{anchor}-{path.stem}"))
    else:
        check_text = (case.path / "check.txt").read_bytes().decode("utf-8")
        entries = parse_check(check_text)
        output = case.output_path().read_bytes() if case.output_path().exists() else None
        blocks.append(module_pair(case.file, case.input_path().read_bytes(), output, encoding, anchor))
    search += " " + " ".join(e["rule"] for e in entries)
    parts.append(
        f'<article class="case" id="{anchor}" data-search="{html.escape(search)}"{" data-ws" if ws else ""}>'
        f'<header class="case-head"><h3><a href="#{anchor}">{inline(case.title)}</a></h3>'
        f'<p class="case-id">tests/fixtures/{html.escape(case.id)}</p></header>'
        f'<div class="case-desc">{paragraphs(case.description)}</div>'
    )
    if toml:
        parts.append(
            '<figure class="toml"><figcaption><span class="pane-title">pyprettyvba.toml</span></figcaption>'
            f"<pre>{toml_html(toml)}</pre></figure>"
        )
    parts.extend(blocks)
    if case.cli is None:
        fixed = sum(1 for e in entries if e["fix"])
        reported = len(entries) - fixed
        summary = []
        if fixed:
            summary.append(f"{fixed} fixed")
        if reported:
            summary.append(f"{reported} reported")
        parts.append(
            f'<details class="findings"{" open" if len(entries) <= 12 else ""}><summary>'
            f'<span class="sum-label">What <code>check</code> reports</span>'
            f'<span class="sum-count">{", ".join(summary) or "nothing"}</span>'
            f'<span class="chips">{rule_counts(entries)}</span></summary>'
            f"{report_html(entries, anchor, project=case.project)}</details>"
        )
    parts.append("</article>")
    return "".join(parts)


def _shell_quote(arg: str) -> str:
    return arg if re.fullmatch(r"[\w./:=,@+-]+", arg) else "'" + arg.replace("'", "'\\''") + "'"


# --------------------------------------------------------------------- page


def build(cases: list[Case]) -> str:
    by_surface: dict[str, list[Case]] = {}
    for case in cases:
        by_surface.setdefault(case.surface, []).append(case)
    summaries = {rule.code: rule.summary for rule in RULES}
    nav: list[str] = []
    sections: list[str] = []
    for key, title, blurb in SURFACES:
        members = by_surface.get(key, [])
        if not members:
            continue
        nav.append(
            f'<li><a href="#{key}"><span>{html.escape(title)}</span><span class="count">{len(members)}</span></a>'
        )
        body: list[str] = []
        if key == "rules":
            groups: dict[str, list[Case]] = {}
            for case in members:
                groups.setdefault(case.id.split("/")[1], []).append(case)
            order = [rule.code for rule in RULES if rule.code in groups]
            sub = []
            for code in order:
                sub.append(f'<li><a href="#rule-{code}">{html.escape(code)}</a></li>')
                body.append(
                    f'<section class="rule-group" id="rule-{code}"><h3 class="rule-name"><code>{html.escape(code)}</code>'
                    f'<span>{inline(summaries.get(code, ""))}</span></h3>'
                    + "".join(case_html(c) for c in groups[code])
                    + "</section>"
                )
            nav.append(f'<ul class="sub">{"".join(sub)}</ul>')
        else:
            body.extend(case_html(case) for case in members)
        nav.append("</li>")
        sections.append(
            f'<section class="surface" id="{key}"><header class="surface-head"><h2>{html.escape(title)}</h2>'
            f'<p>{inline(blurb)}</p></header>{"".join(body)}</section>'
        )
    total = len(cases)
    return PAGE.format(
        title="pyPrettyVBA Fixtures",
        version=html.escape(__version__),
        total=total,
        rules=len(RULES),
        presets=len(PRESETS),
        surfaces=len([s for s in SURFACES if by_surface.get(s[0])]),
        nav="".join(nav),
        sections="".join(sections),
        css=CSS,
        js=JS,
    )


PAGE = """<title>{title}</title>
<style>{css}</style>
<a class="skip" href="#main">Skip to the cases</a>
<div class="shell">
<header class="top">
  <div class="brand">
    <h1>pyPrettyVBA <span>fixtures</span></h1>
    <p class="lede">Every case the formatter's test suite checks, byte for byte: the module as written, its
    configuration, the module as formatted, and what <code>check</code> reports. The suite fails when the
    formatter's output drifts from any of them.</p>
    <p class="stats"><span><b>{total}</b> cases</span><span><b>{rules}</b> rules</span>
    <span><b>{presets}</b> presets</span><span><b>{surfaces}</b> surfaces</span>
    <span>pyPrettyVBA {version}</span></p>
  </div>
  <div class="controls" role="search">
    <label class="search"><span class="sr">Filter the cases</span>
      <input id="filter" type="search" placeholder="Filter by rule, title or text" autocomplete="off" spellcheck="false">
    </label>
    <button type="button" id="ws" class="toggle" aria-pressed="false">Show whitespace and line breaks</button>
    <p class="shown" id="shown" role="status" aria-live="polite"></p>
  </div>
  <p class="legend"><span class="lg k">keyword</span><span class="lg c">comment</span>
  <span class="lg x">not VBA</span><span class="lg before">changed, as written</span>
  <span class="lg after">changed, as formatted</span></p>
</header>
<div class="layout">
<nav class="toc" aria-label="Surfaces"><details open><summary>Contents</summary><ul>{nav}</ul></details></nav>
<main id="main">{sections}
<p class="none" id="none" hidden>No case matches that filter. <button type="button" id="clear">Clear the filter</button></p>
<footer class="foot"><p>Generated from <code>tests/fixtures</code> by <code>tools/build_gallery.py</code>, with
pyPrettyVBA's own lexer coloring the code the way the VBE's code window does. Each case is a test: the output and
report shown here are the ones the suite requires, byte for byte.</p></footer>
</main>
</div>
</div>
<script>{js}</script>
"""

CSS = """
:root {
  --bg: #f3f5f8; --surface: #ffffff; --surface-2: #f7f8fb; --ink: #19202d; --muted: #566174;
  --faint: #8a93a2; --border: #dce1e9; --border-strong: #c5ccd8; --accent: #2141b0; --accent-soft: #e8ecfa;
  --kw: #1b2fa6; --com: #1c7a2c; --bad: #b3261e; --num: #19202d; --str: #19202d;
  --before-line: #fdf2f1; --before-char: #f4c7c1; --after-line: #eef8f0; --after-char: #bfe3c6;
  --ws: #9aa3b3; --trail: #fbd9d4; --hdr: #7b8494; --chip: #edf0f5; --term-bg: #1d2330; --term-ink: #dfe5ee;
  --term-add: #8fdc9c; --term-del: #ff9d94; --term-muted: #8f9aad; --hit: #fff3c4;
  --sans: "Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
  --display: "Segoe UI Variable Display", "Segoe UI Semibold", "Segoe UI", system-ui, -apple-system, sans-serif;
  --mono: "Cascadia Mono", "Cascadia Code", Consolas, "SFMono-Regular", Menlo, "DejaVu Sans Mono", monospace;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --bg: #10141b; --surface: #161b24; --surface-2: #1b212c; --ink: #e2e7ef; --muted: #a3acbb;
    --faint: #6f7888; --border: #283040; --border-strong: #354055; --accent: #93a9ff; --accent-soft: #1f2842;
    --kw: #8fa7ff; --com: #79c483; --bad: #ff8b80; --num: #e2e7ef; --str: #e2e7ef;
    --before-line: rgba(248, 113, 113, 0.09); --before-char: rgba(248, 113, 113, 0.34);
    --after-line: rgba(74, 222, 128, 0.08); --after-char: rgba(74, 222, 128, 0.30);
    --ws: #5b6475; --trail: rgba(248, 113, 113, 0.28); --hdr: #7d8697; --chip: #212836; --term-bg: #0b0e13;
    --term-ink: #dfe5ee; --hit: rgba(250, 204, 21, 0.18);
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --bg: #10141b; --surface: #161b24; --surface-2: #1b212c; --ink: #e2e7ef; --muted: #a3acbb;
  --faint: #6f7888; --border: #283040; --border-strong: #354055; --accent: #93a9ff; --accent-soft: #1f2842;
  --kw: #8fa7ff; --com: #79c483; --bad: #ff8b80; --num: #e2e7ef; --str: #e2e7ef;
  --before-line: rgba(248, 113, 113, 0.09); --before-char: rgba(248, 113, 113, 0.34);
  --after-line: rgba(74, 222, 128, 0.08); --after-char: rgba(74, 222, 128, 0.30);
  --ws: #5b6475; --trail: rgba(248, 113, 113, 0.28); --hdr: #7d8697; --chip: #212836; --term-bg: #0b0e13;
  --term-ink: #dfe5ee; --hit: rgba(250, 204, 21, 0.18);
}
* { box-sizing: border-box; }
body { background: var(--bg); color: var(--ink); font: 15px/1.55 var(--sans); }
a { color: var(--accent); }
code { font-family: var(--mono); font-size: 0.88em; background: var(--chip); padding: 0.05em 0.35em; border-radius: 4px; }
.skip { position: absolute; left: -999px; top: 0; }
.skip:focus { left: 16px; top: 8px; background: var(--surface); padding: 6px 10px; z-index: 10; }
.sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }
.shell { max-width: 1480px; margin: 0 auto; padding-inline: 16px; padding-block: 20px 48px; }
.top { display: grid; gap: 14px; padding-block: 8px 18px; border-bottom: 1px solid var(--border); }
.brand h1 { font: 650 30px/1.15 var(--display); margin: 0 0 6px; letter-spacing: -0.01em; }
.brand h1 span { color: var(--muted); font-weight: 400; }
.lede { margin: 0; max-width: 72ch; color: var(--muted); }
.stats { display: flex; flex-wrap: wrap; gap: 6px 16px; margin: 10px 0 0; color: var(--muted); font-size: 13px; }
.stats b { color: var(--ink); font-variant-numeric: tabular-nums; }
.controls { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
.search input { width: min(360px, 100%); font: inherit; padding: 7px 11px; border-radius: 7px;
  border: 1px solid var(--border-strong); background: var(--surface); color: var(--ink); }
.search { flex: 0 1 360px; min-width: 0; }
.toggle { font: inherit; font-size: 13.5px; padding: 7px 12px; border-radius: 7px; cursor: pointer;
  border: 1px solid var(--border-strong); background: var(--surface); color: var(--ink); }
.toggle[aria-pressed="true"] { background: var(--accent-soft); border-color: var(--accent); color: var(--accent); }
.shown { margin: 0; color: var(--muted); font-size: 13px; }
:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.legend { display: flex; flex-wrap: wrap; gap: 6px 14px; margin: 0; font-size: 12.5px; color: var(--muted); }
.lg::before { content: ""; display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 6px;
  vertical-align: -1px; background: currentColor; }
.lg.k { color: var(--kw); } .lg.c { color: var(--com); } .lg.x { color: var(--bad); }
.lg.before::before { background: var(--before-char); } .lg.after::before { background: var(--after-char); }
.lg.before, .lg.after { color: var(--muted); }
.layout { display: grid; grid-template-columns: 220px minmax(0, 1fr); gap: 28px; margin-top: 22px; }
.toc { position: sticky; top: calc(env(safe-area-inset-top, 0px) + 12px); align-self: start;
  max-height: calc(100vh - 24px); overflow-y: auto; font-size: 13.5px; }
.toc summary { display: none; }
.toc ul { list-style: none; margin: 0; padding: 0; }
.toc > details > ul > li > a { display: flex; justify-content: space-between; padding: 5px 8px; border-radius: 6px;
  color: var(--ink); text-decoration: none; font-weight: 600; }
.toc a:hover { background: var(--chip); }
.toc .count { color: var(--faint); font-weight: 400; font-variant-numeric: tabular-nums; }
.toc .sub { margin: 2px 0 8px 8px; border-left: 1px solid var(--border); }
.toc .sub a { display: block; padding: 2px 10px; color: var(--muted); text-decoration: none; font-family: var(--mono);
  font-size: 12px; }
.toc .sub a:hover { color: var(--accent); }
main { min-width: 0; }
.surface { margin-bottom: 40px; }
.surface-head { border-bottom: 2px solid var(--ink); padding-bottom: 6px; margin-bottom: 18px; }
.surface-head h2 { font: 650 22px/1.2 var(--display); margin: 0; }
.surface-head p { margin: 4px 0 0; color: var(--muted); max-width: 80ch; }
.rule-group { margin-bottom: 26px; }
.rule-name { display: flex; flex-wrap: wrap; gap: 4px 12px; align-items: baseline; font: 600 17px/1.3 var(--display);
  margin: 26px 0 10px; }
.rule-name code { font-size: 15px; color: var(--accent); background: var(--accent-soft); }
.rule-name span { font-weight: 400; color: var(--muted); font-size: 15px; }
.case { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px 18px 14px;
  margin-bottom: 16px; display: grid; gap: 12px; }
.case-head h3 { margin: 0; font: 600 16.5px/1.35 var(--display); text-wrap: balance; }
.case-head h3 a { color: var(--ink); text-decoration: none; }
.case-head h3 a:hover { text-decoration: underline; }
.case-id { margin: 2px 0 0; font: 12px var(--mono); color: var(--faint); }
.case-desc p { margin: 0 0 6px; max-width: 78ch; }
.case-desc p:last-child { margin-bottom: 0; }
figure { margin: 0; }
figcaption { display: flex; flex-wrap: wrap; gap: 4px 10px; align-items: baseline; justify-content: space-between;
  padding: 6px 10px; border-bottom: 1px solid var(--border); background: var(--surface-2); font-size: 12.5px; }
.pane-title { font-weight: 600; color: var(--ink); }
.facts { display: flex; flex-wrap: wrap; gap: 6px; }
.fact { font: 11.5px var(--mono); color: var(--muted); background: var(--chip); padding: 1px 6px; border-radius: 4px; }
.fact.exit { color: var(--bad); } .fact.exit.ok { color: var(--com); }
.pair { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
.pane, .toml, .terminal { border: 1px solid var(--border); border-radius: 8px; overflow: hidden; background: var(--surface); }
.pane.after.same .pane-title { color: var(--muted); font-weight: 500; }
.code { overflow-x: auto; font: 12.8px/1.62 var(--mono); tab-size: 4; }
.rows { display: inline-block; min-width: 100%; padding-block: 6px; }
.ln { display: flex; min-height: 1.62em; }
.ln .no { flex: none; width: 3.6ch; padding-right: 1ch; text-align: right; color: var(--faint);
  user-select: none; font-variant-numeric: tabular-nums; }
.ln .tx { white-space: pre; padding-right: 12px; flex: 1 0 auto; }
.ln.hdr .tx { color: var(--hdr); }
.ln.hdr .tx .k, .ln.hdr .tx .s, .ln.hdr .tx .c { color: var(--hdr); }
.ln.empty .tx { color: var(--faint); font-style: italic; font-family: var(--sans); }
.before .ln.chg, .before .ln.gone { background: var(--before-line); }
.after .ln.chg, .after .ln.new { background: var(--after-line); }
.before .d, .before .ln.gone .tx { background-color: var(--before-char); }
.after .d, .after .ln.new .tx { background-color: var(--after-char); }
.ln.new .tx:empty::after, .ln.gone .tx:empty::after { content: " "; }
.ln.hit { background: var(--hit) !important; }
.k { color: var(--kw); } .c { color: var(--com); } .x { color: var(--bad); text-decoration: underline wavy; }
.s { color: var(--str); } .n { color: var(--num); }
.raw { font-size: 0.78em; color: var(--bad); border: 1px solid currentColor; border-radius: 3px; padding: 0 2px;
  margin: 0 1px; vertical-align: 1px; }
.show-ws .w, .case[data-ws] .w {
  background-image: radial-gradient(circle at 50% 55%, var(--ws) 0 1.3px, transparent 1.7px);
  background-size: 1ch 100%; background-repeat: repeat-x; }
.show-ws .t, .case[data-ws] .t {
  background-image: linear-gradient(var(--ws), var(--ws)); background-size: calc(100% - 0.6ch) 1px;
  background-position: 0.3ch 58%; background-repeat: no-repeat; }
.show-ws .tr, .case[data-ws] .tr { background-color: var(--trail); }
.eol::after { font: 9.5px/1 var(--sans); color: var(--ws); margin-left: 3px; vertical-align: 1px;
  letter-spacing: 0.04em; }
.show-ws .eol.crlf::after, .case[data-ws] .eol.crlf::after { content: "CRLF"; }
.show-ws .eol.lf::after, .case[data-ws] .eol.lf::after { content: "LF"; }
.show-ws .eol.cr::after, .case[data-ws] .eol.cr::after { content: "CR"; }
.show-ws .eol.none::after, .case[data-ws] .eol.none::after { content: "no line break"; font-style: italic; }
.eol.d::after { color: var(--ink); background: var(--before-char); padding: 1px 3px; border-radius: 3px; }
.after .eol.d::after { background: var(--after-char); }
.eol.d.crlf::after { content: "CRLF"; } .eol.d.lf::after { content: "LF"; } .eol.d.cr::after { content: "CR"; }
.toml pre, .terminal pre { margin: 0; padding: 10px 12px; overflow-x: auto; font: 12.8px/1.6 var(--mono); }
.th { color: var(--accent); font-weight: 600; } .tk { color: var(--ink); font-weight: 600; } .tc { color: var(--faint); }
.terminal { background: var(--term-bg); border-color: var(--term-bg); }
.terminal figcaption { background: var(--term-bg); border-color: rgba(255, 255, 255, 0.08); }
.terminal .pane-title { color: var(--term-muted); }
.terminal .fact { background: rgba(255, 255, 255, 0.08); }
.terminal .fact.exit { color: var(--term-del); } .terminal .fact.exit.ok { color: var(--term-add); }
.terminal pre { color: var(--term-ink); }
.tp { color: var(--term-muted); user-select: none; } .ta { color: var(--term-add); } .tg { color: var(--term-del); }
.tm { color: var(--term-muted); }
.findings { border-top: 1px dashed var(--border); padding-top: 8px; }
.findings summary { cursor: pointer; display: flex; flex-wrap: wrap; gap: 6px 12px; align-items: center;
  font-size: 13.5px; list-style: none; }
.findings summary::-webkit-details-marker { display: none; }
.findings summary::before { content: ""; width: 7px; height: 7px; border-right: 2px solid var(--muted);
  border-bottom: 2px solid var(--muted); transform: rotate(-45deg); transition: transform 0.15s; margin-right: 2px; }
.findings[open] summary::before { transform: rotate(45deg); }
.sum-label { font-weight: 600; }
.sum-count { color: var(--muted); }
.chips { display: flex; flex-wrap: wrap; gap: 5px; }
.chip { font: 11.5px var(--mono); background: var(--chip); color: var(--muted); padding: 1px 7px; border-radius: 10px; }
.chip b { color: var(--ink); font-weight: 600; }
.table-wrap { overflow-x: auto; margin-top: 8px; }
.report { border-collapse: collapse; width: 100%; font-size: 13px; }
.report th { text-align: left; font-weight: 600; color: var(--muted); font-size: 12px; padding: 4px 8px;
  border-bottom: 1px solid var(--border); }
.report td { padding: 4px 8px; border-bottom: 1px solid var(--border); vertical-align: top; }
.report tr[data-target] { cursor: pointer; }
.report tr[data-target]:hover td { background: var(--surface-2); }
.report .pos { font: 12px var(--mono); color: var(--muted); white-space: nowrap; font-variant-numeric: tabular-nums; }
.rule { font: 12px var(--mono); color: var(--accent); white-space: nowrap; }
.badge { font-size: 11px; padding: 1px 7px; border-radius: 10px; white-space: nowrap; }
.badge.fix { background: var(--after-line); color: var(--com); border: 1px solid var(--after-char); }
.badge.rep { background: var(--before-line); color: var(--bad); border: 1px solid var(--before-char); }
.quiet { margin: 6px 0 0; color: var(--muted); font-size: 13.5px; }
.none { padding: 24px; text-align: center; color: var(--muted); }
.none button { font: inherit; color: var(--accent); background: none; border: 0; text-decoration: underline; cursor: pointer; }
.foot { border-top: 1px solid var(--border); margin-top: 28px; padding-top: 12px; color: var(--muted); font-size: 13px; }
.foot p { max-width: 90ch; margin: 0; }
[hidden] { display: none !important; }
@media (max-width: 1080px) { .pair { grid-template-columns: minmax(0, 1fr); } }
@media (max-width: 860px) {
  .layout { grid-template-columns: minmax(0, 1fr); gap: 12px; }
  .toc { position: static; max-height: none; }
  .toc summary { display: block; cursor: pointer; font-weight: 600; padding: 6px 0; }
  .toc .sub { display: none; }
  .case { padding: 12px; }
}
@media (prefers-reduced-motion: reduce) { * { transition: none !important; scroll-behavior: auto !important; } }
"""

JS = """
(function () {
  var root = document.documentElement;
  var wsButton = document.getElementById('ws');
  function setWs(on) {
    root.classList.toggle('show-ws', on);
    wsButton.setAttribute('aria-pressed', on ? 'true' : 'false');
    try { localStorage.setItem('pyprettyvba-gallery-ws', on ? '1' : '0'); } catch (e) {}
  }
  var saved = null;
  try { saved = localStorage.getItem('pyprettyvba-gallery-ws'); } catch (e) {}
  setWs(saved === '1');
  wsButton.addEventListener('click', function () { setWs(!root.classList.contains('show-ws')); });

  var cases = Array.prototype.slice.call(document.querySelectorAll('.case'));
  var groups = Array.prototype.slice.call(document.querySelectorAll('.rule-group, .surface'));
  var filter = document.getElementById('filter');
  var shown = document.getElementById('shown');
  var none = document.getElementById('none');
  function apply() {
    var words = filter.value.toLowerCase().split(/\\s+/).filter(Boolean);
    var count = 0;
    cases.forEach(function (c) {
      var text = c.getAttribute('data-search');
      var match = words.every(function (w) { return text.indexOf(w) !== -1; });
      c.hidden = !match;
      if (match) { count += 1; }
    });
    groups.forEach(function (g) { g.hidden = !g.querySelector('.case:not([hidden])'); });
    shown.textContent = words.length ? count + ' of ' + cases.length + ' cases' : '';
    none.hidden = count !== 0;
  }
  filter.addEventListener('input', apply);
  document.getElementById('clear').addEventListener('click', function () { filter.value = ''; apply(); filter.focus(); });

  document.addEventListener('click', function (event) {
    var row = event.target.closest('tr[data-target]');
    if (!row) { return; }
    var line = document.getElementById(row.getAttribute('data-target'));
    if (!line) { return; }
    document.querySelectorAll('.ln.hit').forEach(function (l) { l.classList.remove('hit'); });
    line.classList.add('hit');
    var box = line.closest('.code');
    var top = line.offsetTop - box.clientHeight / 3;
    box.scrollTop = Math.max(0, top);
    line.scrollIntoView({ block: 'nearest' });
  });
})();
"""


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args(argv)
    cases = load_cases(FIXTURES)
    page = build(cases)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(page, encoding="utf-8", newline="\n")
    print(f"wrote {args.out} ({len(page) // 1024} KB, {len(cases)} cases)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
