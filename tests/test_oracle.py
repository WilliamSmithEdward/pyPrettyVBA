"""Replay the VBE's own rendering: the `vbe` preset must write what the VBE wrote.

tests/oracle/vbe_rendering.json records probes pasted into a real VBE
(CodeModule.AddFromString) and exported again (tools/vbe_oracle.py). Every
probe is formatted with the `vbe` preset and compared with the VBE's output,
line for line. The few places pyPrettyVBA deliberately does something else
are listed in DEVIATIONS with the reason; a new mismatch fails the suite.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyprettyvba import Config, format_source

ORACLE = Path(__file__).parent / "oracle" / "vbe_rendering.json"
EVIDENCE: dict[str, Any] = json.loads(ORACLE.read_text(encoding="utf-8"))

# The VBE keeps one spelling per name for the whole project. The recording's
# project also held pyVBAharness's support module, which declares `item`,
# `count`, `text` and `value`, so the VBE spelled those names that way here;
# the `vbe` preset, which sees one module, spells them as Excel's library
# does (`.Item`, `.Count`), as the VBE does in a project without it
# (tests/test_live.py).
HARNESS_NAMES = "the harness module in the recorded project declares item, count, text and value"
# A line the VBE cannot parse is stored exactly as written; pyPrettyVBA
# still spaces the parts it can read, which cannot change a line that does
# not compile.
UNPARSED = "the VBE stores a line it cannot parse as written"

DEVIATIONS: dict[str, str] = {
    "p-with": HARNESS_NAMES,
    "p-member-chain": HARNESS_NAMES,
    "p-member-space-before-dot": HARNESS_NAMES,
    "r-with-paren": HARNESS_NAMES,
    # Also: a keyword used as a member name (`zqo.Print`, `zqo.input`) took
    # the spelling that word was last typed with anywhere in the module, even
    # in keyword position (`debug.print`, `LINE INPUT`), where the VBE shows
    # the keyword's own spelling. The preset leaves such member names alone.
    "k-member-keywords": HARNESS_NAMES + "; member keywords take the word's last typed spelling",
    "p-concat-glued-vars": UNPARSED,
    "p-concat-glued-string": UNPARSED,
    "s-optional-params": UNPARSED,
    "r-call-empty-parens": UNPARSED + " (`zqfoo()` without Call)",
    # A comment one space after its code stays one space after it when the
    # code shrinks; the VBE keeps the column. Both are stable in the VBE.
    "r-comment-after-collapse": "a comment attached to its code moves with it",
    # A two-digit year depends on the machine's date window.
    "d-two-digit-year": "a two-digit year is machine-dependent",
    # The VBE rounds these to 15 (or 7) digits, changing the value; they are
    # reported instead.
    "n-float-big2": "the VBE's spelling changes the value",
    "n-float-16": "the VBE's spelling changes the value",
    "n-float-pi17": "the VBE's spelling changes the value",
    "n-suffix-double": "the VBE's spelling changes the value",
    "n-suffix-single": "the VBE's spelling changes the value",
}

CASES = [c for c in EVIDENCE["rendering"] if c["scope"] != "export" and c["output"] is not None]
CONFIG = Config({"preset": "vbe", "rules": {"end-of-file": False, "line-endings": False}})
MARKER = "'@@case "


def _module_text(module: str) -> str:
    """The probe module exactly as tools/vbe_oracle.py assembled it."""
    probes = [c for c in EVIDENCE["rendering"] if c["module"] == module and c["scope"] != "export"]
    prefix = "ZqQ" if module == "class" else "ZqP"
    head: list[str] = []
    tail: list[str] = []
    for index, case in enumerate(probes, start=1):
        marker = MARKER + case["id"]
        if case["scope"] == "module":
            head += [marker, *case["input"]]
        elif case["scope"] == "body":
            tail += [marker, f"Sub {prefix}{index:03d}()", *case["input"], "End Sub"]
        else:
            tail += [marker, *case["input"]]
    tail.append(MARKER + "@@end")
    return "\r\n".join(head + tail) + "\r\n"


def _rendered_module(module: str) -> dict[str, list[str]]:
    """Format a probe module once and cut the result back into probes."""
    path = "Probe.cls" if module == "class" else "Probe.bas"
    output = format_source(_module_text(module), CONFIG, path=path).output
    chunks: dict[str, list[str]] = {}
    current: str | None = None
    for line in output.split("\r\n"):
        if line.startswith(MARKER):
            current = line[len(MARKER):]
            chunks[current] = []
        elif current is not None:
            chunks[current].append(line)
    by_id = {c["id"]: c for c in CASES}
    for case_id, chunk in chunks.items():
        case = by_id.get(case_id)
        if case is None:
            continue
        while chunk and chunk[-1] == "" and case["input"][-1] != "":
            chunk.pop()
        if case["scope"] == "body":
            chunks[case_id] = chunk[1:-1]
    return chunks


RENDERED = {**_rendered_module("standard"), **_rendered_module("class")}


@pytest.mark.parametrize("case", CASES, ids=[c["id"] for c in CASES])
def test_vbe_rendering(case: dict[str, Any]) -> None:
    rendered = RENDERED[case["id"]]
    expected = case["output"]
    if case["id"] in DEVIATIONS:
        if rendered == expected:
            pytest.fail(f"{case['id']} now matches the VBE; remove it from DEVIATIONS")
        pytest.skip(DEVIATIONS[case["id"]])
    assert rendered == expected


def test_every_deviation_names_a_recorded_case() -> None:
    ids = {c["id"] for c in CASES}
    assert set(DEVIATIONS) <= ids


# Modules the VBE rendered in a project of their own, for the questions of
# how names interact (contextual keywords used as names, a Type member named
# like a library function, a name declared twice, the libraries' names) and
# of blank lines at either end. Most were read in with AddFromString, which
# stores one more (empty) line after the text's last line break; the rest
# were imported from a file (VBComponents.Import), which does not.
ISOLATED = EVIDENCE.get("isolated", [])
ISOLATED_DEVIATIONS: dict[str, str] = {
    # AddFromString drops blank lines above the first line of code, but
    # Import keeps them (f-import-leading-blank-lines); the preset follows
    # Import, which is how a formatted file goes back into the VBE.
    "i-leading-blank-lines": "AddFromString drops leading blank lines; Import keeps them",
}
WHOLE_MODULE_CONFIG = Config({"preset": "vbe", "rules": {"line-endings": False}})


@pytest.mark.parametrize("case", ISOLATED, ids=[c["id"] for c in ISOLATED])
def test_vbe_rendering_of_whole_modules(case: dict[str, Any]) -> None:
    source = "\r\n".join(case["input"]) + "\r\n"
    output = format_source(source, WHOLE_MODULE_CONFIG, path="Probe.bas").output
    imported = case.get("method") == "VBComponents.Import"
    rendered = output.splitlines() if imported else output.split("\r\n")
    if case["id"] in ISOLATED_DEVIATIONS:
        if rendered == case["output"]:
            pytest.fail(f"{case['id']} now matches the VBE; remove it from ISOLATED_DEVIATIONS")
        pytest.skip(ISOLATED_DEVIATIONS[case["id"]])
    assert rendered == case["output"]


COMPILE = {c["id"]: c["outcome"] for c in EVIDENCE["compile"]}
RUNTIME = {c["id"]: c["value"] for c in EVIDENCE["runtime"]}


def test_recorded_language_facts_the_formatter_relies_on() -> None:
    # Each of these decides a rule's behavior; if a re-recording changes one,
    # the rule it names needs revisiting.
    assert COMPILE["if-then-colon-opens-block"] == "rejected"  # document: If x Then: is single-line
    assert RUNTIME["if-then-colon-is-single-line"] == 5
    assert RUNTIME["single-line-if-colon-scope"] == 0  # split-statements never splits a single-line If
    assert RUNTIME["single-line-else-colon-scope"] == 1
    assert COMPILE["comment-continues"] == "accepted"  # lexer: comments continue across ` _`
    assert RUNTIME["comment-continuation-hides-statement"] == 1
    assert RUNTIME["rem-continuation-hides-statement"] == 1
    assert COMPILE["directive-comment-continues"] == "accepted"
    assert RUNTIME["continuation-with-trailing-space"] == 3  # lexer: ` _   ` continues
    assert COMPILE["concat-glued-identifiers"] == "rejected"  # spacing: glued `&` is a type character
    assert COMPILE["bang-with-spaces"] == "rejected"  # spacing: a bang stays glued
    assert COMPILE["with-member-argument"] == "accepted"  # spacing: `Foo .Bar` keeps its space
    assert COMPILE["call-omitted-first-glued"] == COMPILE["call-omitted-first-spaced"] == "rejected"
    assert RUNTIME["call-omitted-first-glued-runs"] == RUNTIME["call-omitted-first-spaced-runs"] == 102
    assert COMPILE["greater-equal-reversed"] == "accepted"  # spacing: => is >=
    assert COMPILE["endif-single-word"] == "accepted"  # keyword-case: EndIf is End If
    assert COMPILE["indented-label"] == "accepted"
