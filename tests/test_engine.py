"""The pipeline: positions, idempotency, safety, reporting."""

from __future__ import annotations

from collections.abc import Iterable

import pytest

from pyprettyvba import Config, SafetyError, format_source
from pyprettyvba.document import Document
from pyprettyvba.engine import UnstableFormattingError, run_pipeline
from pyprettyvba.rules.base import Finding, FormatContext, Rule


def test_violations_point_at_the_original_source() -> None:
    src = "Attribute VB_Name = \"M\"\r\nsub a()\r\n\r\n\r\n\r\nx=1\r\nend sub\r\n"
    result = format_source(src)
    spacing = [v for v in result.violations if v.rule == "spacing"]
    # `x=1` is on line 6 of the file as written, whatever blank lines went.
    assert {(v.line, v.column) for v in spacing} == {(6, 2), (6, 3)}
    blank = [v for v in result.violations if v.rule == "blank-lines"]
    assert blank[0].line == 3


def test_formatting_is_idempotent_on_its_own_output() -> None:
    src = "sub a()\r\nif x then\r\ny=1:z=2\r\nend if\r\nend sub\r\n"
    once = format_source(src).output
    twice = format_source(once)
    assert twice.output == once and twice.violations == []


def test_header_is_never_formatted() -> None:
    src = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\nsub a()\r\nend sub\r\n"
    out = format_source(src, path="C.cls").output
    assert out.startswith("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n")


def test_header_line_endings_follow_the_body() -> None:
    src = 'Attribute VB_Name = "M"\nSub A()\r\nEnd Sub\r\n'
    result = format_source(src, Config({"line-ending": "crlf"}))
    assert result.output == 'Attribute VB_Name = "M"\r\nSub A()\r\nEnd Sub\r\n'
    assert [v.rule for v in result.violations] == ["line-endings"]


class _Breaker(Rule):
    code = "breaker"
    summary = "Deliberately changes the meaning of the code."
    category = "test"

    def run(self, doc: Document) -> Iterable[Finding]:
        for tok in doc.tokens:
            if tok.text == "1":
                yield Finding(tok.start, tok.end, "2", "bump")


class _Flipper(Rule):
    code = "flipper"
    summary = "Never settles."
    category = "test"

    def run(self, doc: Document) -> Iterable[Finding]:
        yield Finding(0, 0, " ", "grow")


def test_output_that_changes_meaning_is_refused() -> None:
    ctx = FormatContext()
    with pytest.raises(SafetyError, match="meaning"):
        run_pipeline("x = 1\n", [_Breaker({}, ctx)], ctx, frozenset({"breaker"}))


def test_rules_that_never_settle_are_an_error() -> None:
    ctx = FormatContext()
    with pytest.raises(UnstableFormattingError):
        run_pipeline("x = 1\n", [_Flipper({}, ctx)], ctx, frozenset({"flipper"}), check_safety=False)


def test_grouped_findings_are_reported_once() -> None:
    result = format_source("Sub A()\nx = 1\r\ny = 2\nEnd Sub\n", Config({"line-ending": "crlf"}))
    assert [v.rule for v in result.violations].count("line-endings") == 1


def test_report_only_rules_do_not_change_text() -> None:
    src = "x = 3.141592653589793\r\n"
    result = format_source(src)
    assert result.output == src
    assert [(v.rule, v.fixable) for v in result.violations] == [("numeric-literals", False)]
