"""Checks against a real VBE. Windows, desktop Excel and pyVBAharness only.

    python -m pytest -m live tests/test_live.py

* The VBE has nothing left to change in the `vbe` preset's output: each
  fixture input, formatted with that preset, written to a file and read in
  with VBComponents.Import (the way an exported module goes back), exports
  exactly as it went in, line endings aside (the VBE writes CRLF).
* Formatting keeps a module's compile outcome: the VBE accepts the formatted
  module exactly when it accepts the original, with the same message.

Each module goes into a workbook of its own, so no name one module declares
can change the casing of another. The round trip runs from VBA in the
harness's workbook but imports into a new workbook it creates, because the
harness's own support module declares names (`count`, `value`) that the
VBE's project-wide name table would otherwise apply to the module.
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from typing import Any

import pytest

from pyprettyvba import Config, format_source
from pyprettyvba.document import split_header
from pyprettyvba.textio import decode

from fixture_support import Case, load_cases

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(sys.platform != "win32", reason="needs desktop Excel on Windows"),
]

VBE = Config({"preset": "vbe"})
KINDS = {".bas": "standard", ".cls": "class"}
CASES = [
    case
    for case in load_cases()
    if case.cli is None and not case.project and Path(case.file).suffix.lower() in KINDS
]
# Where the `vbe` preset does not write what the VBE writes, on purpose.
VALUE_CHANGE = "the VBE's spelling of this literal is a different value; the preset reports it instead"
DEVIATIONS: dict[str, str] = {
    "rules/date-literals/basic": "the VBE reads a two-digit year through the machine's date window",
    "rules/numeric-literals/basic": VALUE_CHANGE,
    "rules/numeric-literals/quiet": VALUE_CHANGE,
    "safety/precision": VALUE_CHANGE,
}
# Cases whose formatting touches something the compiler could notice:
# statement shapes, literal spellings, keyword forms, whole presets.
COMPILE_CASES = [
    case
    for case in CASES
    if case.surface in ("presets", "showcase", "safety", "files")
    or case.id.split("/")[1] in (
        "statement-form", "split-statements", "let-keyword", "numeric-literals",
        "date-literals", "keyword-case", "spacing", "rem-comments",
    )
]
ROUND_TRIP = (
    "Public Function ZqRoundTrip(ByVal zqIn As String, ByVal zqOut As String) As String\r\n"
    "    Dim zqBook As Object, zqComp As Object\r\n"
    "    Set zqBook = Workbooks.Add\r\n"
    "    Set zqComp = zqBook.VBProject.VBComponents.Import(zqIn)\r\n"
    "    zqComp.Export zqOut\r\n"
    "    ZqRoundTrip = zqComp.Name\r\n"
    "    zqBook.Close False\r\n"
    "End Function\r\n"
)


@pytest.fixture(scope="module")
def excel() -> Any:
    harness = pytest.importorskip("pyvbaharness")
    # Office automation is one session at a time: wait for a turn
    # (PYPRETTYVBA_LIVE_LOCK_WAIT seconds, 300 by default) before giving up.
    wait = float(os.environ.get("PYPRETTYVBA_LIVE_LOCK_WAIT", "300"))
    try:
        session = harness.ExcelSession(harness.HarnessConfig(lock_wait_s=wait))
    except harness.SessionLockHeld:
        pytest.skip(f"another pyVBAharness session held Excel for {wait:.0f} seconds")
    with session:
        session.new_workbook()
        yield session


def _formatted(case: Case, config: Config) -> tuple[str, str]:
    """The input's text and its formatted text."""
    text = decode(case.input_path().read_bytes(), config.settings_for(case.file).encoding).text
    return text, format_source(text, config, path=case.file).output


def _cp1252(text: str) -> bytes:
    try:
        return text.encode("cp1252")
    except UnicodeEncodeError:
        pytest.skip("the VBE stores modules in the ANSI code page, which cannot hold this text")


def _round_trip(excel: Any, case: Case, text: str) -> str:
    """Import ``text`` as a module file into a new workbook and export it again."""
    excel.add_module("ZqRoundTrip", ROUND_TRIP)
    with tempfile.TemporaryDirectory() as tmp:
        source = Path(tmp) / Path(case.file).name
        target = Path(tmp) / ("out" + source.suffix)
        source.write_bytes(_cp1252(text))
        result = excel.run_macro("ZqRoundTrip.ZqRoundTrip", str(source), str(target), timeout=60)
        assert result.outcome == "passed", result
        return target.read_bytes().decode("cp1252")


def _compile(excel: Any, case: Case, text: str) -> tuple[str, str]:
    excel.new_workbook()
    excel.add_module("ZqLive", split_header(text)[1], kind=KINDS[Path(case.file).suffix.lower()])
    result = excel.compile_project(watch_seconds=20)
    excel.remove_module("ZqLive")
    return result.outcome, result.message


def _lines(text: str) -> list[str]:
    return split_header(text)[1].replace("\r\n", "\n").replace("\r", "\n").split("\n")


@pytest.mark.parametrize("case", CASES, ids=[c.id for c in CASES])
def test_the_vbe_leaves_the_vbe_preset_output_alone(excel: Any, case: Case) -> None:
    if case.surface == "suppression":
        pytest.skip("suppressed lines are left as written on purpose")
    _source, formatted = _formatted(case, VBE)
    if not split_header(formatted)[1].strip():
        pytest.skip("an empty module")
    exported = _round_trip(excel, case, formatted)
    if case.id in DEVIATIONS:
        if _lines(exported) == _lines(formatted):
            pytest.fail(f"{case.id} now matches the VBE; remove it from DEVIATIONS")
        pytest.skip(DEVIATIONS[case.id])
    assert _lines(exported) == _lines(formatted)


@pytest.mark.parametrize("case", COMPILE_CASES, ids=[c.id for c in COMPILE_CASES])
def test_formatting_keeps_the_compile_outcome(excel: Any, case: Case) -> None:
    source, formatted = _formatted(case, Config(case.config))
    if formatted == source:
        pytest.skip("formatting changes nothing")
    before = _compile(excel, case, source)
    after = _compile(excel, case, formatted)
    assert "infrastructure-failure" not in (before[0], after[0]), (before, after)
    assert after == before
