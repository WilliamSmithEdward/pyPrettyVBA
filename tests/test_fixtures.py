"""Run every fixture case (see tests/fixture_support.py for the layout)."""

from __future__ import annotations

import pytest

from pyprettyvba import Config, format_source
from pyprettyvba.textio import decode

from fixture_support import Case, load_cases, run_case

CASES = load_cases()


def test_there_are_cases_for_every_surface() -> None:
    surfaces = {case.surface for case in CASES}
    assert {"rules", "presets", "suppression", "config", "files", "safety", "cli", "project", "showcase"} <= surfaces


def test_every_rule_has_a_case() -> None:
    from pyprettyvba import RULES

    covered = {case.path.relative_to(case.path.parents[1]).parts[0] for case in CASES if case.surface == "rules"}
    missing = {rule.code for rule in RULES} - covered - {"suppression-directive"}
    assert not missing, f"rules without a fixture: {sorted(missing)}"


@pytest.mark.parametrize("case", CASES, ids=[c.id for c in CASES])
def test_case(case: Case) -> None:
    outcome = run_case(case)
    if case.cli is not None:
        assert outcome.stdout == (case.path / "stdout.txt").read_bytes().decode("utf-8")
        assert outcome.exit_code == case.cli.get("exit", 0)
        expected_output = case.output_path()
        if expected_output.exists():
            assert outcome.output == expected_output.read_bytes()
        return
    assert outcome.check == (case.path / "check.txt").read_bytes().decode("utf-8")
    if case.project:
        for name, data in outcome.outputs.items():
            expected = case.path / "output" / name
            source = case.path / "input" / name
            assert data == (expected.read_bytes() if expected.exists() else source.read_bytes()), name
        return
    expected = case.output_path()
    assert outcome.output == (expected.read_bytes() if expected.exists() else case.input_path().read_bytes())


@pytest.mark.parametrize(
    "case", [c for c in CASES if c.cli is None and not c.project], ids=[c.id for c in CASES if c.cli is None and not c.project]
)
def test_case_output_is_stable(case: Case) -> None:
    """Formatting a case's output again changes nothing and fixes nothing."""
    config = Config(case.config)
    path = case.output_path() if case.output_path().exists() else case.input_path()
    text = decode(path.read_bytes(), config.settings_for(case.file).encoding).text
    again = format_source(text, config, path=case.file)
    assert again.output == text
    assert [v for v in again.violations if v.fixable] == []
