"""Properties that hold for every input, checked with Hypothesis.

The generators are in tests/vba_strategies.py. Set HYPOTHESIS_PROFILE=thorough
for a longer run (see tests/conftest.py).
"""

from __future__ import annotations

import pytest
from hypothesis import given
from hypothesis import strategies as st

from pyprettyvba import PRESETS, Config, format_source
from pyprettyvba.chars import is_wsc
from pyprettyvba.document import Document
from pyprettyvba.lexer import tokenize
from pyprettyvba.structure import analyze_structure

from vba_strategies import GeneratedModule, canonical_modules, token_soup

ANY_TEXT = st.one_of(st.text(), token_soup())
CONFIGS = {preset: Config({"preset": preset}) for preset in PRESETS}


@given(ANY_TEXT)
def test_the_lexer_loses_nothing(text: str) -> None:
    tokens = tokenize(text)
    assert "".join(tok.text for tok in tokens) == text
    position = 0
    for tok in tokens:
        assert tok.start == position and tok.end == tok.start + len(tok.text) and tok.text
        position = tok.end


@given(ANY_TEXT)
def test_every_text_has_a_document_and_a_structure(text: str) -> None:
    doc = Document(text)
    structure = analyze_structure(doc)
    assert len(structure.levels) == len(doc.lines)


def _non_ascii(text: str) -> list[str]:
    return [c for c in text if ord(c) > 127 and not is_wsc(c)]


@pytest.mark.parametrize("preset", PRESETS)
@given(text=token_soup())
def test_every_preset_is_safe_and_settles(preset: str, text: str) -> None:
    config = CONFIGS[preset]
    # format_source refuses (raises) output that changes the code's meaning.
    result = format_source(text, config)
    again = format_source(result.output, config)
    assert again.output == result.output
    assert [v for v in again.violations if v.fixable] == []
    # Formatting changes ASCII characters and whitespace, nothing else.
    assert _non_ascii(result.output) == _non_ascii(text)
    lines = text.count("\n") + text.count("\r") + 1
    for violation in result.violations:
        assert 1 <= violation.line <= lines and violation.column >= 1


@given(text=token_soup())
def test_the_none_preset_changes_nothing(text: str) -> None:
    result = format_source(text, CONFIGS["none"])
    assert result.output == text
    # It still checks the suppression directives.
    assert {v.rule for v in result.violations} <= {"suppression-directive"}


@given(canonical_modules())
def test_scrambled_modules_format_back_to_the_canonical_form(module: GeneratedModule) -> None:
    assert format_source(module.scrambled).output == module.canonical


@given(canonical_modules())
def test_the_canonical_form_has_nothing_to_fix(module: GeneratedModule) -> None:
    assert format_source(module.canonical).violations == []
