"""The fuzz corpus, replayed through the fuzz targets on every run.

tests/fuzz_corpus/vba seeds fuzz/fuzz_formatter.py; a fuzz finding joins it
as a regression seed. Each seed's first byte picks a preset, as in the
format target, and the rest is the source.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyprettyvba import PRESETS, Config, format_source
from pyprettyvba.chars import is_wsc
from pyprettyvba.lexer import tokenize

CORPUS = Path(__file__).parent / "fuzz_corpus" / "vba"
SEEDS = sorted(CORPUS.glob("*"))


def _non_ascii(text: str) -> list[str]:
    return [c for c in text if ord(c) > 127 and not is_wsc(c)]


def test_the_corpus_is_present() -> None:
    assert len(SEEDS) >= 50


@pytest.mark.parametrize("path", SEEDS, ids=lambda p: p.name)
def test_a_seed_lexes_and_formats_safely(path: Path) -> None:
    data = path.read_bytes()
    preset = PRESETS[data[0] % len(PRESETS)]
    text = data[1:].decode("utf-8", errors="replace")
    assert "".join(token.text for token in tokenize(text)) == text
    config = Config({"preset": preset})
    result = format_source(text, config)
    if preset == "none":
        assert result.output == text
        return
    assert format_source(result.output, config).output == result.output
    assert _non_ascii(result.output) == _non_ascii(text)
