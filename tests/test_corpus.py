"""Format a corpus of real modules and check what must hold for every one.

Point PYPRETTYVBA_CORPUS at one or more directories of exported modules
(separated by the platform's path separator) and run:

    python -m pytest -m corpus tests/test_corpus.py

For every .bas, .cls, .frm and .doccls file under them, and every preset:

* formatting succeeds (output that would change the code's meaning, or rules
  that never settle, are errors);
* formatting the output again changes nothing;
* with pyVBAanalysis installed, the default preset's output gets no
  diagnostic the original did not, and loses none but the few the
  analyzer raises on valid spellings the formatter replaces (see
  MAY_DISAPPEAR).

Without PYPRETTYVBA_CORPUS the tests are skipped.
"""

from __future__ import annotations

import os
from collections import Counter
from pathlib import Path

import pytest

from pyprettyvba import PRESETS, Config, format_source
from pyprettyvba.document import split_header
from pyprettyvba.textio import decode

pytestmark = pytest.mark.corpus

EXTENSIONS = {".bas", ".cls", ".frm", ".doccls"}
SKIP_DIRS = {".git", "node_modules", ".venv", "__pycache__"}


def _corpus() -> list[Path]:
    roots = [Path(p) for p in os.environ.get("PYPRETTYVBA_CORPUS", "").split(os.pathsep) if p]
    files: list[Path] = []
    for root in roots:
        for path in sorted(root.rglob("*")):
            if path.suffix.lower() in EXTENSIONS and path.is_file() and not SKIP_DIRS & set(path.parts):
                files.append(path)
    return files


FILES = _corpus()
CONFIGS = {preset: Config({"preset": preset}) for preset in PRESETS}

if not FILES:
    pytest.skip("set PYPRETTYVBA_CORPUS to run the corpus tests", allow_module_level=True)


def _text(path: Path) -> str:
    return decode(path.read_bytes()).text


@pytest.mark.parametrize("path", FILES, ids=[str(p) for p in FILES])
def test_every_preset_formats_the_module_and_settles(path: Path) -> None:
    text = _text(path)
    for preset, config in CONFIGS.items():
        output = format_source(text, config, path=path).output
        again = format_source(output, config, path=path)
        assert again.output == output, f"{preset} is not stable"


# pyVBAanalysis reports some valid VBA that the formatter rewrites to the
# VBE's own spelling, and the report goes with it: `1.`, `&17` and `=>`
# (xlide_vscode#87), and a block If closed with a one-word `EndIf` (#88).
MAY_DISAPPEAR = {"invalid-expression-syntax", "unreachable-code"}


@pytest.mark.parametrize("path", FILES, ids=[str(p) for p in FILES])
def test_diagnostics_survive_formatting(path: Path) -> None:
    analysis = pytest.importorskip("pyvbaanalysis")
    text = _text(path)
    output = format_source(text, CONFIGS["default"], path=path).output
    if output == text:
        return

    def codes(source: str) -> Counter[str]:
        return Counter(d.code for d in analysis.analyze_module(split_header(source)[1]))

    before, after = codes(text), codes(output)
    assert not after - before, "formatting added diagnostics"
    assert set(before - after) <= MAY_DISAPPEAR, "formatting removed diagnostics"
