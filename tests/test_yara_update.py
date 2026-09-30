""".github/security/yara.json, the pins the weekly YARA update moves.

The updater, .github/security/yara_update.py, is the standard one shared by
every repository and tested where it is kept (WilliamSmithEdward/repo-standards).
Here the pin file is checked to be one it reads and writes.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).parents[1]


def test_the_pin_file_is_one_the_standard_updater_accepts() -> None:
    spec = importlib.util.spec_from_file_location("yara_update", ROOT / ".github" / "security" / "yara_update.py")
    assert spec is not None and spec.loader is not None
    updater = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(updater)
    pins = json.loads((ROOT / ".github" / "security" / "yara.json").read_text(encoding="utf-8"))
    assert set(pins) == {"yara_forge", "yara_x"}
    updater.check_move(pins, pins)
