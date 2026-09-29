"""scripts/security/yara_update.py, the weekly move of the YARA pins.

YARA Forge's newest release is taken at once, a YARA-X release once it is a
week old, and every digest comes from GitHub's record of the asset.
"""

from __future__ import annotations

import datetime as dt
import importlib.util
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
SCRIPT = ROOT / "scripts" / "security" / "yara_update.py"
PINS = ROOT / ".github" / "security" / "yara.env"
NOW = dt.datetime(2026, 9, 29, tzinfo=dt.UTC)


def load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("yara_update", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def release(tag: str, days_old: int, digest: str = "sha256:" + "a" * 64, **flags: bool) -> dict[str, Any]:
    published = (NOW - dt.timedelta(days=days_old)).strftime("%Y-%m-%dT%H:%M:%SZ")
    asset = f"yara-x-{tag}-x86_64-unknown-linux-gnu.tar.gz"
    return {"tag_name": tag, "published_at": published, "assets": [{"name": asset, "digest": digest}], **flags}


def test_yara_x_waits_a_week_and_skips_prereleases() -> None:
    releases = [release("v1.22.0", 2), release("v1.21.1", 9, prerelease=True), release("v1.21.0", 10)]
    pin = load().yara_x_pin(releases, NOW)
    assert pin == {"YARA_X_VERSION": "v1.21.0", "YARA_X_SHA256": "a" * 64}


def test_a_release_with_no_recorded_sha256_is_refused() -> None:
    with pytest.raises(ValueError, match="no SHA-256"):
        load().yara_x_pin([release("v1.21.0", 10, digest="")], NOW)


def test_forge_takes_the_full_rule_set_digest() -> None:
    forge = {"tag_name": "20261001", "assets": [{"name": "yara-forge-rules-full.zip", "digest": "sha256:" + "b" * 64}]}
    assert load().forge_pin(forge) == {"YARA_FORGE_RELEASE": "20261001", "YARA_FORGE_SHA256": "b" * 64}


def test_the_pins_are_rewritten_in_place_and_the_moves_named() -> None:
    module = load()
    text = PINS.read_text(encoding="utf-8")
    old = module.read_pins(text)
    new = {**old, "YARA_FORGE_RELEASE": "29991231", "YARA_FORGE_SHA256": "c" * 64}
    rewritten = module.with_pins(text, new)
    assert module.read_pins(rewritten) == new
    assert [line for line in rewritten.splitlines() if line.startswith("#")] == [
        line for line in text.splitlines() if line.startswith("#")
    ]
    assert module.moves(old, new) == [f"YARA Forge {old['YARA_FORGE_RELEASE']} -> 29991231"]


def test_the_repository_s_pins_are_complete() -> None:
    pins = load().read_pins(PINS.read_text(encoding="utf-8"))
    assert set(pins) == {"YARA_X_VERSION", "YARA_X_SHA256", "YARA_FORGE_RELEASE", "YARA_FORGE_SHA256"}
    assert all(len(pins[key]) == 64 for key in ("YARA_X_SHA256", "YARA_FORGE_SHA256"))
