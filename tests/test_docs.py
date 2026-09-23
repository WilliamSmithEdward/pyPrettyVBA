"""The generated rule reference matches the rules."""

from __future__ import annotations

import importlib.util
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def test_docs_rules_md_is_current() -> None:
    spec = importlib.util.spec_from_file_location("build_docs", REPO / "tools" / "build_docs.py")
    assert spec is not None and spec.loader is not None
    build_docs = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(build_docs)
    written = (REPO / "docs" / "rules.md").read_text(encoding="utf-8")
    assert written == build_docs.render(), "docs/rules.md is stale: run python tools/build_docs.py"
