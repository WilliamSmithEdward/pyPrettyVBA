"""Configuration: presets, files, overrides, validation."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyprettyvba import PRESETS, RULES, Config, ConfigError, format_source
from pyprettyvba.config import find_config_file


def enabled(config: Config, path: str | None = None) -> set[str]:
    return set(config.settings_for(path).rules)


def test_default_preset_enables_the_default_rules() -> None:
    assert enabled(Config()) == {rule.code for rule in RULES if rule.default_enabled}


def test_every_preset_resolves() -> None:
    for preset in PRESETS:
        assert enabled(Config({"preset": preset}))


def test_vbe_preset_only_does_what_the_vbe_does() -> None:
    settings = Config({"preset": "vbe"}).settings_for(None)
    assert "blank-lines" not in settings.rules and "continuation-indent" not in settings.rules
    assert {"spacing", "keyword-case", "numeric-literals", "statement-form"} <= set(settings.rules)
    # Indentation is kept; only labels move, as the VBE moves them.
    assert settings.rules["indent"]["reindent"] is False


def test_rule_tables() -> None:
    config = Config({"rules": {"indent": False, "comment-space": True, "blank-lines": {"max-consecutive": 1}}})
    settings = config.settings_for(None)
    assert "indent" not in settings.rules
    assert "comment-space" in settings.rules
    assert settings.rules["blank-lines"]["max-consecutive"] == 1
    assert settings.rules["blank-lines"]["between-procedures"] == 1  # default kept
    assert "spacing" not in enabled(Config({"rules": {"spacing": {"enabled": False}}}))


@pytest.mark.parametrize(
    "table, message",
    [
        ({"prest": "vbe"}, "Did you mean 'preset'"),
        ({"preset": "fancy"}, "unknown preset"),
        ({"rules": {"indnet": True}}, "Did you mean 'indent'"),
        ({"rules": {"indent": {"comment": "code"}}}, "Did you mean 'comments'"),
        ({"rules": {"indent": {"comments": "sideways"}}}, "must be one of"),
        ({"rules": {"blank-lines": {"max-consecutive": -3}}}, "at least 0"),
        ({"rules": {"indent": "yes"}}, "true, false, or a table"),
        ({"indent-width": 0}, "between 1 and 16"),
        ({"indent-style": "tabs"}, "must be one of"),
        ({"hosts": ["exel"]}, "Did you mean 'excel'"),
        ({"encoding": "latin-99"}, "unknown encoding"),
        ({"overrides": [{"rules": {}}]}, "needs a files list"),
    ],
)
def test_validation_names_the_problem(table: dict[str, object], message: str) -> None:
    with pytest.raises(ConfigError, match=message):
        Config(table)


def test_overrides_apply_by_glob() -> None:
    config = Config(
        {"overrides": [{"files": ["legacy/**"], "rules": {"indent": False}, "indent-width": 2}]},
        root=Path("/project"),
    )
    assert "indent" in enabled(config, "/project/src/A.bas")
    legacy = config.settings_for("/project/legacy/old/B.bas")
    assert "indent" not in legacy.rules and legacy.indent_width == 2


def test_discovery_and_extend(tmp_path: Path) -> None:
    (tmp_path / "base.toml").write_text('indent-width = 2\n[rules]\ncomment-space = true\n', encoding="utf-8")
    (tmp_path / "project").mkdir()
    (tmp_path / "project" / "pyprettyvba.toml").write_text(
        'extend = "../base.toml"\n[rules]\nlet-keyword = true\n', encoding="utf-8"
    )
    module = tmp_path / "project" / "src" / "A.bas"
    module.parent.mkdir()
    module.write_text("x\n", encoding="utf-8")
    assert find_config_file(module) == tmp_path / "project" / "pyprettyvba.toml"
    config = Config.discover(module)
    settings = config.settings_for(module)
    assert settings.indent_width == 2
    assert {"comment-space", "let-keyword"} <= set(settings.rules)


def test_pyproject_table_is_used_only_when_present(tmp_path: Path) -> None:
    (tmp_path / "pyproject.toml").write_text("[project]\nname = 'x'\n", encoding="utf-8")
    assert find_config_file(tmp_path) is None
    (tmp_path / "pyproject.toml").write_text("[tool.pyprettyvba]\npreset = 'vbe'\n", encoding="utf-8")
    assert Config.discover(tmp_path).preset == "vbe"


def test_extend_cycle_is_an_error(tmp_path: Path) -> None:
    (tmp_path / "a.toml").write_text('extend = "b.toml"\n', encoding="utf-8")
    (tmp_path / "b.toml").write_text('extend = "a.toml"\n', encoding="utf-8")
    with pytest.raises(ConfigError, match="cycle"):
        Config.from_file(tmp_path / "a.toml")


def test_include_and_exclude(tmp_path: Path) -> None:
    config = Config({"extend-exclude": ["vendor/"]}, root=tmp_path)
    assert config.is_included(tmp_path / "A.BAS")
    assert not config.is_included(tmp_path / "notes.txt")
    assert config.is_excluded(tmp_path / "vendor" / "x.bas")
    assert config.is_excluded(tmp_path / "a" / "node_modules" / "x.bas")
    assert not config.is_excluded(tmp_path / "src" / "x.bas")


def test_command_line_changes_layer_on_top() -> None:
    config = Config({"rules": {"indent": False}}).with_changes(
        preset="strict", globals_={"indent-width": 3}, rules={"indent": True, "max-line-length": {"max": 80}}
    )
    settings = config.settings_for(None)
    assert settings.indent_width == 3
    assert settings.rules["max-line-length"]["max"] == 80
    assert "indent" in settings.rules


def test_tabs() -> None:
    out = format_source("Sub A()\nx = 1\nEnd Sub\n", Config({"indent-style": "tab"})).output
    assert out == "Sub A()\n\tx = 1\nEnd Sub\n"
