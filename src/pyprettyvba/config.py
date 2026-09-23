"""Configuration: presets, files, per-file overrides, and validation.

A configuration is a TOML table, read from the first of these found walking
up from the file being formatted:

    pyprettyvba.toml          (the table is the whole file)
    .pyprettyvba.toml         (the same, as a hidden file)
    pyproject.toml            (the [tool.pyprettyvba] table, if present)

Keys (all optional):

    preset = "default"            # default | vbe | xlide | strict | minimal | none
    extend = "../base.toml"       # start from another configuration file
    indent-width = 4
    indent-style = "space"        # space | tab
    tab-width = 4                 # columns a tab advances to
    line-ending = "auto"          # auto | crlf | lf
    encoding = "auto"             # auto, or a codec name such as cp1252
    hosts = ["excel"]             # libraries identifier-case knows names from
    include = ["*.bas", "*.cls", "*.frm", "*.doccls"]
    exclude = [".git", "node_modules"]
    extend-exclude = ["vendor/"]

    [rules]
    blank-lines = { max-consecutive = 1 }
    comment-space = true
    numeric-literals = false

    [[overrides]]
    files = ["legacy/**"]
    rules = { indent = false }

A rule set to `true` is enabled with its preset options; a table enables it
(unless it says `enabled = false`) and sets options; `false` disables it.
Unknown keys, rules and options are errors, reported with the closest name
that exists.
"""

from __future__ import annotations

import copy
import difflib
import re
import tomllib
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any

from .names import HOSTS
from .rules import RULES, RULES_BY_CODE, Rule

__all__ = [
    "CONFIG_FILES",
    "Config",
    "ConfigError",
    "DEFAULT_EXCLUDE",
    "DEFAULT_INCLUDE",
    "PRESETS",
    "Settings",
    "find_config_file",
]

CONFIG_FILES = ("pyprettyvba.toml", ".pyprettyvba.toml", "pyproject.toml")
DEFAULT_INCLUDE = ("*.bas", "*.cls", "*.frm", "*.doccls")
DEFAULT_EXCLUDE = (
    ".git", ".hg", ".svn", ".venv", "venv", "node_modules", "__pycache__", "build", "dist",
)

_GLOBAL_KEYS: dict[str, tuple[Any, ...]] = {
    "indent-width": (int, 1, 16),
    "indent-style": (("space", "tab"),),
    "tab-width": (int, 1, 16),
    "line-ending": (("auto", "crlf", "lf"),),
    "encoding": (str,),
    "hosts": (list,),
}
_FILE_KEYS = {"preset", "extend", "include", "exclude", "extend-exclude", "rules", "overrides"}
_OVERRIDE_KEYS = {"files", "exclude-files", "rules", *_GLOBAL_KEYS}


class ConfigError(Exception):
    """A configuration could not be read or is invalid."""


def _preset_rules(name: str) -> dict[str, dict[str, Any] | None]:
    """Rule code -> options (None when disabled) for a preset."""
    rules: dict[str, dict[str, Any] | None] = {}
    for rule in RULES:
        rules[rule.code] = {} if rule.default_enabled else None
    if name == "default":
        return rules
    if name == "vbe":
        # Only what the VBE itself does to code it reads in.
        for rule in RULES:
            rules[rule.code] = {} if rule.vbe_canonical else None
        rules["indent"] = {"reindent": False}
        rules["identifier-case"] = {"scope": "module", "undeclared": "last"}
        rules["trailing-whitespace"] = {"blank-lines": "keep"}
        rules["end-of-file"] = {"trailing-blank-lines": "keep"}
        rules["line-endings"] = {}
        rules["suppression-directive"] = {}
        return rules
    if name == "xlide":
        # XLIDE's Format Document: indentation, casing, inserted spaces,
        # trailing whitespace; never removes a space or a line.
        enabled: dict[str, dict[str, Any]] = {
            "indent": {},
            "continuation-indent": {"style": "relative", "closing-paren": "indent"},
            "keyword-case": {"expand-endif": False},
            "identifier-case": {"library": "globals", "members": "keep"},
            "spacing": {"collapse": False},
            "trailing-whitespace": {"blank-lines": "indent"},
            "suppression-directive": {},
        }
        return {code: enabled.get(code) for code in rules}
    if name == "strict":
        for code in ("comment-space", "rem-comments", "let-keyword", "split-statements", "align-declarations"):
            rules[code] = {}
        rules["max-line-length"] = {"max": 120}
        rules["blank-lines"] = {"max-consecutive": 1}
        return rules
    if name == "minimal":
        keep = {"trailing-whitespace", "end-of-file", "line-endings", "suppression-directive"}
        return {code: ({} if code in keep else None) for code in rules}
    if name == "none":
        # Nothing but directive checking: enable rules one by one.
        return {code: ({} if code == "suppression-directive" else None) for code in rules}
    raise ConfigError(f"Unknown preset {name!r}.{_suggest(name, PRESETS)}")


PRESETS = ("default", "vbe", "xlide", "strict", "minimal", "none")


def _suggest(name: str, choices: Sequence[str] | Mapping[str, Any]) -> str:
    match = difflib.get_close_matches(name, list(choices), n=1)
    return f" Did you mean {match[0]!r}?" if match else ""


@dataclass(frozen=True)
class Settings:
    """Everything that decides how one file is formatted."""

    indent_width: int = 4
    indent_style: str = "space"
    tab_width: int = 4
    line_ending: str = "auto"
    encoding: str = "auto"
    hosts: tuple[str, ...] = ("excel",)
    # Rule code -> resolved options, for every enabled rule.
    rules: Mapping[str, Mapping[str, Any]] = field(default_factory=dict)

    def enabled(self, code: str) -> bool:
        return code in self.rules

    def instantiate(self, context: Any) -> list[Rule]:
        """Rule instances for this file, in pipeline order."""
        return [rule(self.rules[rule.code], context) for rule in RULES if rule.code in self.rules]


@dataclass
class _Override:
    files: list[str]
    exclude_files: list[str]
    table: dict[str, Any]


class Config:
    """A validated configuration: presets, global settings, rules, overrides."""

    def __init__(self, table: Mapping[str, Any] | None = None, *, root: Path | None = None,
                 source: Path | None = None) -> None:
        table = dict(table or {})
        self.source = source
        self.root = root.resolve() if root is not None else None
        where = str(source) if source else "configuration"
        _check_keys(table, _FILE_KEYS | set(_GLOBAL_KEYS), where)
        self.preset: str = table.get("preset", "default")
        if not isinstance(self.preset, str) or self.preset not in PRESETS:
            raise ConfigError(f"{where}: unknown preset {self.preset!r}.{_suggest(str(self.preset), PRESETS)}")
        self.globals: dict[str, Any] = {}
        for key in _GLOBAL_KEYS:
            if key in table:
                self.globals[key] = _check_global(key, table[key], where)
        self.include: list[str] = _string_list(table.get("include", list(DEFAULT_INCLUDE)), "include", where)
        exclude = _string_list(table.get("exclude", list(DEFAULT_EXCLUDE)), "exclude", where)
        exclude += _string_list(table.get("extend-exclude", []), "extend-exclude", where)
        self.exclude: list[str] = exclude
        self.rule_table: dict[str, Any] = _check_rules(table.get("rules", {}), where)
        self.overrides: list[_Override] = []
        raw_overrides = table.get("overrides", [])
        if not isinstance(raw_overrides, list):
            raise ConfigError(f"{where}: overrides must be an array of tables.")
        for number, item in enumerate(raw_overrides, start=1):
            label = f"{where}: overrides[{number}]"
            if not isinstance(item, dict):
                raise ConfigError(f"{label} must be a table.")
            _check_keys(item, _OVERRIDE_KEYS, label)
            if "files" not in item:
                raise ConfigError(f"{label} needs a files list.")
            override_table = {k: v for k, v in item.items() if k not in ("files", "exclude-files")}
            for key in _GLOBAL_KEYS:
                if key in override_table:
                    override_table[key] = _check_global(key, override_table[key], label)
            if "rules" in override_table:
                override_table["rules"] = _check_rules(override_table["rules"], label)
            self.overrides.append(
                _Override(
                    files=_string_list(item["files"], "files", label),
                    exclude_files=_string_list(item.get("exclude-files", []), "exclude-files", label),
                    table=override_table,
                )
            )
        # The resolved settings for a file with no matching override.
        self._base = self._resolve(self.globals, self.rule_table)

    # -- construction --------------------------------------------------------

    @classmethod
    def default(cls) -> Config:
        return cls()

    @classmethod
    def from_file(cls, path: str | Path) -> Config:
        """Read a configuration file (following `extend`)."""
        path = Path(path).resolve()
        table = _load_table(path)
        return cls(table, root=path.parent, source=path)

    @classmethod
    def discover(cls, start: str | Path) -> Config:
        """The configuration governing ``start`` (a file or directory)."""
        found = find_config_file(Path(start))
        return cls.from_file(found) if found is not None else cls(root=None)

    # -- queries ---------------------------------------------------------------

    def settings_for(self, path: str | Path | None = None) -> Settings:
        """The settings for one file, with matching overrides applied."""
        matching = [o for o in self.overrides if path is not None and self._matches(o, Path(path))]
        if not matching:
            return self._base
        globals_ = dict(self.globals)
        rules = copy.deepcopy(self.rule_table)
        for override in matching:
            for key, value in override.table.items():
                if key == "rules":
                    for code, setting in value.items():
                        rules[code] = _combine_raw(rules.get(code), setting)
                else:
                    globals_[key] = value
        return self._resolve(globals_, rules)

    def with_changes(
        self,
        *,
        preset: str | None = None,
        globals_: Mapping[str, Any] | None = None,
        rules: Mapping[str, Any] | None = None,
    ) -> Config:
        """A copy with command-line changes applied on top."""
        clone = copy.copy(self)
        if preset is not None:
            if preset not in PRESETS:
                raise ConfigError(f"unknown preset {preset!r}.{_suggest(preset, PRESETS)}")
            clone.preset = preset
        clone.globals = dict(self.globals)
        for key, value in (globals_ or {}).items():
            clone.globals[key] = _check_global(key, value, "command line")
        clone.rule_table = copy.deepcopy(self.rule_table)
        for code, setting in _check_rules(dict(rules or {}), "command line").items():
            clone.rule_table[code] = _combine_raw(clone.rule_table.get(code), setting)
        clone.overrides = list(self.overrides)
        clone._base = clone._resolve(clone.globals, clone.rule_table)
        return clone

    def is_excluded(self, path: Path) -> bool:
        rel = self._relative(path)
        return any(_glob_match(pattern, rel) for pattern in self.exclude)

    def is_included(self, path: Path) -> bool:
        rel = self._relative(path)
        return any(_glob_match(pattern, rel) for pattern in self.include)

    # -- internals -----------------------------------------------------------

    def _relative(self, path: Path) -> str:
        """The path globs are matched against: relative to the configuration's folder."""
        if self.root is None:
            # A configuration with no file: a relative path is matched as given.
            return path.as_posix() if not path.is_absolute() else path.resolve().as_posix()
        resolved = path.resolve()
        try:
            return resolved.relative_to(self.root).as_posix()
        except ValueError:
            return resolved.as_posix()

    def _matches(self, override: _Override, path: Path) -> bool:
        rel = self._relative(path)
        if not any(_glob_match(p, rel) for p in override.files):
            return False
        return not any(_glob_match(p, rel) for p in override.exclude_files)

    def _resolve(self, globals_: Mapping[str, Any], rule_table: Mapping[str, Any]) -> Settings:
        rules = _preset_rules(self.preset)
        for code, setting in rule_table.items():
            rules[code] = _merge_rule(rules.get(code), setting)
        resolved: dict[str, dict[str, Any]] = {}
        for rule in RULES:
            options = rules.get(rule.code)
            if options is None:
                continue
            values = {option.name: option.default for option in rule.options}
            values.update(options)
            resolved[rule.code] = values
        return Settings(
            indent_width=int(globals_.get("indent-width", 4)),
            indent_style=str(globals_.get("indent-style", "space")),
            tab_width=int(globals_.get("tab-width", 4)),
            line_ending=str(globals_.get("line-ending", "auto")),
            encoding=str(globals_.get("encoding", "auto")),
            hosts=tuple(globals_.get("hosts", ("excel",))),
            rules=resolved,
        )


def _combine_raw(base: Any, override: Any) -> Any:
    """Layer one raw rule setting (true, false or a table) over another."""
    if isinstance(override, bool):
        return override
    if isinstance(base, dict):
        merged = dict(base)
        merged.update(override)
        return merged
    return dict(override)


def _merge_rule(current: dict[str, Any] | None, setting: Any) -> dict[str, Any] | None:
    """Apply one rule setting (bool or table) over the current options."""
    if setting is False:
        return None
    if setting is True:
        return dict(current or {})
    table = dict(setting)
    enabled = table.pop("enabled", True)
    if not enabled:
        return None
    merged = dict(current or {})
    merged.update(table)
    return merged


def _check_keys(table: Mapping[str, Any], allowed: set[str], where: str) -> None:
    for key in table:
        if key not in allowed:
            raise ConfigError(f"{where}: unknown setting {key!r}.{_suggest(key, sorted(allowed))}")


def _check_global(key: str, value: Any, where: str) -> Any:
    spec = _GLOBAL_KEYS[key]
    if isinstance(spec[0], tuple):
        if value not in spec[0]:
            allowed = ", ".join(repr(v) for v in spec[0])
            raise ConfigError(f"{where}: {key} must be one of {allowed}, not {value!r}.")
        return value
    kind = spec[0]
    if kind is int:
        if isinstance(value, bool) or not isinstance(value, int):
            raise ConfigError(f"{where}: {key} must be an integer, not {value!r}.")
        low, high = spec[1], spec[2]
        if not low <= value <= high:
            raise ConfigError(f"{where}: {key} must be between {low} and {high}, not {value}.")
        return value
    if key == "hosts":
        hosts = _string_list(value, key, where)
        for host in hosts:
            if host not in HOSTS:
                raise ConfigError(f"{where}: unknown host {host!r}; hosts are {', '.join(HOSTS)}.{_suggest(host, HOSTS)}")
        return hosts
    if key == "encoding":
        if not isinstance(value, str):
            raise ConfigError(f"{where}: encoding must be a string.")
        if value != "auto":
            import codecs

            try:
                codecs.lookup(value)
            except LookupError:
                raise ConfigError(f"{where}: unknown encoding {value!r}.") from None
        return value
    return value


def _check_rules(table: Any, where: str) -> dict[str, Any]:
    if not isinstance(table, dict):
        raise ConfigError(f"{where}: rules must be a table.")
    checked: dict[str, Any] = {}
    for code, setting in table.items():
        rule = RULES_BY_CODE.get(code)
        if rule is None:
            raise ConfigError(f"{where}: unknown rule {code!r}.{_suggest(code, RULES_BY_CODE)}")
        if isinstance(setting, bool):
            checked[code] = setting
            continue
        if not isinstance(setting, dict):
            raise ConfigError(f"{where}: rules.{code} must be true, false, or a table of options.")
        names = {option.name for option in rule.options} | {"enabled"}
        for name, value in setting.items():
            if name not in names:
                raise ConfigError(
                    f"{where}: rule {code} has no option {name!r}.{_suggest(name, sorted(names))}"
                )
            if name == "enabled":
                if not isinstance(value, bool):
                    raise ConfigError(f"{where}: rules.{code}.enabled must be true or false.")
                continue
            problem = rule.option(name).validate(value)
            if problem is not None:
                raise ConfigError(f"{where}: rules.{code}.{name} {problem}.")
        checked[code] = dict(setting)
    return checked


def _string_list(value: Any, key: str, where: str) -> list[str]:
    if not isinstance(value, list) or not all(isinstance(v, str) for v in value):
        raise ConfigError(f"{where}: {key} must be a list of strings.")
    return list(value)


def _load_table(path: Path, seen: tuple[Path, ...] = ()) -> dict[str, Any]:
    if path in seen:
        raise ConfigError(f"{path}: extend forms a cycle.")
    try:
        data: dict[str, Any] = tomllib.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise ConfigError(f"{path}: cannot read ({exc.strerror or exc}).") from None
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{path}: invalid TOML ({exc}).") from None
    if path.name == "pyproject.toml":
        data = data.get("tool", {}).get("pyprettyvba", {})
    base_name = data.pop("extend", None)
    if base_name is None:
        return data
    if not isinstance(base_name, str):
        raise ConfigError(f"{path}: extend must be a path.")
    base_path = (path.parent / base_name).resolve()
    base = _load_table(base_path, seen + (path,))
    merged = dict(base)
    for key, value in data.items():
        if key == "rules" and isinstance(value, dict) and isinstance(base.get("rules"), dict):
            rules = dict(base["rules"])
            rules.update(value)
            merged["rules"] = rules
        else:
            merged[key] = value
    return merged


def find_config_file(start: Path) -> Path | None:
    """The nearest configuration file at or above ``start``."""
    start = start.resolve()
    directory = start if start.is_dir() else start.parent
    for folder in (directory, *directory.parents):
        for name in CONFIG_FILES:
            candidate = folder / name
            if not candidate.is_file():
                continue
            if name == "pyproject.toml" and not _has_tool_table(candidate):
                continue
            return candidate
    return None


def _has_tool_table(path: Path) -> bool:
    try:
        data: dict[str, Any] = tomllib.loads(path.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError):
        return False
    return isinstance(data.get("tool", {}).get("pyprettyvba"), dict)


# --- globs ------------------------------------------------------------------

_GLOB_CACHE: dict[str, re.Pattern[str]] = {}


def _glob_regex(pattern: str) -> re.Pattern[str]:
    cached = _GLOB_CACHE.get(pattern)
    if cached is not None:
        return cached
    anchored = "/" in pattern.rstrip("/")
    body = pattern.strip("/")
    out = []
    i = 0
    while i < len(body):
        c = body[i]
        if body.startswith("**/", i):
            out.append("(?:.*/)?")
            i += 3
            continue
        if body.startswith("**", i):
            out.append(".*")
            i += 2
            continue
        if c == "*":
            out.append("[^/]*")
        elif c == "?":
            out.append("[^/]")
        elif c == "[":
            close = body.find("]", i + 1)
            if close == -1:
                out.append(re.escape(c))
            else:
                out.append("[" + body[i + 1 : close].replace("!", "^", 1) + "]")
                i = close
        else:
            out.append(re.escape(c))
        i += 1
    prefix = "" if anchored else "(?:.*/)?"
    compiled = re.compile("^" + prefix + "".join(out) + "(?:/.*)?$", re.IGNORECASE)
    _GLOB_CACHE[pattern] = compiled
    return compiled


def _glob_match(pattern: str, relative: str) -> bool:
    """Match a gitignore-style glob against a posix relative path.

    A pattern without a slash matches a file or directory name at any depth;
    one with a slash is anchored to the configuration's directory. A match on
    a directory covers everything under it.
    """
    return _glob_regex(pattern).match(PurePosixPath(relative).as_posix()) is not None
