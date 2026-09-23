"""The Python API: format and check source text, files, and whole projects."""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from .config import Config, ConfigError, Settings, find_config_file
from .document import Document, split_header
from .engine import FormatResult, UnstableFormattingError, Violation, run_pipeline
from .rules import RULES
from .rules.base import FormatContext
from .rules.layout import detect_newline
from .safety import SafetyError
from .scopes import collect_declarations
from .textio import decode, encode

__all__ = [
    "FileResult",
    "ProjectNames",
    "check_source",
    "format_file",
    "format_paths",
    "format_source",
    "module_kind",
    "project_names",
]

KNOWN_CODES = frozenset(rule.code for rule in RULES)

_VB_NAME_RE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.IGNORECASE | re.MULTILINE)
_VB_BASE_RE = re.compile(r'^Attribute\s+VB_Base\s*=\s*"([^"]*)"', re.IGNORECASE | re.MULTILINE)


def module_kind(path: str | Path | None, source: str) -> str | None:
    """standard, class, form or document, from the file name and its header."""
    header, _body = split_header(source)
    suffix = Path(path).suffix.lower() if path is not None else ""
    if suffix == ".frm":
        return "form"
    if suffix == ".doccls":
        return "document"
    base = _VB_BASE_RE.search(header)
    if base is not None and base.group(1).startswith("0{"):
        return "document"
    if suffix == ".cls" or header.upper().startswith("VERSION 1.0 CLASS"):
        return "class"
    if suffix == ".bas":
        return "standard"
    return None


def _module_name(path: str | Path | None, source: str) -> str | None:
    match = _VB_NAME_RE.search(split_header(source)[0])
    if match:
        return match.group(1)
    return Path(path).stem if path is not None else None


@dataclass(frozen=True)
class ProjectNames:
    """What the modules of a project declare, lower-case name -> spelling.

    ``public`` holds the names a module can use from the others: the Public
    declarations of standard modules and the names of class, form and
    document modules. ``members`` holds the names reachable after a `.`:
    what classes and modules do not declare Private, and Type and Enum
    members. ``declared`` holds every declaration of every module, public or
    not, a later one replacing an earlier one: the VBE keeps one spelling per
    name for the whole project, and every declaration feeds it.
    """

    public: Mapping[str, str] = field(default_factory=dict)
    declared: Mapping[str, str] = field(default_factory=dict)
    members: Mapping[str, str] = field(default_factory=dict)


ProjectLike = ProjectNames | Mapping[str, str] | None


def _project(project: ProjectLike) -> ProjectNames:
    if project is None:
        return ProjectNames()
    if isinstance(project, ProjectNames):
        return project
    return ProjectNames(dict(project), dict(project), dict(project))


def _context(settings: Settings, source: str, path: str | Path | None, project: ProjectLike) -> FormatContext:
    names = _project(project)
    return FormatContext(
        indent_width=settings.indent_width,
        indent_style=settings.indent_style,
        tab_width=settings.tab_width,
        line_ending=settings.line_ending,
        filename=str(path) if path is not None else None,
        module_kind=module_kind(path, source),
        hosts=settings.hosts,
        project_names=dict(names.public),
        project_declared=dict(names.declared),
        project_members=dict(names.members),
        known_codes=KNOWN_CODES,
        newline=detect_newline(source, settings.line_ending),
        rule_settings=settings.rules,
    )


def format_source(
    source: str,
    config: Config | None = None,
    *,
    path: str | Path | None = None,
    project: ProjectLike = None,
    check_safety: bool = True,
) -> FormatResult:
    """Format a module's source text.

    ``path`` selects per-file overrides and tells the formatter what kind of
    module this is (by extension); it is not read. ``project`` holds what the
    other modules of the project declare (see ``project_names``), or a plain
    mapping of lower-case names to their spelling. The result carries the
    formatted text and every violation found, with 1-based positions in
    ``source``.

    Raises ``SafetyError`` if the output would mean something different from
    the input (a formatter bug; nothing is lost, the input is untouched).
    """
    config = config or Config.default()
    settings = config.settings_for(path)
    context = _context(settings, source, path, project)
    rules = settings.instantiate(context)
    header_newline = context.newline if settings.enabled("line-endings") else None
    return run_pipeline(
        source, rules, context, KNOWN_CODES, check_safety=check_safety, header_newline=header_newline
    )


def check_source(
    source: str,
    config: Config | None = None,
    *,
    path: str | Path | None = None,
    project: ProjectLike = None,
) -> list[Violation]:
    """The violations formatting ``source`` would fix or report."""
    return format_source(source, config, path=path, project=project).violations


@dataclass
class FileResult:
    """The outcome of formatting one file."""

    path: Path
    violations: list[Violation] = field(default_factory=list)
    changed: bool = False
    written: bool = False
    encoding: str | None = None
    error: str | None = None
    source: str | None = None
    output: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None


def project_names(sources: Iterable[tuple[str | Path | None, str]]) -> ProjectNames:
    """What the modules of a project declare (see ``ProjectNames``).

    ``sources`` are (path, text) pairs. When two modules declare one name,
    the later module (in the order given) gives its spelling, as a later
    declaration does inside a module.
    """
    public: dict[str, str] = {}
    declared: dict[str, str] = {}
    members: dict[str, str] = {}
    for path, source in sources:
        kind = module_kind(path, source)
        name = _module_name(path, source)
        if name and kind in ("class", "form", "document"):
            public[name.lower()] = name
            declared[name.lower()] = name
        _header, body = split_header(source)
        declarations = collect_declarations(Document(body), kind)
        public.update(declarations.public)
        declared.update(declarations.anywhere)
        members.update(declarations.member_names)
    return ProjectNames(public, declared, members)


def format_file(
    path: str | Path,
    config: Config | None = None,
    *,
    write: bool = False,
    project: ProjectLike = None,
) -> FileResult:
    """Format one file; write it back when ``write`` is true and it changed."""
    path = Path(path)
    result = FileResult(path=path)
    try:
        config = config or Config.discover(path)
        settings = config.settings_for(path)
        decoded = decode(path.read_bytes(), settings.encoding)
    except (OSError, ConfigError) as exc:
        result.error = str(exc)
        return result
    result.encoding = decoded.encoding
    result.source = decoded.text
    try:
        formatted = format_source(decoded.text, config, path=path, project=project)
    except (SafetyError, UnstableFormattingError) as exc:
        result.error = f"not formatted: {exc}"
        return result
    result.violations = formatted.violations
    result.output = formatted.output
    result.changed = formatted.changed
    if write and formatted.changed:
        try:
            path.write_bytes(encode(formatted.output, decoded.encoding, decoded.bom))
        except OSError as exc:
            result.error = f"cannot write: {exc}"
            return result
        result.written = True
    return result


def format_paths(
    paths: Sequence[str | Path],
    config: Config | None = None,
    *,
    write: bool = False,
    project_casing: bool = True,
) -> list[FileResult]:
    """Format files and directories (searched with each configuration's include/exclude).

    Files that share a configuration (or, with none, a directory) are one
    project: with ``project_casing``, names one of them declares are spelled
    the same way in the others.
    """
    files = collect_files(paths, config)
    groups: dict[str, list[Path]] = {}
    for file in files:
        key = _project_key(file, config)
        groups.setdefault(key, []).append(file)
    results: list[FileResult] = []
    for _key, members in groups.items():
        project = ProjectNames()
        if project_casing and len(members) > 1:
            sources = []
            for member in members:
                try:
                    member_config = config or Config.discover(member)
                    text = decode(member.read_bytes(), member_config.settings_for(member).encoding).text
                except (OSError, ConfigError):
                    continue
                sources.append((member, text))
            project = project_names(sources)
        for member in members:
            results.append(format_file(member, config, write=write, project=project))
    return results


def _project_key(path: Path, config: Config | None) -> str:
    if config is not None and config.root is not None:
        return str(config.root)
    found = find_config_file(path)
    return str(found.parent if found is not None else path.resolve().parent)


def collect_files(paths: Sequence[str | Path], config: Config | None = None) -> list[Path]:
    """Expand directories into the module files they hold, in a stable order."""
    found: list[Path] = []
    seen: set[Path] = set()
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            dir_config = config or Config.discover(path)
            for candidate in sorted(path.rglob("*")):
                if not candidate.is_file():
                    continue
                if dir_config.is_excluded(candidate) or not dir_config.is_included(candidate):
                    continue
                resolved = candidate.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    found.append(candidate)
        else:
            resolved = path.resolve()
            if resolved not in seen:
                seen.add(resolved)
                found.append(path)
    return found
