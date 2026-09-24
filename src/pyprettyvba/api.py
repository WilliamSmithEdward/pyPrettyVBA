"""The Python API: format and check source text, files, and whole projects."""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from . import office
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
    "format_office_file",
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


def _context(settings: Settings, source: str, path: str | Path | None, project: ProjectLike,
             kind: str | None = None) -> FormatContext:
    names = _project(project)
    return FormatContext(
        indent_width=settings.indent_width,
        indent_style=settings.indent_style,
        tab_width=settings.tab_width,
        line_ending=settings.line_ending,
        filename=str(path) if path is not None else None,
        module_kind=kind or module_kind(path, source),
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
    kind: str | None = None,
    check_safety: bool = True,
) -> FormatResult:
    """Format a module's source text.

    ``path`` selects per-file overrides and tells the formatter what kind of
    module this is (by extension); it is not read. ``kind`` (standard,
    class, form or document) says so directly, for text that comes from
    somewhere other than a file. ``project`` holds what the other modules of
    the project declare (see ``project_names``), or a plain mapping of
    lower-case names to their spelling. The result carries the formatted
    text and every violation found, with 1-based positions in ``source``.

    Raises ``SafetyError`` if the output would mean something different from
    the input (a formatter bug; nothing is lost, the input is untouched).
    """
    config = config or Config.default()
    settings = config.settings_for(path)
    context = _context(settings, source, path, project, kind)
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
    """The outcome of formatting one file, or one module of an Office file."""

    path: Path
    violations: list[Violation] = field(default_factory=list)
    changed: bool = False
    written: bool = False
    encoding: str | None = None
    error: str | None = None
    source: str | None = None
    output: str | None = None
    # The module's name, for a module of an Office file.
    module: str | None = None
    # Whether the Office file was written without the digital signature it had.
    signature_removed: bool = False

    @property
    def ok(self) -> bool:
        return self.error is None

    @property
    def label(self) -> str:
        """How reports name it: the path, and for an Office file `path:Module`."""
        return f"{self.path}:{self.module}" if self.module is not None else str(self.path)


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
    """Format one module file; write it back when ``write`` is true and it changed.

    An Office file holds several modules: see ``format_office_file``.
    """
    path = Path(path)
    result = FileResult(path=path)
    if office.is_office_file(path):
        result.error = "an Office file holds several modules: format it with format_office_file or format_paths"
        return result
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


def format_office_file(
    path: str | Path,
    config: Config | None = None,
    *,
    write: bool = False,
    project_casing: bool = True,
    remove_signatures: bool = False,
) -> list[FileResult]:
    """Format the modules of the VBA project inside an Office file.

    Excel (.xlsm .xlsb .xlam .xls), Word (.docm .dotm .doc), PowerPoint
    (.pptm .potm .ppt) and Access (.accdb .mdb) files are read and written
    through pyOpenVBA. There is one result per module, named by
    ``FileResult.module``, or a single result carrying the error when the
    file cannot be read.

    The modules are one project, as they are in the VBE: with
    ``project_casing``, a name one of them declares is spelled its way in
    the others. Their line endings are CRLF whatever the configuration
    says, since that is how a VBA project stores its code. With ``write``,
    the changed modules go back into the file: it is saved beside the
    original and then moved over it, so an interrupted save changes
    nothing. A project with a digital signature is not written, since
    editing it would invalidate the signature, unless ``remove_signatures``:
    then it is written without one, to be signed again, and its results
    say ``signature_removed``. The signature is found in the zip-based
    formats, not in .xls, .doc, .ppt or Access files.
    """
    path = Path(path)
    try:
        config = (config or Config.discover(path)).with_changes(globals_={"line-ending": "crlf"})
    except ConfigError as exc:
        return [FileResult(path=path, error=str(exc))]
    results: list[FileResult] = []
    saved: Path | None = None
    signature_removed = False
    try:
        with office.open_host(path) as host:
            modules = office.modules_of(host)
            project = ProjectNames()
            if project_casing:
                project = project_names((_office_module_path(m), m.source) for m in modules)
            changes: dict[str, str] = {}
            for module in modules:
                result = FileResult(path=path, module=module.name, source=module.source)
                results.append(result)
                try:
                    formatted = format_source(module.source, config, path=path, project=project, kind=module.kind)
                except (SafetyError, UnstableFormattingError) as exc:
                    result.error = f"not formatted: {exc}"
                    continue
                result.violations = formatted.violations
                result.output = formatted.output
                result.changed = formatted.changed
                if formatted.changed:
                    changes[module.name] = formatted.output
            if write and changes:
                for name, text in changes.items():
                    host.set_module(name, text)
                saved, signature_removed = office.save_beside(host, path, remove_signatures=remove_signatures)
    except office.SignedProjectError as exc:
        return _not_written(results, f"not written: {exc}")
    except office.FILE_ERRORS as exc:
        if not results:
            return [FileResult(path=path, error=f"cannot read the VBA project: {exc}")]
        return _not_written(results, f"cannot write: {exc}")
    if saved is not None:
        try:
            office.replace(saved, path)
        except OSError as exc:
            return _not_written(results, f"cannot write: {exc}")
        for result in results:
            result.written = result.changed
            result.signature_removed = signature_removed
    return results


def _office_module_path(module: office.OfficeModule) -> Path:
    """A file name that tells ``project_names`` the module's kind."""
    return Path(module.name + (".bas" if module.kind == "standard" else ".cls"))


def _not_written(results: list[FileResult], error: str) -> list[FileResult]:
    for result in results:
        if result.ok and result.changed:
            result.error = error
    return results


def format_paths(
    paths: Sequence[str | Path],
    config: Config | None = None,
    *,
    write: bool = False,
    project_casing: bool = True,
    remove_signatures: bool = False,
) -> list[FileResult]:
    """Format files and directories (searched with each configuration's include/exclude).

    Module files that share a configuration (or, with none, a directory) are
    one project: with ``project_casing``, names one of them declares are
    spelled the same way in the others. An Office file is a project of its
    own (see ``format_office_file``, and its ``remove_signatures``).
    """
    files = collect_files(paths, config)
    results: list[FileResult] = []
    for file in files:
        if office.is_office_file(file):
            results.extend(
                format_office_file(
                    file, config, write=write, project_casing=project_casing, remove_signatures=remove_signatures
                )
            )
    files = [file for file in files if not office.is_office_file(file)]
    groups: dict[str, list[Path]] = {}
    for file in files:
        key = _project_key(file, config)
        groups.setdefault(key, []).append(file)
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
