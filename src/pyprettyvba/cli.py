"""The command line.

    pyprettyvba format [PATHS]        rewrite files in place
    pyprettyvba format --check        list files that would change (exit 1)
    pyprettyvba format --diff         print the changes as a diff (exit 1)
    pyprettyvba check [PATHS]         report violations by rule (exit 1)
    pyprettyvba check --fix           fix what can be fixed, report the rest
    pyprettyvba rules [RULE]          list the rules, or describe one
    pyprettyvba config [PATH]         show the settings that apply to a file
    pyprettyvba init                  write a starter pyprettyvba.toml

A path of `-` reads standard input and writes standard output.

Exit codes: 0 when nothing needs changing, 1 when files would change or
violations were found, 2 on an error (configuration, I/O, or output the
safety check refused).
"""

from __future__ import annotations

import argparse
import difflib
import inspect
import json
import sys
import textwrap
import tomllib
from collections import Counter
from collections.abc import Sequence
from pathlib import Path
from typing import Any, TextIO

from . import __version__
from .api import FileResult, ProjectNames, collect_files, format_file, format_source, project_names
from .config import PRESETS, Config, ConfigError
from .engine import UnstableFormattingError, Violation
from .rules import RULES, RULES_BY_CODE
from .safety import SafetyError
from .textio import decode

EXIT_CLEAN = 0
EXIT_FINDINGS = 1
EXIT_ERROR = 2


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pyprettyvba",
        description="A configurable formatter for VBA, with autofix and rule suppression.",
    )
    parser.add_argument("--version", action="version", version=f"pyprettyvba {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    def common(p: argparse.ArgumentParser) -> None:
        p.add_argument("paths", nargs="*", default=["."], help="files or directories (- for stdin)")
        p.add_argument("--config", type=Path, help="use this configuration file instead of discovering one")
        p.add_argument("--isolated", action="store_true", help="ignore configuration files")
        p.add_argument("--preset", choices=PRESETS, help="start from this preset")
        p.add_argument("--enable", action="append", default=[], metavar="RULE", help="enable a rule")
        p.add_argument("--disable", action="append", default=[], metavar="RULE", help="disable a rule")
        p.add_argument(
            "--set", action="append", default=[], metavar="RULE.OPTION=VALUE",
            help="set a rule option (the value is read as TOML)",
        )
        p.add_argument("--indent-width", type=int)
        p.add_argument("--indent-style", choices=("space", "tab"))
        p.add_argument("--line-ending", choices=("auto", "crlf", "lf"))
        p.add_argument("--stdin-filename", help="the file name to treat standard input as")
        p.add_argument(
            "--no-project-casing", action="store_true",
            help="do not spell names from other modules of the project their way",
        )
        p.add_argument("-q", "--quiet", action="store_true", help="print only errors")

    fmt = sub.add_parser("format", help="format files in place")
    common(fmt)
    fmt.add_argument("--check", action="store_true", help="do not write; exit 1 if files would change")
    fmt.add_argument("--diff", action="store_true", help="do not write; print a diff of the changes")

    chk = sub.add_parser("check", help="report violations by rule")
    common(chk)
    chk.add_argument("--fix", action="store_true", help="apply the fixes, then report what is left")
    chk.add_argument("--diff", action="store_true", help="print a diff of the fixes instead of applying them")
    chk.add_argument(
        "--output-format", choices=("text", "grouped", "json", "github"), default="text",
        help="how to print violations",
    )
    chk.add_argument("--statistics", action="store_true", help="print a count per rule instead")

    rules = sub.add_parser("rules", help="list the rules, or describe one")
    rules.add_argument("rule", nargs="?")
    rules.add_argument("--json", action="store_true")

    cfg = sub.add_parser("config", help="show the settings that apply to a file")
    cfg.add_argument("path", nargs="?", default=".")
    cfg.add_argument("--config", type=Path)
    cfg.add_argument("--json", action="store_true")

    init = sub.add_parser("init", help="write a starter pyprettyvba.toml")
    init.add_argument("--preset", choices=PRESETS, default="default")
    init.add_argument("--force", action="store_true", help="overwrite an existing file")
    init.add_argument("--path", type=Path, default=Path("pyprettyvba.toml"))
    return parser


def main(argv: Sequence[str] | None = None, *, stdout: TextIO | None = None,
         stderr: TextIO | None = None, stdin: TextIO | None = None) -> int:
    out = stdout or sys.stdout
    err = stderr or sys.stderr
    args = _parser().parse_args(argv)
    try:
        if args.command == "rules":
            return _cmd_rules(args, out, err)
        if args.command == "config":
            return _cmd_config(args, out)
        if args.command == "init":
            return _cmd_init(args, out, err)
        return _cmd_run(args, out, err, stdin or sys.stdin)
    except ConfigError as exc:
        print(f"pyprettyvba: configuration error: {exc}", file=err)
        return EXIT_ERROR


# --- configuration from flags -------------------------------------------------

def _config_for(args: argparse.Namespace, anchor: Path) -> Config:
    if args.isolated:
        base = Config()
    elif args.config is not None:
        base = Config.from_file(args.config)
    else:
        base = Config.discover(anchor)
    return _apply_flags(base, args)


def _apply_flags(base: Config, args: argparse.Namespace) -> Config:
    rules: dict[str, Any] = {}
    for code in args.enable:
        rules[code] = True
    for code in args.disable:
        rules[code] = False
    for item in args.set:
        name, sep, raw = item.partition("=")
        code, dot, option = name.strip().partition(".")
        if not sep or not dot:
            raise ConfigError(f"--set expects RULE.OPTION=VALUE, not {item!r}.")
        try:
            value = tomllib.loads(f"v = {raw}")["v"]
        except tomllib.TOMLDecodeError:
            value = raw
        table = rules.get(code)
        table = dict(table) if isinstance(table, dict) else {}
        table[option] = value
        rules[code] = table
    globals_: dict[str, Any] = {}
    if args.indent_width is not None:
        globals_["indent-width"] = args.indent_width
    if args.indent_style is not None:
        globals_["indent-style"] = args.indent_style
    if args.line_ending is not None:
        globals_["line-ending"] = args.line_ending
    return base.with_changes(preset=args.preset, globals_=globals_, rules=rules)


# --- format and check -----------------------------------------------------------

def _cmd_run(args: argparse.Namespace, out: TextIO, err: TextIO, stdin: TextIO) -> int:
    if args.paths == ["-"]:
        return _run_stdin(args, out, err, stdin)
    configs: dict[Path, Config] = {}
    files: list[Path] = []
    for raw in args.paths:
        path = Path(raw)
        if not path.exists():
            print(f"pyprettyvba: {raw}: no such file or directory", file=err)
            return EXIT_ERROR
        config = _config_for(args, path)
        for file in collect_files([path], config):
            configs.setdefault(file, config)
            files.append(file)
    if not files:
        if not args.quiet:
            print("pyprettyvba: no VBA files found", file=err)
        return EXIT_CLEAN

    write = _writes(args)
    results = _format_all(files, configs, write=write, project_casing=not args.no_project_casing)
    if args.command == "format":
        return _report_format(args, results, out, err)
    return _report_check(args, results, out, err)


def _writes(args: argparse.Namespace) -> bool:
    if args.command == "format":
        return not (args.check or args.diff)
    return bool(args.fix) and not args.diff


def _format_all(files: list[Path], configs: dict[Path, Config], *, write: bool,
                project_casing: bool) -> list[FileResult]:
    groups: dict[int, list[Path]] = {}
    for file in files:
        groups.setdefault(id(configs[file]), []).append(file)
    results: list[FileResult] = []
    for members in groups.values():
        config = configs[members[0]]
        project = ProjectNames()
        if project_casing and len(members) > 1:
            sources = []
            for member in members:
                try:
                    text = decode(member.read_bytes(), config.settings_for(member).encoding).text
                except OSError:
                    continue
                sources.append((member, text))
            project = project_names(sources)
        for member in members:
            results.append(format_file(member, config, write=write, project=project))
    return results


def _run_stdin(args: argparse.Namespace, out: TextIO, err: TextIO, stdin: TextIO) -> int:
    name = args.stdin_filename
    anchor = Path(name) if name else Path.cwd()
    config = _config_for(args, anchor)
    source = stdin.read()
    try:
        result = format_source(source, config, path=name)
    except (SafetyError, UnstableFormattingError) as exc:
        print(f"pyprettyvba: {name or '<stdin>'}: not formatted: {exc}", file=err)
        out.write(source)
        return EXIT_ERROR
    label = name or "<stdin>"
    if args.command == "check" and not args.fix and not args.diff:
        _print_violations(args, [(label, result.violations)], out)
        return EXIT_FINDINGS if result.violations else EXIT_CLEAN
    if getattr(args, "check", False):
        return EXIT_FINDINGS if result.changed else EXIT_CLEAN
    if args.diff:
        out.write(_diff(label, result.source, result.output))
        return EXIT_FINDINGS if result.changed else EXIT_CLEAN
    out.write(result.output)
    if args.command == "check":
        remaining = [v for v in result.violations if not v.fixable]
        _print_violations(args, [(label, remaining)], err)
        return EXIT_FINDINGS if remaining else EXIT_CLEAN
    return EXIT_CLEAN


def _report_format(args: argparse.Namespace, results: list[FileResult], out: TextIO, err: TextIO) -> int:
    errors = [r for r in results if not r.ok]
    changed = [r for r in results if r.ok and r.changed]
    for result in errors:
        print(f"pyprettyvba: {result.path}: {result.error}", file=err)
    if args.diff:
        for result in changed:
            assert result.source is not None and result.output is not None
            out.write(_diff(str(result.path), result.source, result.output))
    elif args.check:
        for result in changed:
            print(f"Would reformat: {result.path}", file=out)
    if not args.quiet:
        unchanged = len(results) - len(changed) - len(errors)
        verb = "reformatted" if not (args.check or args.diff) else "would be reformatted"
        print(
            f"{len(changed)} file{'s' if len(changed) != 1 else ''} {verb}, "
            f"{unchanged} file{'s' if unchanged != 1 else ''} left unchanged"
            + (f", {len(errors)} failed" if errors else ""),
            file=err,
        )
    if errors:
        return EXIT_ERROR
    if (args.check or args.diff) and changed:
        return EXIT_FINDINGS
    return EXIT_CLEAN


def _report_check(args: argparse.Namespace, results: list[FileResult], out: TextIO, err: TextIO) -> int:
    errors = [r for r in results if not r.ok]
    for result in errors:
        print(f"pyprettyvba: {result.path}: {result.error}", file=err)
    if args.diff:
        for result in results:
            if result.ok and result.changed:
                assert result.source is not None and result.output is not None
                out.write(_diff(str(result.path), result.source, result.output))
    reported: list[tuple[str, list[Violation]]] = []
    for result in results:
        if not result.ok:
            continue
        violations = result.violations
        if args.fix:
            violations = [v for v in violations if not v.fixable]
        reported.append((str(result.path), violations))
    total = sum(len(v) for _p, v in reported)
    if args.statistics:
        counts = Counter(v.rule for _p, vs in reported for v in vs)
        fixable = Counter(v.rule for _p, vs in reported for v in vs if v.fixable)
        for rule, count in counts.most_common():
            print(f"{count:6d}  {rule:<22} {'(fixable)' if fixable[rule] == count else ''}".rstrip(), file=out)
    else:
        _print_violations(args, reported, out)
    if not args.quiet and args.output_format in ("text", "grouped"):
        fixed = sum(1 for r in results if r.written)
        fixable_count = sum(1 for _p, vs in reported for v in vs if v.fixable)
        summary = f"Found {total} violation{'s' if total != 1 else ''}"
        if args.fix:
            summary += f"; fixed {fixed} file{'s' if fixed != 1 else ''}"
        elif fixable_count:
            summary += f" ({fixable_count} fixable with `pyprettyvba format` or `check --fix`)"
        print(summary + ".", file=err)
    if errors:
        return EXIT_ERROR
    return EXIT_FINDINGS if total else EXIT_CLEAN


def _print_violations(args: argparse.Namespace, reported: list[tuple[str, list[Violation]]], out: TextIO) -> None:
    fmt = getattr(args, "output_format", "text")
    if fmt == "json":
        payload = [
            {
                "path": path,
                "line": v.line,
                "column": v.column,
                "end_line": v.end_line,
                "end_column": v.end_column,
                "rule": v.rule,
                "message": v.message,
                "fixable": v.fixable,
            }
            for path, violations in reported
            for v in violations
        ]
        json.dump(payload, out, indent=2)
        out.write("\n")
        return
    for path, violations in reported:
        if fmt == "grouped" and violations:
            print(path, file=out)
        for v in violations:
            if fmt == "github":
                message = v.message.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
                print(
                    f"::warning file={path},line={v.line},col={v.column},"
                    f"endLine={v.end_line},endColumn={v.end_column},title=pyprettyvba ({v.rule})::{message}",
                    file=out,
                )
            elif fmt == "grouped":
                mark = " [*]" if v.fixable else ""
                print(f"  {v.line}:{v.column} {v.rule} {v.message}{mark}", file=out)
            else:
                mark = " [*]" if v.fixable else ""
                print(f"{path}:{v.line}:{v.column}: {v.rule} {v.message}{mark}", file=out)


def _diff(label: str, before: str, after: str) -> str:
    lines = difflib.unified_diff(
        before.splitlines(keepends=True),
        after.splitlines(keepends=True),
        fromfile=f"a/{label}",
        tofile=f"b/{label}",
    )
    text = "".join(line if line.endswith(("\n", "\r")) else line + "\n" for line in lines)
    return text.replace("\r\n", "\n").replace("\r", "\n")


# --- rules, config, init -------------------------------------------------------

def _cmd_rules(args: argparse.Namespace, out: TextIO, err: TextIO) -> int:
    defaults = Config().settings_for(None)
    if args.rule:
        rule = RULES_BY_CODE.get(args.rule)
        if rule is None:
            import difflib as _difflib

            close = _difflib.get_close_matches(args.rule, list(RULES_BY_CODE), n=1)
            hint = f"; did you mean {close[0]!r}?" if close else ""
            print(f"pyprettyvba: unknown rule {args.rule!r}{hint}", file=err)
            return EXIT_ERROR
        if args.json:
            json.dump(_rule_json(rule, defaults.enabled(rule.code)), out, indent=2)
            out.write("\n")
            return EXIT_CLEAN
        print(f"{rule.code}: {rule.summary}", file=out)
        status = "enabled" if defaults.enabled(rule.code) else "disabled"
        traits = [status + " by default", "fixable" if rule.fixable else "report only"]
        if rule.vbe_canonical:
            traits.append("does what the VBE does")
        print(f"({', '.join(traits)}; category: {rule.category})", file=out)
        print(file=out)
        # cleandoc, not dedent: before 3.13 a docstring keeps the indentation
        # of every line but its first, which dedent then leaves in place.
        print(inspect.cleandoc(rule.__doc__ or ""), file=out)
        if rule.options:
            print("\nOptions:", file=out)
            for option in rule.options:
                choices = f" ({' | '.join(map(str, option.choices))})" if option.choices else ""
                print(f"  {option.name} = {json.dumps(option.default)}{choices}", file=out)
                for line in textwrap.wrap(option.doc, 72):
                    print(f"      {line}", file=out)
        return EXIT_CLEAN
    if args.json:
        json.dump([_rule_json(rule, defaults.enabled(rule.code)) for rule in RULES], out, indent=2)
        out.write("\n")
        return EXIT_CLEAN
    width = max(len(rule.code) for rule in RULES)
    for rule in RULES:
        flag = "on " if defaults.enabled(rule.code) else "off"
        fix = "fix" if rule.fixable else "   "
        print(f"{rule.code:<{width}}  {flag}  {fix}  {rule.summary}", file=out)
    return EXIT_CLEAN


def _rule_json(rule: Any, enabled: bool) -> dict[str, Any]:
    return {
        "code": rule.code,
        "summary": rule.summary,
        "category": rule.category,
        "enabled_by_default": enabled,
        "fixable": rule.fixable,
        "vbe_canonical": rule.vbe_canonical,
        "description": inspect.cleandoc(rule.__doc__ or ""),
        "options": [
            {"name": o.name, "default": o.default, "choices": list(o.choices) if o.choices else None, "doc": o.doc}
            for o in rule.options
        ],
    }


def _cmd_config(args: argparse.Namespace, out: TextIO) -> int:
    path = Path(args.path)
    config = Config.from_file(args.config) if args.config else Config.discover(path)
    settings = config.settings_for(path if path.is_file() else None)
    data = {
        "config-file": str(config.source) if config.source else None,
        "preset": config.preset,
        "indent-width": settings.indent_width,
        "indent-style": settings.indent_style,
        "tab-width": settings.tab_width,
        "line-ending": settings.line_ending,
        "encoding": settings.encoding,
        "hosts": list(settings.hosts),
        "include": config.include,
        "exclude": config.exclude,
        "rules": {code: dict(options) for code, options in settings.rules.items()},
        "disabled": [rule.code for rule in RULES if rule.code not in settings.rules],
    }
    if args.json:
        json.dump(data, out, indent=2)
        out.write("\n")
        return EXIT_CLEAN
    print(f"configuration: {data['config-file'] or '(none found; built-in defaults)'}", file=out)
    print(f"preset: {data['preset']}", file=out)
    for key in ("indent-width", "indent-style", "tab-width", "line-ending", "encoding", "hosts"):
        print(f"{key}: {data[key]}", file=out)
    print("enabled rules:", file=out)
    for code, options in data["rules"].items():  # type: ignore[union-attr]
        shown = ", ".join(f"{k}={json.dumps(v)}" for k, v in options.items())
        print(f"  {code}" + (f" ({shown})" if shown else ""), file=out)
    print("disabled rules: " + (", ".join(data["disabled"]) or "none"), file=out)  # type: ignore[arg-type]
    return EXIT_CLEAN


_INIT_TEMPLATE = """\
# pyPrettyVBA configuration. Every key is optional; see
# `pyprettyvba rules` for the rules and `pyprettyvba rules NAME` for options.

preset = "{preset}"   # default | vbe | xlide | strict | minimal
indent-width = 4
indent-style = "space"
line-ending = "auto"  # crlf is the safe choice for code the VBE imports

[rules]
# blank-lines = {{ max-consecutive = 1 }}
# comment-space = true
# align-declarations = true

# [[overrides]]
# files = ["legacy/**"]
# rules = {{ indent = false }}
"""


def _cmd_init(args: argparse.Namespace, out: TextIO, err: TextIO) -> int:
    target: Path = args.path
    if target.exists() and not args.force:
        print(f"pyprettyvba: {target} exists; use --force to overwrite", file=err)
        return EXIT_ERROR
    target.write_text(_INIT_TEMPLATE.format(preset=args.preset), encoding="utf-8")
    print(f"wrote {target}", file=out)
    return EXIT_CLEAN


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
