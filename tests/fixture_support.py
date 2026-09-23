"""The fixture suite: cases that show each surface of pyPrettyVBA, as data.

A case is a directory under tests/fixtures holding:

    case.toml        title, description, the file name to format the input as,
                     optional [config] (a pyprettyvba.toml table), and for a
                     command-line case a [cli] table (args, exit, stdin)
    input.*          the module exactly as written (bytes)
    output.*         the module after formatting (absent when nothing changes)
    check.txt        what `check` reports: `line:col rule message [*]`
    stdout.txt       for a command-line case, what the command prints

A case with an input/ directory instead of input.* formats every file in it
as one project, and expects output/ to mirror it.

The unit suite (tests/test_fixtures.py) runs every case; tools/fixtures.py
bless writes the expected files from current behavior, for review; and
tools/build_gallery.py renders the cases as a page.
"""

from __future__ import annotations

import io
import os
import shutil
import tempfile
import tomllib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pyprettyvba import Config, Violation, format_source
from pyprettyvba.api import format_paths
from pyprettyvba.cli import main as cli_main
from pyprettyvba.textio import decode, encode

FIXTURES = Path(__file__).resolve().parent / "fixtures"


@dataclass
class Case:
    path: Path
    title: str
    description: str
    file: str
    config: dict[str, Any]
    cli: dict[str, Any] | None
    tags: list[str] = field(default_factory=list)

    @property
    def id(self) -> str:
        return self.path.relative_to(FIXTURES).as_posix()

    @property
    def surface(self) -> str:
        return self.path.relative_to(FIXTURES).parts[0]

    @property
    def project(self) -> bool:
        return (self.path / "input").is_dir()

    def input_path(self) -> Path:
        return self.path / ("input" + Path(self.file).suffix)

    def output_path(self) -> Path:
        return self.path / ("output" + Path(self.file).suffix)


def load_cases(root: Path = FIXTURES) -> list[Case]:
    cases = []
    for toml_path in sorted(root.rglob("case.toml")):
        data = tomllib.loads(toml_path.read_text(encoding="utf-8"))
        cases.append(
            Case(
                path=toml_path.parent,
                title=data["title"],
                description=data.get("description", "").strip(),
                file=data.get("file", "Module1.bas"),
                config=data.get("config", {}),
                cli=data.get("cli"),
                tags=list(data.get("tags", [])),
            )
        )
    return cases


def format_violations(violations: list[Violation]) -> str:
    lines = [
        f"{v.line}:{v.column} {v.rule} {v.message}{' [*]' if v.fixable else ''}" for v in violations
    ]
    return "".join(line + "\n" for line in lines)


@dataclass
class Outcome:
    """What a case produces now."""

    output: bytes | None = None  # formatted bytes (single-file cases)
    check: str = ""
    outputs: dict[str, bytes] = field(default_factory=dict)  # project cases
    stdout: str | None = None
    exit_code: int | None = None


def run_case(case: Case) -> Outcome:
    if case.cli is not None:
        return _run_cli(case)
    if case.project:
        return _run_project(case)
    config = Config(case.config)
    decoded = decode(case.input_path().read_bytes(), config.settings_for(case.file).encoding)
    result = format_source(decoded.text, config, path=case.file)
    return Outcome(
        output=encode(result.output, decoded.encoding, decoded.bom),
        check=format_violations(result.violations),
    )


def _run_project(case: Case) -> Outcome:
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp) / "project"
        shutil.copytree(case.path / "input", work)
        config = Config(case.config, root=work)
        results = format_paths([work], config, write=True)
        outcome = Outcome()
        lines = []
        for result in sorted(results, key=lambda r: r.path.name):
            name = result.path.relative_to(work).as_posix()
            outcome.outputs[name] = result.path.read_bytes()
            for v in result.violations:
                lines.append(f"{name}:{v.line}:{v.column} {v.rule} {v.message}{' [*]' if v.fixable else ''}")
        outcome.check = "".join(line + "\n" for line in lines)
        return outcome


def _run_cli(case: Case) -> Outcome:
    assert case.cli is not None
    with tempfile.TemporaryDirectory() as tmp:
        # Resolved, because the command prints resolved paths: on Windows the
        # temporary directory can be named in 8.3 form (C:\Users\RUNNER~1\...).
        work = Path(tmp).resolve()
        if case.input_path().exists():
            target = work / case.file
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(case.input_path(), target)
        extra = case.path / "files"
        if extra.is_dir():
            shutil.copytree(extra, work, dirs_exist_ok=True)
        args = [str(a).replace("{file}", case.file) for a in case.cli["args"]]
        stdout, stderr = io.StringIO(), io.StringIO()
        stdin = io.StringIO(case.cli.get("stdin", ""))
        previous = os.getcwd()
        os.chdir(work)
        try:
            code = cli_main(args, stdout=stdout, stderr=stderr, stdin=stdin)
        finally:
            os.chdir(previous)
        text = stdout.getvalue()
        if case.cli.get("stderr", False):
            text += "--- stderr ---\n" + stderr.getvalue()
        text = text.replace(str(work) + os.sep, "").replace(str(work), ".")
        outcome = Outcome(stdout=text.replace("\\", "/"), exit_code=code)
        written = work / case.file
        if written.exists() and case.input_path().exists():
            data = written.read_bytes()
            if data != case.input_path().read_bytes():
                outcome.output = data
        return outcome


def bless(case: Case) -> list[str]:
    """Write a case's expected files from what it produces now."""
    outcome = run_case(case)
    written: list[str] = []

    def put(path: Path, data: bytes | None) -> None:
        if data is None:
            if path.exists():
                path.unlink()
                written.append(f"removed {path.name}")
            return
        if not path.exists() or path.read_bytes() != data:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)
            written.append(f"wrote {path.relative_to(case.path).as_posix()}")

    if case.cli is not None:
        put(case.path / "stdout.txt", (outcome.stdout or "").encode("utf-8"))
        put(case.output_path(), outcome.output)
        return written
    if case.project:
        for name, data in outcome.outputs.items():
            source = (case.path / "input" / name).read_bytes()
            put(case.path / "output" / name, data if data != source else None)
    else:
        source = case.input_path().read_bytes()
        put(case.output_path(), outcome.output if outcome.output != source else None)
    put(case.path / "check.txt", outcome.check.encode("utf-8"))
    return written
