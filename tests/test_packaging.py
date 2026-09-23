"""The shipped package uses nothing outside the standard library.

Third-party packages are for development only (the `dev` and `verify`
extras): the formatter itself must install and run with Python alone.
"""

from __future__ import annotations

import ast
import sys
import tomllib
from collections.abc import Iterator
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
PACKAGE = REPO / "src" / "pyprettyvba"


def _imported_modules(path: Path) -> Iterator[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                yield alias.name.split(".")[0]
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            yield node.module.split(".")[0]


def test_the_package_imports_only_the_standard_library() -> None:
    foreign: dict[str, list[str]] = {}
    for path in sorted(PACKAGE.rglob("*.py")):
        for name in _imported_modules(path):
            if name != "pyprettyvba" and name not in sys.stdlib_module_names:
                foreign.setdefault(name, []).append(path.relative_to(REPO).as_posix())
    assert not foreign, f"third-party imports in the package: {foreign}"


def test_the_package_declares_no_dependencies() -> None:
    project = tomllib.loads((REPO / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    assert project.get("dependencies", []) == []
