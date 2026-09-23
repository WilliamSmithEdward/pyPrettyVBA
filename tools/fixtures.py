"""Maintain the fixture suite (tests/fixtures; layout in tests/fixture_support.py).

    python tools/fixtures.py list
    python tools/fixtures.py bless [CASE_PATH_PREFIX ...]

`bless` writes each case's expected files from what pyPrettyVBA does now.
It is for writing a new case or accepting a deliberate change: review the
diff it leaves before keeping it, since a blessed mistake becomes the test.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(REPO / "tests"))

from fixture_support import bless, load_cases  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("list")
    blessing = sub.add_parser("bless")
    blessing.add_argument("prefixes", nargs="*")
    args = parser.parse_args(argv)
    cases = load_cases()
    if args.command == "list":
        for case in cases:
            print(f"{case.id:55} {case.title}")
        print(f"{len(cases)} cases")
        return 0
    for case in cases:
        if args.prefixes and not any(case.id.startswith(p) for p in args.prefixes):
            continue
        for change in bless(case):
            print(f"{case.id}: {change}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
