"""Compare pyPrettyVBA's `xlide` preset with XLIDE's own Format Document.

XLIDE (the VS Code extension) ships a formatter of its own; the `xlide`
preset is meant to lay code out the way it does. This tool runs both over
the same modules and reports every line they format differently, grouped by
a short description of the difference, so a disagreement is either a known,
documented choice or a bug in one of them.

    python tools/xlide_differential.py PATH [PATH ...] [--xlide-root DIR] [--show N]

XLIDE is read from --xlide-root, the XLIDE_ROOT environment variable, or a
sibling checkout (../xlide/xlide_vscode, then ../xlide_vscode). Its
formatter is bundled with that checkout's own esbuild into artifacts/xlide/,
so nothing is installed. Needs Node.js.
"""

from __future__ import annotations

import argparse
import collections
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from pyprettyvba import Config, collect_files, format_source  # noqa: E402
from pyprettyvba.document import split_header  # noqa: E402
from pyprettyvba.textio import decode  # noqa: E402


def xlide_root(explicit: str | None) -> Path:
    candidates = [explicit, os.environ.get("XLIDE_ROOT")]
    candidates += [str(REPO.parent / "xlide" / "xlide_vscode"), str(REPO.parent / "xlide_vscode")]
    for candidate in candidates:
        if candidate and (Path(candidate) / "src" / "analyzer" / "format" / "formatModule.ts").is_file():
            return Path(candidate)
    raise SystemExit("No XLIDE checkout with src/analyzer/format/formatModule.ts; pass --xlide-root.")


def bundle(root: Path) -> Path:
    out = REPO / "artifacts" / "xlide" / "formatModule.mjs"
    out.parent.mkdir(parents=True, exist_ok=True)
    esbuild = root / "node_modules" / ".bin" / ("esbuild.cmd" if os.name == "nt" else "esbuild")
    subprocess.run(
        [str(esbuild), "src/analyzer/format/formatModule.ts", "--bundle", "--platform=node",
         "--format=esm", f"--outfile={out}", "--log-level=warning"],
        cwd=root, check=True,
    )
    return out


def xlide_format(bundle_path: Path, sources: dict[str, str]) -> dict[str, dict[str, str]]:
    with tempfile.TemporaryDirectory() as tmp:
        requests = Path(tmp) / "requests.json"
        results = Path(tmp) / "results.json"
        requests.write_text(json.dumps([{"id": k, "source": v} for k, v in sources.items()]), encoding="utf-8")
        subprocess.run(
            ["node", str(REPO / "tools" / "xlide_format.mjs"), str(bundle_path), str(requests), str(results)],
            check=True,
        )
        return {item["id"]: item for item in json.loads(results.read_text(encoding="utf-8"))}


def classify(ours: str, theirs: str) -> str:
    """A short description of how two versions of one line differ."""
    if ours.strip() == theirs.strip():
        return "indentation"
    if ours.rstrip() == theirs.rstrip():
        return "trailing whitespace"
    if ours.replace(" ", "") == theirs.replace(" ", ""):
        return "spacing"
    if ours.lower() == theirs.lower():
        return "letter case"
    return "other"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--xlide-root")
    parser.add_argument("--show", type=int, default=3, help="examples to print per kind of difference")
    args = parser.parse_args(argv)

    root = xlide_root(args.xlide_root)
    bundle_path = bundle(root)
    # The preset without identifier casing: XLIDE takes its casing from a
    # project index this comparison does not build.
    config = Config({"preset": "xlide", "rules": {"identifier-case": False}})
    files = collect_files(args.paths, Config())
    bodies: dict[str, str] = {}
    ours: dict[str, str] = {}
    for path in files:
        text = decode(path.read_bytes()).text
        _header, body = split_header(text)
        bodies[str(path)] = body
        ours[str(path)] = format_source(body, config, path=path).output
    theirs = xlide_format(bundle_path, bodies)

    kinds: collections.Counter[str] = collections.Counter()
    examples: dict[str, list[str]] = collections.defaultdict(list)
    refused = 0
    for key in bodies:
        result = theirs[key]
        if "text" not in result:
            refused += 1
            print(f"XLIDE declined {key}: {result.get('refusal') or result.get('error')}")
            continue
        a = ours[key].splitlines()
        b = result["text"].splitlines()
        if len(a) != len(b):
            kinds["line count"] += 1
            continue
        for number, (x, y) in enumerate(zip(a, b, strict=True), start=1):
            if x == y:
                continue
            kind = classify(x, y)
            kinds[kind] += 1
            if len(examples[kind]) < args.show:
                examples[kind].append(f"{key}:{number}\n    pyprettyvba: {x!r}\n    xlide:       {y!r}")
    print(f"{len(files)} files, {refused} declined by XLIDE")
    for kind, count in kinds.most_common():
        print(f"\n{count} lines differ in {kind}:")
        for example in examples[kind]:
            print("  " + example)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
