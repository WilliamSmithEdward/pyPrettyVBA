"""Build the built-in name table identifier-case uses.

The VBE writes a name from the VBA library or an Office host library with
that library's spelling: `msgbox` becomes `MsgBox`, `vbcrlf` becomes
`vbCrLf`, `range` becomes `Range`. The spellings come from pyVBAanalysis's
vendored data, which XLIDE generates from the type libraries and verifies
against Office (see pyvbaanalysis/data/manifest.json for the XLIDE commit).
This script reads that data and writes src/pyprettyvba/data/names.json.gz.

    python tools/build_names.py [--source PATH_TO_pyvbaanalysis_data]

The output records the pyVBAanalysis version and XLIDE commit it came from.
Re-run it after a pyVBAanalysis sync and review the diff of the counts.
"""

from __future__ import annotations

import argparse
import gzip
import json
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
OUT = REPO / "src" / "pyprettyvba" / "data" / "names.json.gz"

HOSTS = {
    "excel": "excel_host_model.json",
    "word": "word_host_model.json",
    "powerpoint": "powerpoint_host_model.json",
    "access": "access_host_model.json",
}

# VBA library classes the runtime tables do not list as functions.
VBA_CLASSES = ("Collection", "ErrObject", "VBA")

# Compiler constants, spelled as the VBE writes them (tests/oracle: `#if vba6
# or win16 or win32 or mac` becomes `#If VBA6 Or Win16 Or Win32 Or Mac`).
# TWINBASIC is not one: the VBE leaves `twinbasic` as written.
COMPILER_CONSTANTS = ("VBA6", "VBA7", "Win16", "Win32", "Win64", "Mac")

# Names the vendored data spells differently from the VBE, with the VBE's
# spelling as measured (tests/oracle/vbe_rendering.json, probe
# i-vba-library-members: the VBE writes `Err.LastDllError`).
MEASURED_SPELLINGS = {"lastdllerror": "LastDllError"}


def _measured(name: str) -> str:
    return MEASURED_SPELLINGS.get(name.lower(), name)


def _default_source() -> Path:
    import pyvbaanalysis

    return Path(pyvbaanalysis.__file__).resolve().parent / "data"


def build(source: Path) -> dict[str, object]:
    runtime = json.loads((source / "vba_runtime_tables.json").read_text(encoding="utf-8"))
    manifest = json.loads((source / "manifest.json").read_text(encoding="utf-8"))
    vba: list[str] = []
    vba.extend(f["name"] for f in runtime["functions"])
    vba.extend(c["name"] for c in runtime["constants"])
    vba.extend(o["name"] for o in runtime["objects"])
    vba.extend(VBA_CLASSES)
    members = {
        o["name"].lower(): sorted({_measured(m["name"]) for m in o.get("members", [])})
        for o in runtime["objects"]
    }
    hosts: dict[str, list[str]] = {}
    for host, file_name in HOSTS.items():
        model = json.loads((source / file_name).read_text(encoding="utf-8"))
        names: set[str] = set(model.get("globals", {}))
        global_type = model.get("globalType")
        types = model.get("types", {})
        if global_type and global_type in types:
            names.update(m["name"] for m in types[global_type].get("members", []))
        names.update(model.get("constants", {}))
        for enum in model.get("enums", {}).values():
            if enum.get("displayName"):
                names.add(enum["displayName"])
        for type_model in types.values():
            if type_model.get("displayName"):
                names.add(type_model["displayName"])
        host_name = model.get("hostName")
        if host_name:
            names.add(host_name)
        hosts[host] = sorted(n for n in names if n and not n.startswith("_") and n.replace("_", "").isalnum())
    import pyvbaanalysis

    library: dict[str, dict[str, str]] = {"vba": _library_vba(runtime)}
    for host, file_name in HOSTS.items():
        library[host] = _library_host(json.loads((source / file_name).read_text(encoding="utf-8")))
    return {
        "source": {
            "pyvbaanalysis": pyvbaanalysis.__version__,
            "xlideVersion": manifest.get("xlideVersion"),
            "xlideCommit": manifest.get("xlideCommit"),
        },
        "vba": sorted(set(vba)),
        "members": members,
        "compiler": list(COMPILER_CONSTANTS),
        "hosts": hosts,
        "library": library,
    }


def _usable(name: str) -> bool:
    return bool(name) and not name.startswith("_") and name.replace("_", "").isalnum()


def _resolve(entries: list[tuple[int, str]]) -> dict[str, str]:
    """Lower-case name -> spelling, from (rank, spelling) pairs.

    The lowest rank that names a word decides its spelling. Two spellings at
    that rank make the word ambiguous: it maps to "" and is left as written.
    """
    ranked: dict[str, tuple[int, set[str]]] = {}
    for rank, raw in entries:
        spelling = _measured(raw)
        if not _usable(spelling):
            continue
        key = spelling.lower()
        best = ranked.get(key)
        if best is None or rank < best[0]:
            ranked[key] = (rank, {spelling})
        elif rank == best[0]:
            best[1].add(spelling)
    return {key: (next(iter(s)) if len(s) == 1 else "") for key, (_rank, s) in sorted(ranked.items())}


def _library_vba(runtime: dict[str, object]) -> dict[str, str]:
    """Every name the VBA library defines: functions, constants, objects, members."""
    entries: list[tuple[int, str]] = []
    for key in ("functions", "constants", "objects"):
        for item in runtime[key]:  # type: ignore[attr-defined]
            entries.append((0, item["name"]))
            for member in item.get("members", []):
                entries.append((0, member["name"]))
    entries.extend((0, name) for name in VBA_CLASSES)
    return _resolve(entries)


def _library_host(model: dict[str, object]) -> dict[str, str]:
    """Every name a host's libraries define, the host's own before Office's.

    These are the names the VBE's project name table holds for a project
    that references the host (measured: an undeclared `sql` is written `Sql`,
    after Excel's QueryTable.Sql). Parameter names are not among them: the
    VBE leaves `savechanges:=` and `prompt:=` as written.
    """
    entries: list[tuple[int, str]] = []
    entries.extend((0, name) for name in model.get("globals", {}))  # type: ignore[union-attr]
    entries.extend((0, name) for name in model.get("constants", {}))  # type: ignore[union-attr]
    for enum_name, enum in model.get("enums", {}).items():  # type: ignore[union-attr]
        # An enum's own name ranks below its constants: the VBE writes
        # `xlConstants` (the constant), not `XlConstants` (the enum).
        entries.append((1, enum.get("displayName") or enum_name.split(".")[-1]))
    for type_name, type_model in model.get("types", {}).items():  # type: ignore[union-attr]
        rank = _rank(type_name)
        entries.append((rank, type_model.get("displayName") or type_name.split(".")[-1]))
        entries.extend((rank, member["name"]) for member in type_model.get("members", []))
    host_name = model.get("hostName")
    if isinstance(host_name, str):
        entries.append((0, host_name))
    return _resolve(entries)


def _rank(qualified: str) -> int:
    """A host library's own names come before the shared Office library's.

    Measured: the VBE writes `ID`, `Filename` and `Url` (Excel's) rather than
    `Id`, `FileName` and `URL` (Office's).
    """
    return 2 if qualified.startswith("Office.") else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--source", type=Path, default=None)
    args = parser.parse_args(argv)
    table = build(args.source or _default_source())
    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(table, separators=(",", ":"), sort_keys=True).encode("utf-8")
    # mtime=0 keeps the file byte-identical across rebuilds of the same data.
    OUT.write_bytes(gzip.compress(payload, mtime=0))
    counts = {host: len(names) for host, names in table["hosts"].items()}  # type: ignore[union-attr]
    print(f"wrote {OUT} ({len(payload)} bytes raw): vba={len(table['vba'])} hosts={counts}")  # type: ignore[arg-type]
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
