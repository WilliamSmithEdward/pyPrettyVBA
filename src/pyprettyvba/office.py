"""VBA inside Office files, through pyOpenVBA.

pyOpenVBA reads the VBA project out of an Excel workbook, Word document,
PowerPoint presentation or Access database, and writes edited modules
back so that the file still opens cleanly in its application. This module
is the thin layer pyPrettyVBA needs on top of it: open a file by its
extension, list its modules with their kinds and code, and save it.

Saving never writes the file in place. The file is saved beside the
original first, and replaces it only when the save went through cleanly,
so an interrupted or refused save leaves the original as it was. A project
with a digital signature is refused, since any edit invalidates the
signature, unless the caller asks for the signature to be removed. A
password-protected project is written: its password material is kept as it
was.

A zip-based file (.xlsm, .docm, .pptm and the like) keeps the signature of
its VBA project in parts of its own, beside vbaProject.bin and named by its
relationships; Office's own signed add-ins carry three, the legacy, agile
and V3 signatures. pyOpenVBA looks for signatures inside vbaProject.bin
only, so this module finds those parts itself. The older binary formats
keep the signature elsewhere again (in a property set or a string table),
and a signed one is not recognized.
"""

from __future__ import annotations

import os
import posixpath
import re
import warnings
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote

from pyopenvba import AccessDatabase, ExcelFile, PowerPointFile, VBAModuleKind, WordFile
from pyopenvba.exceptions import PyOpenVBAError

__all__ = [
    "FILE_ERRORS",
    "OFFICE_SUFFIXES",
    "OfficeModule",
    "SignedProjectError",
    "is_office_file",
    "modules_of",
    "open_host",
    "replace",
    "save_beside",
    "signature_parts",
]

# The file extensions pyOpenVBA reads and writes VBA in, and the class that does it.
OFFICE_SUFFIXES: dict[str, Any] = {
    ".xlsm": ExcelFile,
    ".xlsb": ExcelFile,
    ".xlam": ExcelFile,
    ".xls": ExcelFile,
    ".docm": WordFile,
    ".dotm": WordFile,
    ".doc": WordFile,
    ".pptm": PowerPointFile,
    ".potm": PowerPointFile,
    ".ppt": PowerPointFile,
    ".accdb": AccessDatabase,
    ".mdb": AccessDatabase,
}


# What reading or saving a file can raise: a file that is missing or locked,
# that is not a zip package or compound file, or that has no VBA project.
FILE_ERRORS: tuple[type[BaseException], ...] = (PyOpenVBAError, OSError, zipfile.BadZipFile)


class SignedProjectError(Exception):
    """The VBA project is digitally signed; editing it would invalidate the signature."""


_SIGNED = "the VBA project is digitally signed, and editing it would invalidate the signature"

# The relationship types that point from vbaProject.bin to its signatures.
_SIGNATURE_TYPES = (b"/vbaProjectSignature", b"/vbaProjectSignatureAgile", b"/vbaProjectSignatureV3")
_RELATIONSHIP_RE = re.compile(rb"<Relationship\b[^>]*?(?:/>|>\s*</Relationship\s*>)")
_OVERRIDE_RE = re.compile(rb"<Override\b[^>]*?(?:/>|>\s*</Override\s*>)")
_ATTRIBUTE_RE = re.compile(rb"""([\w:]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")


@dataclass(frozen=True)
class OfficeModule:
    """One module of the VBA project in an Office file."""

    name: str
    # "standard" for a standard module; "class" for a class, document or
    # form module, which pyOpenVBA does not tell apart.
    kind: str
    source: str


def is_office_file(path: str | Path) -> bool:
    return Path(path).suffix.lower() in OFFICE_SUFFIXES


def open_host(path: str | Path) -> Any:
    """Open an Office file with the pyOpenVBA class for its extension."""
    path = Path(path)
    host_class = OFFICE_SUFFIXES.get(path.suffix.lower())
    if host_class is None:
        raise ValueError(f"{path}: not an Office file pyPrettyVBA can read VBA from")
    return host_class(path)


def modules_of(host: Any) -> list[OfficeModule]:
    """The modules of an open file, in the project's order."""
    project = host.vba_project()
    listed = project.modules() if callable(project.modules) else project.modules
    kinds = {module.name: _kind(getattr(module, "kind", None)) for module in listed}
    return [OfficeModule(name, kinds.get(name, "standard"), host.get_module(name)) for name in host.module_names()]


def _kind(kind: object) -> str:
    if kind is None or kind == VBAModuleKind.standard or kind == "module":
        return "standard"
    return "class"


def signature_parts(path: str | Path) -> list[str]:
    """The parts of a zip-based Office file that hold the signatures of its VBA project."""
    try:
        with zipfile.ZipFile(path) as package:
            return [part for _rels, _element, part in _signature_links(package)]
    except zipfile.BadZipFile:
        return []


def _signature_links(package: zipfile.ZipFile) -> list[tuple[str, bytes, str]]:
    """(relationships part, relationship element, signature part) for each signature."""
    names = {name.lower(): name for name in package.namelist()}
    links: list[tuple[str, bytes, str]] = []
    for name in package.namelist():
        folder, base = posixpath.split(name)
        rels = names.get(posixpath.join(folder, "_rels", base + ".rels").lower())
        if base.lower() != "vbaproject.bin" or rels is None:
            continue
        for element in _RELATIONSHIP_RE.findall(package.read(rels)):
            attributes = _attributes(element)
            if attributes.get(b"TargetMode") == b"External":
                continue
            if not attributes.get(b"Type", b"").endswith(_SIGNATURE_TYPES):
                continue
            target = unquote(attributes.get(b"Target", b"").decode("utf-8", "replace"))
            if target.startswith("/"):
                part = target.lstrip("/")
            else:
                part = posixpath.normpath(posixpath.join(folder, target))
            links.append((rels, element, names.get(part.lower(), part)))
    return links


def _attributes(element: bytes) -> dict[bytes, bytes]:
    return {key: double or single for key, double, single in _ATTRIBUTE_RE.findall(element)}


def _remove_signature_parts(path: Path) -> None:
    """Remove the signature parts of a zip-based file, with their relationships and content types.

    What is left is the package Office writes for an unsigned project.
    """
    with zipfile.ZipFile(path) as package:
        links = _signature_links(package)
        entries = [(info, package.read(info)) for info in package.infolist()]
    parts = {part.lower() for _rels, _element, part in links}
    elements: dict[str, list[bytes]] = {}
    for rels, element, _part in links:
        elements.setdefault(rels, []).append(element)
    with zipfile.ZipFile(path, "w") as package:
        for info, data in entries:
            name = info.filename
            if name.lower() in parts:
                continue
            if name in elements:
                for element in elements[name]:
                    data = data.replace(element, b"", 1)
                if not _RELATIONSHIP_RE.search(data):
                    continue  # no relationships left: an unsigned project has no such part
            elif name == "[Content_Types].xml":
                data = _OVERRIDE_RE.sub(lambda m: b"" if _part_name(m.group(0)) in parts else m.group(0), data)
            package.writestr(info, data)


def _part_name(override: bytes) -> str:
    return _attributes(override).get(b"PartName", b"").decode("utf-8", "replace").lstrip("/").lower()


def save_beside(host: Any, path: Path, *, remove_signatures: bool = False) -> tuple[Path, bool]:
    """Save the open file next to ``path``; return where, and whether a signature was removed.

    A digitally signed project raises SignedProjectError, and nothing is
    written, unless ``remove_signatures``: the saved file then has no
    signature, since the one it had would not match the edited code.
    """
    signed = bool(signature_parts(path))
    if signed and not remove_signatures:
        raise SignedProjectError(_SIGNED)
    temporary = path.with_name(f".{path.stem}.pyprettyvba{path.suffix}")
    try:
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            host.save(temporary, allow_protected=True)
        # pyOpenVBA drops the signature streams it finds inside vbaProject.bin, and warns.
        dropped = any("signature" in str(warning.message).lower() for warning in caught)
        if dropped and not remove_signatures:
            raise SignedProjectError(_SIGNED)
        if signed:
            _remove_signature_parts(temporary)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise
    return temporary, signed or dropped


def replace(temporary: Path, path: Path) -> None:
    """Move a file saved by ``save_beside`` over the original."""
    try:
        os.replace(temporary, path)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise
