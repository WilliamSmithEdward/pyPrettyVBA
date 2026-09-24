"""The VBA inside Office files: formatted in place through pyOpenVBA.

The files are made with pyOpenVBA as each test runs, so none is kept in the
repository.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile
from pathlib import Path
from typing import Any

import pytest
from pyopenvba import AccessDatabase, ExcelFile, PowerPointFile, VBAModuleKind, WordFile
from pyopenvba.cfb import CFB
from pyopenvba.exceptions import VBAProjectError

from pyprettyvba import Config, format_file, format_office_file, format_paths, format_source
from pyprettyvba.cli import main as cli_main
from pyprettyvba.office import signature_parts

MESSY = (
    "option explicit\r\n"
    "public function total(ws as worksheet) as long\r\n"
    "dim r as long\r\n"
    "for r=2 to ws.cells(ws.rows.count,1).end(xlup).row\r\n"
    "total=total+ws.cells(r,2).value\r\n"
    "next r\r\n"
    "end function\r\n"
)
TIDY_BODY = (
    "Option Explicit\r\n"
    "\r\n"
    "Public Function total(ws As Worksheet) As Long\r\n"
    "    Dim r As Long\r\n"
    "    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row\r\n"
    "        total = total + ws.Cells(r, 2).Value\r\n"
    "    Next r\r\n"
    "End Function\r\n"
)


def _put(document: Any, name: str, source: str, kind: VBAModuleKind = VBAModuleKind.standard) -> None:
    """Replace a module the new file already has, or add one."""
    if name in document.module_names():
        document.set_module(name, source)
    else:
        document.vba_project().add_module(name, source, kind=kind)


def _workbook(path: Path, modules: dict[str, str] | None = None, classes: dict[str, str] | None = None) -> Path:
    with ExcelFile.create_new(path) as wb:
        for name, source in (modules or {"Module1": MESSY}).items():
            _put(wb, name, source)
        for name, source in (classes or {}).items():
            _put(wb, name, source, VBAModuleKind.other)
        wb.save()
    return path


def _module(path: Path, name: str) -> str:
    with ExcelFile(path) as wb:
        return str(wb.get_module(name))


def test_a_workbook_is_formatted_module_by_module(tmp_path: Path) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    results = format_office_file(book, Config(), write=True)
    by_module = {r.module: r for r in results}
    assert set(by_module) == {"ThisWorkbook", "Sheet1", "Module1"}
    assert by_module["Module1"].changed and by_module["Module1"].written
    assert not by_module["Sheet1"].changed and not by_module["Sheet1"].written
    assert by_module["Module1"].label == f"{book}:Module1"
    stored = _module(book, "Module1")
    assert stored.endswith(TIDY_BODY)
    # The stored module is formatted: formatting it again changes nothing.
    assert format_source(stored, Config()).output == stored
    assert not list(tmp_path.glob(".*pyprettyvba*"))


def test_checking_writes_nothing(tmp_path: Path) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    before = book.read_bytes()
    results = format_office_file(book, Config(), write=False)
    assert any(r.changed for r in results)
    assert not any(r.written for r in results)
    assert book.read_bytes() == before


def test_the_modules_of_a_file_are_one_project(tmp_path: Path) -> None:
    book = _workbook(
        tmp_path / "Book1.xlsm",
        modules={
            "Settings": "Public Const MaxItems As Long = 10\r\n",
            "Orders": "Sub Add()\r\n    If n > MAXITEMS Then Exit Sub\r\n    Dim b As New shoppingbasket\r\nEnd Sub\r\n",
        },
        classes={"ShoppingBasket": "Public Sub Clear()\r\nEnd Sub\r\n"},
    )
    format_office_file(book, Config(), write=True)
    orders = _module(book, "Orders")
    assert "If n > MaxItems Then" in orders
    assert "As New ShoppingBasket" in orders


def test_the_code_keeps_crlf_line_endings(tmp_path: Path) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    format_office_file(book, Config({"line-ending": "lf"}), write=True)
    stored = _module(book, "Module1")
    assert "\n" not in stored.replace("\r\n", "")


@pytest.mark.parametrize(
    ("suffix", "host"),
    [(".docm", WordFile), (".pptm", PowerPointFile), (".accdb", AccessDatabase)],
)
def test_word_powerpoint_and_access_files(tmp_path: Path, suffix: str, host: Any) -> None:
    path = tmp_path / f"Sample{suffix}"
    with host.create_new(path) as document:
        _put(document, "Module1", MESSY)
        document.save()
    results = format_office_file(path, Config(), write=True)
    module = next(r for r in results if r.module == "Module1")
    assert module.written, module.error
    with host(path) as document:
        assert document.get_module("Module1").rstrip("\r\n").endswith(TIDY_BODY.rstrip("\r\n"))


def test_format_paths_takes_office_files_named_and_leaves_them_out_of_directories(tmp_path: Path) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    (tmp_path / "Module2.bas").write_bytes(b"sub a()\r\nend sub\r\n")
    in_directory = {r.label for r in format_paths([tmp_path], Config())}
    assert in_directory == {str(tmp_path / "Module2.bas")}
    named = {r.module for r in format_paths([book], Config())}
    assert named == {"ThisWorkbook", "Sheet1", "Module1"}
    included = format_paths([tmp_path], Config({"include": ["*.bas", "*.xlsm"]}))
    assert {r.module for r in included} == {None, "ThisWorkbook", "Sheet1", "Module1"}


def test_a_file_without_a_vba_project_is_an_error(tmp_path: Path) -> None:
    garbage = tmp_path / "Broken.xlsm"
    garbage.write_bytes(b"not a workbook")
    [result] = format_office_file(garbage, Config())
    assert not result.ok and "cannot read the VBA project" in (result.error or "")
    [missing] = format_office_file(tmp_path / "Missing.xlsm", Config())
    assert not missing.ok


def test_format_file_points_office_files_elsewhere(tmp_path: Path) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    result = format_file(book, Config())
    assert not result.ok and "format_office_file" in (result.error or "")


# How Office's own signed add-ins keep their VBA signatures: three parts
# beside vbaProject.bin, each with a relationship and a content type.
SIGNATURES = [
    ("rId1", "2006/relationships/vbaProjectSignature", "vbaProjectSignature.bin"),
    ("rId2", "2014/relationships/vbaProjectSignatureAgile", "vbaProjectSignatureAgile.bin"),
    ("rId3", "2020/07/relationships/vbaProjectSignatureV3", "vbaProjectSignatureV3.bin"),
]


def _sign(path: Path) -> list[str]:
    """Give a zip-based Office file signature parts; return their names."""
    with zipfile.ZipFile(path) as archive:
        entries = [(info, archive.read(info)) for info in archive.infolist()]
    folder = next(info.filename for info, _data in entries if info.filename.endswith("/vbaProject.bin"))
    folder = folder.rpartition("/")[0]
    relationships = "".join(
        f'<Relationship Id="{rid}" Type="http://schemas.microsoft.com/office/{kind}" Target="{target}"/>'
        for rid, kind, target in SIGNATURES
    )
    overrides = "".join(
        f'<Override PartName="/{folder}/{target}" ContentType="application/vnd.ms-office.{target[:-4]}"/>'
        for _rid, _kind, target in SIGNATURES
    )
    with zipfile.ZipFile(path, "w") as archive:
        for info, data in entries:
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", overrides.encode() + b"</Types>")
            archive.writestr(info, data)
        archive.writestr(
            f"{folder}/_rels/vbaProject.bin.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{relationships}</Relationships>",
        )
        for _rid, _kind, target in SIGNATURES:
            archive.writestr(f"{folder}/{target}", b"0\x82 a PKCS #7 signature")
    return [f"{folder}/{target}" for _rid, _kind, target in SIGNATURES]


@pytest.mark.parametrize(("suffix", "host"), [(".xlsm", ExcelFile), (".docm", WordFile), (".pptm", PowerPointFile)])
def test_a_signed_file_is_not_written(tmp_path: Path, suffix: str, host: Any) -> None:
    path = tmp_path / f"Signed{suffix}"
    with host.create_new(path) as document:
        _put(document, "Module1", MESSY)
        document.save()
    assert signature_parts(path) == []
    parts = _sign(path)
    assert signature_parts(path) == parts
    before = path.read_bytes()
    results = format_office_file(path, Config(), write=True)
    changed = [r for r in results if r.changed]
    assert changed and all("digitally signed" in (r.error or "") for r in changed)
    assert not any(r.written for r in results)
    assert path.read_bytes() == before
    assert not list(tmp_path.glob(".*pyprettyvba*"))


def test_signature_streams_pyopenvba_finds_are_refused_too(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    before = book.read_bytes()
    save = ExcelFile.save

    def signed_save(self: ExcelFile, dest: Any = None, **kwargs: Any) -> None:
        save(self, dest, **kwargs)
        warnings.warn("Dropped stale VBA digital signature streams (legacy).", UserWarning, stacklevel=2)

    monkeypatch.setattr(ExcelFile, "save", signed_save)
    results = format_office_file(book, Config(), write=True)
    changed = [r for r in results if r.changed]
    assert changed and all("digitally signed" in (r.error or "") for r in changed)
    assert not any(r.written for r in results)
    assert book.read_bytes() == before
    assert not list(tmp_path.glob(".*pyprettyvba*"))


def _protect(path: Path) -> None:
    """Give a workbook's VBA project a password, as pyOpenVBA reads one.

    pyOpenVBA takes a long DPB record in the PROJECT stream for a password
    hash, without checking the hash.
    """
    with zipfile.ZipFile(path) as archive:
        entries = [(info, archive.read(info)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w") as archive:
        for info, data in entries:
            if info.filename.lower().endswith("vbaproject.bin"):
                compound = CFB.from_bytes(data)
                project = compound.get_stream("PROJECT").decode("latin-1")
                project = re.sub('DPB="[0-9A-F]*"', 'DPB="' + "AB" * 36 + '"', project)
                compound.write_stream("PROJECT", project.encode("latin-1"))
                data = compound.to_bytes()
            archive.writestr(info, data)


def test_a_password_protected_project_is_written_and_keeps_its_password(tmp_path: Path) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    _protect(book)
    with ExcelFile(book) as wb:
        password = wb.vba_project().protection.dpb
        # pyOpenVBA on its own refuses to write the project.
        wb.set_module("Module1", MESSY + "' edited\r\n")
        with pytest.raises(VBAProjectError, match="password-protected"):
            wb.save(tmp_path / "Refused.xlsm")
    [module] = [r for r in format_office_file(book, Config(), write=True) if r.module == "Module1"]
    assert module.written, module.error
    assert _module(book, "Module1").endswith(TIDY_BODY)
    with ExcelFile(book) as wb:
        assert wb.vba_project().protection.dpb == password


def test_a_failed_save_leaves_the_file_as_it_was(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    book = _workbook(tmp_path / "Book1.xlsm")
    before = book.read_bytes()

    def failing_save(self: ExcelFile, dest: Any = None, **kwargs: Any) -> None:
        raise PermissionError("the file is open in Excel")

    monkeypatch.setattr(ExcelFile, "save", failing_save)
    results = format_office_file(book, Config(), write=True)
    assert any("cannot write" in (r.error or "") for r in results)
    assert book.read_bytes() == before


def _cli(args: list[str], cwd: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[int, str, str]:
    monkeypatch.chdir(cwd)
    out, err = io.StringIO(), io.StringIO()
    code = cli_main(args, stdout=out, stderr=err, stdin=io.StringIO())
    return code, out.getvalue(), err.getvalue()


def test_the_command_line_names_modules(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _workbook(tmp_path / "Book1.xlsm")
    code, out, err = _cli(["check", "--isolated", "Book1.xlsm"], tmp_path, monkeypatch)
    assert code == 1
    # Line 1 of the module is its `Attribute VB_Name` line.
    assert out.splitlines()[0].startswith("Book1.xlsm:Module1:2:1: keyword-case")
    code, out, err = _cli(["format", "--isolated", "--diff", "Book1.xlsm"], tmp_path, monkeypatch)
    assert code == 1 and "--- a/Book1.xlsm:Module1" in out
    assert "1 file would be reformatted" in err
    code, out, err = _cli(["format", "--isolated", "Book1.xlsm"], tmp_path, monkeypatch)
    assert code == 0 and "1 file reformatted" in err
    code, out, err = _cli(["check", "--isolated", "Book1.xlsm"], tmp_path, monkeypatch)
    assert code == 0 and out == ""
