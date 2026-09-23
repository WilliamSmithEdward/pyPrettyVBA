"""The document model: headers, logical lines, statements, classification."""

from __future__ import annotations

from pyprettyvba.document import Document, LineKind, StatementKind, split_header

CLASS_EXPORT = (
    "VERSION 1.0 CLASS\r\n"
    "BEGIN\r\n"
    "  MultiUse = -1  'True\r\n"
    "END\r\n"
    'Attribute VB_Name = "Class1"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Option Explicit\r\n"
)

FORM_EXPORT = (
    "VERSION 5.00\r\n"
    "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\r\n"
    '   Caption         =   "UserForm1"\r\n'
    "   Begin VB.CommandButton Command1\r\n"
    "      BeginProperty Font\r\n"
    "      EndProperty\r\n"
    "   End\r\n"
    "End\r\n"
    'Attribute VB_Name = "UserForm1"\r\n'
    "Private Sub Command1_Click()\r\n"
    "End Sub\r\n"
)


def test_class_header() -> None:
    header, body = split_header(CLASS_EXPORT)
    assert header + body == CLASS_EXPORT
    assert body == "Option Explicit\r\n"


def test_form_header_with_nested_designer_blocks() -> None:
    header, body = split_header(FORM_EXPORT)
    assert body.startswith("Private Sub Command1_Click()")
    assert header.endswith('Attribute VB_Name = "UserForm1"\r\n')


def test_plain_module_keeps_only_attribute_lines_as_header() -> None:
    header, body = split_header('Attribute VB_Name = "M"\nSub A()\nEnd Sub\n')
    assert header == 'Attribute VB_Name = "M"\n'
    assert split_header("Sub A()\n") == ("", "Sub A()\n")


def kinds_of(src: str) -> list[list[StatementKind]]:
    doc = Document(src)
    return [[s.kind for s in line.statements] for line in doc.lines]


def test_single_line_if_owns_the_rest_of_the_line() -> None:
    doc = Document("If a Then b: c Else d: e\n")
    assert [s.kind for s in doc.lines[0].statements] == [StatementKind.IF_SINGLE]


def test_if_then_colon_is_single_line() -> None:
    # Measured: `If False Then:` followed by a statement runs the statement.
    assert kinds_of("If x Then:\n")[0] == [StatementKind.IF_SINGLE]
    assert kinds_of("If x Then ' note\n")[0] == [StatementKind.IF_BLOCK]


def test_block_else_and_elseif_end_where_their_keywords_do() -> None:
    assert kinds_of("Else If d Then\n")[0] == [StatementKind.ELSE, StatementKind.IF_BLOCK]
    assert kinds_of("ElseIf c Then x = 1: y = 2\n")[0] == [
        StatementKind.ELSEIF, StatementKind.OTHER, StatementKind.OTHER,
    ]


def test_labels_and_line_numbers() -> None:
    doc = Document("ErrHandler: Resume Next\n10 x = 1\n20: y = 2\nBeep:\n")
    assert [doc.tokens[line.label].text for line in doc.lines] == ["ErrHandler", "10", "20", "Beep"]
    # An identifier and a colon at the start of a line is always a label.
    assert doc.lines[3].statements == []


def test_classification() -> None:
    src = (
        "Private Declare PtrSafe Function F Lib \"k\" () As Long\n"
        "Public Static Sub S()\n"
        "Friend Property Get P() As Long\n"
        "Private Type T\n"
        "Private Enum LongPtr\n"
        "Static n As Long\n"
        "Next i, j\n"
        "Select Case x\n"
        "Case 1\n"
        "DefInt A-Z\n"
        "Public Event Changed()\n"
        "End\n"
    )
    doc = Document(src)
    got = [line.statements[0].kind for line in doc.lines]
    assert got == [
        StatementKind.DECLARE, StatementKind.PROC_START, StatementKind.PROC_START,
        StatementKind.TYPE_START, StatementKind.ENUM_START, StatementKind.VARIABLES,
        StatementKind.NEXT, StatementKind.SELECT, StatementKind.CASE, StatementKind.DEFTYPE,
        StatementKind.EVENT, StatementKind.OTHER,
    ]
    assert doc.lines[6].statements[0].count == 2


def test_line_kinds_and_logical_lines() -> None:
    doc = Document("x = 1 + _\n    2\n' note\n\n#If A Then\nAttribute X.VB_UserMemId = 0\n")
    assert [line.kind for line in doc.lines] == [
        LineKind.CODE, LineKind.COMMENT, LineKind.BLANK, LineKind.DIRECTIVE, LineKind.ATTRIBUTE,
    ]
    assert (doc.lines[0].first_physical, doc.lines[0].last_physical) == (0, 1)


def test_member_names_are_not_statement_keywords() -> None:
    doc = Document("wb.Close\n.Print x\n")
    assert all(s.kind is StatementKind.OTHER for line in doc.lines for s in line.statements)
