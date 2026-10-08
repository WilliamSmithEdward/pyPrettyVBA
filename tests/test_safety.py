"""The safety check: what counts as the same code, and what does not."""

from __future__ import annotations

import pytest

from pyprettyvba.safety import first_difference

SAME = [
    ("x=1+2", "x = 1 + 2"),
    ("dim x as long", "Dim x As Long"),
    ("x = 1.0", "x = 1#"),
    ("x = &hff", "x = &HFF"),
    ("d = #2020-01-15#", "d = #1/15/2020#"),
    ("x = 1: y = 2", "x = 1\r\ny = 2"),
    ("Let x = 1", "x = 1"),
    ("If x Then\r\ny\r\nEndIf", "If x Then\r\ny\r\nEnd If"),
    ("Rem note", "' note"),
    ("'note", "' note"),
    ("x = 1 + _\r\n2", "x = 1 + 2"),
    ("If a => b Then c", "If a >= b Then c"),
    ("Call Foo()", "Call Foo"),
    # A Call is the plain call it means.
    ("Call Foo(a, b)", "Foo a, b"),
    ("Call Reset", "Reset"),
    ("Call Save()", "Save"),
    ("Call Log((x))", "Log (x)"),
    ('Call ws.Range("A1").Select', 'ws.Range("A1").Select'),
    ('s = "a"&"b"', 's = "a" & "b"'),
    ("x = (y)^2", "x = (y) ^ 2"),
    ("Label1: x = 1", "Label1:\r\nx = 1"),
    ("Else x = 1", "Else: x = 1"),
    ("Foo , 2", "Foo, 2"),
    # A declaration of several variables is one declaration each.
    ("Dim a As Long, b As Long", "Dim a As Long\r\nDim b As Long"),
    ("Dim x, y As Long", "Dim x\r\nDim y As Long"),
    ("Public WithEvents app As Application, name As String", "Public WithEvents app As Application\r\nPublic name As String"),
    ("Dim buffer(1 To 10) As Byte, flag As Boolean", "Dim buffer(1 To 10) As Byte\r\nDim flag As Boolean"),
    ("Const A = 1, B As String = \"b\"", "Const A = 1\r\nConst B As String = \"b\""),
    # A single-line If is the block If it means.
    ("If x Then Exit Sub", "If x Then\r\nExit Sub\r\nEnd If"),
    ("If x Then\r\n\r\nExit Sub\r\n\r\nEnd If", "If x Then Exit Sub"),
    ("If a Then b: c Else d", "If a Then\r\nb\r\nc\r\nElse\r\nd\r\nEnd If"),
    ("If a Then b Else c: d", "If a Then\r\nb\r\nElse\r\nc\r\nd\r\nEnd If"),
    # A comment on an Attribute line compares by its text, like any comment:
    # trailing-whitespace trims the space after it.
    ('Attribute VB_Name = "M" \'note ', 'Attribute VB_Name = "M" \'note'),
]


@pytest.mark.parametrize("before, after", SAME)
def test_allowed_changes(before: str, after: str) -> None:
    assert first_difference(before, after) is None


DIFFERENT = [
    # A single-line If owns its colons: splitting it changes what runs.
    ("If a Then b: c", "If a Then b\r\nc"),
    # A glued `&` is a type character; spaced, a concatenation.
    ("s = s&t", "s = s & t"),
    # `Foo .Bar` passes a With member; `Foo.Bar` calls a member.
    ("Foo .Bar", "Foo.Bar"),
    # A bang must stay glued.
    ("x = rs!Field", "x = rs! Field"),
    # A label must stay a label.
    ("x = 1: Foo: y = 2", "x = 1\r\nFoo:\r\ny = 2"),
    # The name of a Declare without Alias is the DLL entry point.
    (
        'Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
        'Declare PtrSafe Function gettickcount Lib "kernel32" () As Long',
    ),
    # Values and types of literals.
    ("x = 3.141592653589793", "x = 3.14159265358979"),
    ("x = &HFFFF", "x = &HFFFF&"),
    # Strings and comments.
    ('s = "a  b"', 's = "a b"'),
    ("' a  b", "' a b"),
    ("x = 1", "x = 2"),
    # The parentheses around one argument are its own: they pass it by value.
    ("Call Log((x))", "Log x"),
    # A Call that calls nothing is left to the compiler as written.
    ("Call Call Sub", "Call Sub"),
    ("Call Foo(a, b)", "Foo (a), b"),
    # Splitting a declaration must keep each item's own type: in
    # `Dim x, y As Long` only y is Long.
    ("Dim x, y As Long", "Dim x As Long\r\nDim y As Long"),
    # WithEvents belongs to one variable, not to the whole list.
    ("Public WithEvents app As Application, name As String", "Public WithEvents app As Application\r\nPublic WithEvents name As String"),
    # A ReDim is one operation; two ReDims resize at two moments.
    ("ReDim a(5), b(6)", "ReDim a(5)\r\nReDim b(6)"),
    # A comma inside bounds is not an item boundary.
    ("Dim grid(1 To 2, 1 To 3) As Long", "Dim grid(1 To 2)\r\nDim (1 To 3) As Long"),
    # A block If with two statements is not one single-line If and a statement.
    ("If x Then\r\na\r\nb\r\nEnd If", "If x Then a\r\nb"),
    # The Else arm belongs inside the If.
    ("If a Then b Else c", "If a Then b\r\nc"),
]


@pytest.mark.parametrize("before, after", DIFFERENT)
def test_meaning_changes_are_caught(before: str, after: str) -> None:
    assert first_difference(before, after) is not None


def test_declare_with_alias_can_be_recased() -> None:
    before = 'Declare PtrSafe Function tick Lib "kernel32" Alias "GetTickCount" () As Long'
    after = 'Declare PtrSafe Function Tick Lib "kernel32" Alias "GetTickCount" () As Long'
    assert first_difference(before, after) is None
