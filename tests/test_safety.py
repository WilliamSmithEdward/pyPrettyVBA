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
    ('s = "a"&"b"', 's = "a" & "b"'),
    ("x = (y)^2", "x = (y) ^ 2"),
    ("Label1: x = 1", "Label1:\r\nx = 1"),
    ("Else x = 1", "Else: x = 1"),
    ("Foo , 2", "Foo, 2"),
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
]


@pytest.mark.parametrize("before, after", DIFFERENT)
def test_meaning_changes_are_caught(before: str, after: str) -> None:
    assert first_difference(before, after) is not None


def test_declare_with_alias_can_be_recased() -> None:
    before = 'Declare PtrSafe Function tick Lib "kernel32" Alias "GetTickCount" () As Long'
    after = 'Declare PtrSafe Function Tick Lib "kernel32" Alias "GetTickCount" () As Long'
    assert first_difference(before, after) is None
