"""Block structure: levels, conditional compilation, and broken regions."""

from __future__ import annotations

from pyprettyvba.document import Document
from pyprettyvba.structure import analyze_structure


def levels(src: str, **options: bool) -> list[int]:
    return analyze_structure(Document(src), **options).levels


def test_blocks() -> None:
    src = (
        "Sub A()\n"
        "If x Then\n"
        "For i = 1 To 2\n"
        "y\n"
        "Next i\n"
        "ElseIf z Then\n"
        "Select Case q\n"
        "Case 1\n"
        "w\n"
        "End Select\n"
        "End If\n"
        "End Sub\n"
    )
    assert levels(src) == [0, 1, 2, 3, 2, 1, 2, 3, 4, 2, 1, 0]
    assert levels(src, case_arms=False) == [0, 1, 2, 3, 2, 1, 2, 2, 3, 2, 1, 0]
    assert levels(src, procedure_body=False) == [0, 0, 1, 2, 1, 0, 1, 2, 3, 1, 0, 0]


def test_one_line_blocks_net_to_zero() -> None:
    assert levels("Sub A(): x = 1: End Sub\nFor i = 1 To 3: y: Next\nz\n") == [0, 0, 0]


def test_next_closes_several_loops() -> None:
    assert levels("For i = 1 To 2\nFor j = 1 To 2\nx\nNext j, i\ny\n") == [0, 1, 2, 0, 0]


def test_alternative_procedure_headers_in_directive_arms() -> None:
    src = (
        "#If VBA7 Then\n"
        "Private Sub Tick(ByVal ms As LongPtr)\n"
        "#Else\n"
        "Private Sub Tick(ByVal ms As Long)\n"
        "#End If\n"
        "x = ms\n"
        "End Sub\n"
    )
    structure = analyze_structure(Document(src))
    assert structure.problems == []
    # The arms are unbalanced, so they take no extra level.
    assert structure.levels == [0, 0, 0, 0, 0, 1, 0]


def test_balanced_directive_block_indents_its_content() -> None:
    src = "#If VBA7 Then\nPrivate Declare PtrSafe Sub S Lib \"k\" ()\n#Else\nPrivate Declare Sub S Lib \"k\" ()\n#End If\n"
    assert levels(src) == [0, 1, 0, 1, 0]
    assert levels(src, directive_arms=False) == [0, 0, 0, 0, 0]


def test_broken_regions_are_reported() -> None:
    structure = analyze_structure(Document("Sub A()\nEnd If\nx\nEnd Sub\nSub B()\ny\nEnd Sub\n"))
    assert structure.problems and "End If has nothing to close" in structure.problems[0][1]
    assert structure.is_broken(1)
    assert not structure.is_broken(5)


def test_unclosed_block_at_end() -> None:
    structure = analyze_structure(Document("Sub A()\nIf x Then\ny\nEnd Sub\n"))
    assert any("never closed" in message for _line, message in structure.problems)
    assert structure.is_broken(2)


def test_end_function_closes_a_property() -> None:
    # The VBE accepts any procedure closer (XLIDE #81).
    structure = analyze_structure(Document("Property Get P()\nP = 1\nEnd Function\nx\n"))
    assert structure.problems == []
    assert structure.levels == [0, 1, 0, 0]
