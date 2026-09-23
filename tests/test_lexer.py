"""The lexer: lossless, and reading VBA the way the VBE reads it."""

from __future__ import annotations

import pytest

from pyprettyvba.lexer import TokenKind, tokenize


def kinds(src: str) -> list[tuple[str, str]]:
    return [(t.kind.value, t.text) for t in tokenize(src) if t.kind is not TokenKind.WHITESPACE]


SAMPLES = [
    "",
    "x",
    "Sub A()\r\n    x = 1\r\nEnd Sub\r\n",
    "x = 1 + _\r\n    2\n",
    "' comment _\r\ncontinued\r\nx = 1",
    's = "a ""quoted"" string"',
    's = "unterminated',
    "d = #1/1/2000#: t = #10:30 PM#",
    "Print #1, x; y",
    "x& = 1: y! = 2: z# = 3: w@ = 4: s$ = \"\"",
    "rs!Field = rs![Other Field]",
    "#If VBA7 Then\n#Else\n#End If\n",
    "10 x = 1\r\n20: y = 2",
    "Rem a remark _\n  still the remark\nx = 1",
    "x = &HFF& + &O17 + &17 + 1.5E+3! + .5 + 1.",
    "If a => b Then c = 1",
    "\t x 　= 1",
    "x = 1 _   \r\n  + 2",
    "\r\r\n\n",
]


@pytest.mark.parametrize("src", SAMPLES)
def test_round_trip(src: str) -> None:
    tokens = tokenize(src)
    assert "".join(t.text for t in tokens) == src
    position = 0
    for token in tokens:
        assert token.start == position
        assert token.end == position + len(token.text)
        position = token.end


def test_comment_continues_across_line_continuation() -> None:
    # Measured: a statement on the line after `' note _` does not run.
    tokens = tokenize("' note _\r\n  x = 2\r\ny = 3")
    assert tokens[0].kind is TokenKind.COMMENT
    assert tokens[0].text == "' note _\r\n  x = 2"


def test_rem_comment_continues_too() -> None:
    tokens = tokenize("Rem note _\n  x = 2\ny")
    assert tokens[0].kind is TokenKind.COMMENT
    assert tokens[0].text == "Rem note _\n  x = 2"


def test_comment_without_space_before_underscore_does_not_continue() -> None:
    tokens = tokenize("' see my_var_\nx = 1")
    assert tokens[0].text == "' see my_var_"


def test_continuation_with_trailing_whitespace() -> None:
    # Measured: `1 + _   ` followed by `2` evaluates to 3.
    tokens = tokenize("x = 1 + _   \r\n  2")
    continuation = [t for t in tokens if t.kind is TokenKind.CONTINUATION]
    assert [t.text for t in continuation] == ["_   \r\n"]


def test_underscore_needs_whitespace_before_it() -> None:
    tokens = tokenize("x = y +_\r\n2")
    assert not any(t.kind is TokenKind.CONTINUATION for t in tokens)


def test_rem_only_at_statement_start() -> None:
    assert kinds("x = 1: Rem c")[-1] == ("comment", "Rem c")
    assert kinds("10 Rem c")[-1] == ("comment", "Rem c")
    assert kinds("Remark = 1")[0] == ("identifier", "Remark")


def test_directive_only_at_line_start() -> None:
    assert kinds("#If x Then")[0] == ("directive", "#")
    assert ("directive", "#") not in kinds("Print #1, x")
    assert ("operator", "#") in kinds("Print #1, x")


def test_date_literals() -> None:
    assert ("date", "#1/1/2000#") in kinds("d = #1/1/2000#")
    assert ("date", "#January 15, 2020#") in kinds("d = #January 15, 2020#")
    # Not a date body: a file number marker.
    assert ("date", '#1, "#') not in kinds('Print #1, "#"')


def test_type_suffixes() -> None:
    assert kinds("Left$(s, 1)")[:2] == [("identifier", "Left"), ("type-suffix", "$")]
    # `&`, `!`, `#` and `^` double as operators, so they stay operators.
    assert kinds("x& = 1")[:2] == [("identifier", "x"), ("operator", "&")]


def test_numbers() -> None:
    assert kinds("&HFF&") == [("integer", "&HFF&")]
    assert kinds("&17") == [("integer", "&17")]
    assert kinds("1.5E+3!") == [("float", "1.5E+3!")]
    assert kinds("1.") == [("float", "1.")]
    assert kinds("1#") == [("float", "1#")]
    assert kinds("10^") == [("integer", "10^")]
    # `&H` without a digit is an operator and a name.
    assert kinds("a &Hx")[1:] == [("operator", "&"), ("identifier", "Hx")]


def test_reversed_comparisons_are_one_operator() -> None:
    assert ("operator", "=>") in kinds("If a => b Then")


def test_keywords_and_contextual_words() -> None:
    assert kinds("Dim x")[0] == ("keyword", "Dim")
    # Contextual keywords are names to the lexer.
    assert kinds("Option Explicit")[1] == ("identifier", "Explicit")


def test_physical_lines_are_counted() -> None:
    tokens = tokenize("a\r\nb _\r\n c\n' x _\n y\nd")
    by_text = {t.text: t.line for t in tokens if t.kind is TokenKind.IDENTIFIER}
    assert by_text == {"a": 0, "b": 1, "c": 2, "d": 5}


def test_bracketed_names() -> None:
    assert kinds("[A1] = 1")[0] == ("bracketed", "[A1]")
    assert kinds("x = [unterminated")[-1] == ("bracketed", "[unterminated")


def test_unterminated_string_gives_back_a_continuation() -> None:
    # MS-VBAL 3.3.4: a string can end at a line continuation.
    tokens = tokenize('s = "abc _\r\n x')
    assert [t.kind for t in tokens if t.kind is TokenKind.CONTINUATION]
    assert any(t.kind is TokenKind.STRING and t.text == '"abc' for t in tokens)
