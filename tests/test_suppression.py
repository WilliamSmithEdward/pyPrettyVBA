"""Suppression directives."""

from __future__ import annotations

from pyprettyvba import Config, format_source

MESSY = "Sub A()\r\nx=1\r\nEnd Sub\r\n"


def fmt(src: str, **config: object) -> str:
    return format_source(src, Config(dict(config))).output


def rules_hit(src: str) -> set[str]:
    return {v.rule for v in format_source(src, Config()).violations}


def test_line_directive_suppresses_its_own_line() -> None:
    src = "Sub A()\r\nx=1 '@prettyvba-ignore\r\ny=2\r\nEnd Sub\r\n"
    assert fmt(src) == "Sub A()\r\nx=1 '@prettyvba-ignore\r\n    y = 2\r\nEnd Sub\r\n"


def test_directive_with_rule_list_suppresses_only_those_rules() -> None:
    src = "Sub A()\r\nx=1 '@prettyvba-ignore: spacing\r\nEnd Sub\r\n"
    assert fmt(src) == "Sub A()\r\n    x=1 '@prettyvba-ignore: spacing\r\nEnd Sub\r\n"


def test_next_line_skips_blank_lines_and_stacked_directives() -> None:
    src = (
        "Sub A()\r\n"
        "    '@prettyvba-ignore-next-line: spacing\r\n"
        "    '@prettyvba-ignore-next-line: indent\r\n"
        "\r\n"
        "  x=1\r\n"
        "End Sub\r\n"
    )
    assert "  x=1\r\n" in fmt(src)


def test_file_directive() -> None:
    src = "'@prettyvba-ignore-file: indent\r\nSub A()\r\nx=1\r\nEnd Sub\r\n"
    assert fmt(src) == "'@prettyvba-ignore-file: indent\r\nSub A()\r\nx = 1\r\nEnd Sub\r\n"


def test_region() -> None:
    src = (
        "Sub A()\r\n"
        "    '@prettyvba-ignore-start\r\n"
        "  x=1\r\n"
        "  y=2\r\n"
        "    '@prettyvba-ignore-end\r\n"
        "w=3\r\n"
        "End Sub\r\n"
    )
    out = fmt(src)
    assert "  x=1\r\n  y=2\r\n" in out
    assert "    w = 3\r\n" in out


def test_continuation_lines_belong_to_their_line() -> None:
    src = "Sub A()\r\n    '@prettyvba-ignore-next-line\r\n    x = Array(1,  0, _\r\n              0,  1)\r\nEnd Sub\r\n"
    assert fmt(src) == src


def test_bad_directives_are_reported_and_suppress_nothing() -> None:
    for text, message in [
        ("x = 1 '@prettyvba-ignore: nonsense\r\n", "Unknown rule 'nonsense'"),
        ("x = 1 '@prettyvba-ignored\r\n", None),
        ("x = 1 '@prettyvba-ignore-sometimes\r\n", "Unknown suppression directive"),
        ("'@prettyvba-ignore-end\r\n", "no matching -start"),
        ("'@prettyvba-ignore-start\r\nx = 1\r\n", "never closed"),
        ("x = 1\r\n'@prettyvba-ignore-file\r\n", "before the first line of code"),
        ("'@prettyvba-ignore: all, indent\r\n", "'all' cannot be combined"),
    ]:
        violations = [v for v in format_source(text, Config()).violations if v.rule == "suppression-directive"]
        if message is None:
            assert violations == [], text
        else:
            assert violations and message in violations[0].message, text


def test_rem_and_doc_comments_are_never_directives() -> None:
    src = "Sub A()\r\nx=1 ''' @prettyvba-ignore\r\nEnd Sub\r\n"
    assert "x = 1" in fmt(src)


def test_directive_rule_itself_cannot_be_suppressed() -> None:
    src = "'@prettyvba-ignore-file\r\nx = 1 '@prettyvba-ignore: bogus\r\n"
    assert "suppression-directive" in rules_hit(src)
