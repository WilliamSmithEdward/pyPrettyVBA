"""Literal values and the VBE's spelling of them."""

from __future__ import annotations

from decimal import Decimal

import pytest

from pyprettyvba.literals import (
    NumberValue,
    canonical_date,
    canonical_number,
    number_value,
    parse_date_body,
    precision_loss,
)

# (literal, the VBE's spelling) - every pair measured, see tests/oracle.
VBE_NUMBERS = [
    ("1.0", "1#"), ("1.50", "1.5"), (".5", "0.5"), ("1e3", "1000#"), ("1.5e+3", "1500#"),
    ("1.5E-3", "0.0015"), ("1d2", "100#"), ("1D+2", "100#"), ("100000000000000000000", "1E+20"),
    ("123456789012345", "123456789012345#"), ("1E15", "1E+15"), ("1E16", "1E+16"),
    ("3.14159265358979", "3.14159265358979"), ("0.1", "0.1"), ("100.0", "100#"), ("0.0", "0#"),
    ("0.000001", "0.000001"), ("0.0000001", "0.0000001"), ("1E-20", "1E-20"),
    ("1.23456789012345E+300", "1.23456789012345E+300"), ("1.", "1#"), ("5E+0", "5#"),
    ("0.1#", "0.1"), ("1#", "1#"), ("5!", "5!"), ("1.0!", "1!"), ("1.5!", "1.5!"),
    ("1.5E3!", "1500!"), ("2.5@", "2.5@"), ("0.5@", "0.5@"),
    ("922337203685477.5807@", "922337203685477.5807@"), ("&hff", "&HFF"), ("&HFF&", "&HFF&"),
    ("&h00ff", "&HFF"), ("&h0", "&H0"), ("&hFFFF", "&HFFFF"), ("&hFFFFFFFF", "&HFFFFFFFF"),
    ("&H8000", "&H8000"), ("&H8000&", "&H8000&"), ("&o17", "&O17"), ("&O17&", "&O17&"),
    ("&17", "&O17"), ("32768", "32768"), ("2147483648", "2147483648#"), ("10&", "10&"),
    ("10%", "10"), ("00012", "12"), ("32767%", "32767"),
    ("1E-8", "0.00000001"), ("1E-15", "0.000000000000001"), ("1E-16", "1E-16"),
    ("1.5E-10", "0.00000000015"), ("1.23E-5", "0.0000123"),
    ("0.000123456789012345", "1.23456789012345E-04"), ("1.2345E-8", "0.000000012345"),
    ("0.00001!", "0.00001!"), ("1E-10!", "1E-10!"), ("32768&", "32768"), ("&HFF%", "&HFF"),
    ("&HFFFF&", "&HFFFF&"), ("70000&", "70000"), ("10^", "10^"),
]


@pytest.mark.parametrize("literal, vbe", VBE_NUMBERS)
def test_numbers_are_spelled_as_the_vbe_spells_them(literal: str, vbe: str) -> None:
    assert canonical_number(literal) == vbe
    assert number_value(vbe) == number_value(literal)


# Literals the VBE rewrites to a different value: never rewritten, reported.
LOSSY = [
    ("123456789012345678", "1.23456789012346E+17"),
    ("1234567890123456", "1.23456789012346E+15"),
    ("3.141592653589793", "3.14159265358979"),
    ("12345678901234567890#", "1.23456789012346E+19"),
    ("16777217!", "1.677722E+07!"),
]


@pytest.mark.parametrize("literal, vbe", LOSSY)
def test_value_changing_spellings_are_refused_and_reported(literal: str, vbe: str) -> None:
    assert canonical_number(literal) is None
    assert precision_loss(literal) == vbe


def test_values_follow_the_spec_tables() -> None:
    assert number_value("&HFFFF") == NumberValue("Integer", -1)
    assert number_value("&HFFFF&") == NumberValue("Long", 65535)
    assert number_value("&HFFFFFFFF") == NumberValue("Long", -1)
    assert number_value("32767") == NumberValue("Integer", 32767)
    assert number_value("32768") == NumberValue("Long", 32768)
    assert number_value("2147483648") == NumberValue("Double", 2147483648.0)
    assert number_value("1.5@") == NumberValue("Currency", Decimal("1.5000"))
    assert number_value("32768%") is None  # an Integer cannot hold it
    assert number_value("1E400") is None  # beyond a Double


def test_single_rounding_is_exact() -> None:
    # 16777217 has no Single: it rounds to 16777216, ties to even.
    assert number_value("16777217!") == NumberValue("Single", 16777216.0)


VBE_DATES = [
    ("#2020-01-15#", "#1/15/2020#"), ("#1/15/2020#", "#1/15/2020#"), ("#15/1/2020#", "#1/15/2020#"),
    ("#1/1/2000 10:30#", "#1/1/2000 10:30:00 AM#"), ("#10:30:00#", "#10:30:00 AM#"),
    ("#10:30 PM#", "#10:30:00 PM#"), ("#January 15, 2020#", "#1/15/2020#"),
    ("#2020/1/15#", "#1/15/2020#"), ("#12:00:00 AM#", "#12:00:00 AM#"), ("#1-Jan-2020#", "#1/1/2020#"),
    ("#  1/1/2000  #", "#1/1/2000#"), ("#00:00#", "#12:00:00 AM#"),
    ("#1/1/2000 00:00:00#", "#1/1/2000#"), ("#12/31/2020 23:59:59#", "#12/31/2020 11:59:59 PM#"),
]


@pytest.mark.parametrize("literal, vbe", VBE_DATES)
def test_dates_are_spelled_as_the_vbe_spells_them(literal: str, vbe: str) -> None:
    assert canonical_date(literal) == vbe


@pytest.mark.parametrize("literal", ["#1/1/99#", "#1/15#", "#Jan 2020#", "#13/13/2020#", "##"])
def test_machine_dependent_or_invalid_dates_are_left_alone(literal: str) -> None:
    # A two-digit year or a missing year depends on the machine's settings.
    assert canonical_date(literal) is None


def test_date_body_grammar() -> None:
    assert parse_date_body("1/1/2000") is not None
    assert parse_date_body("10:30 PM") is not None
    assert parse_date_body("1, ") is None
    assert parse_date_body("not a date") is None
