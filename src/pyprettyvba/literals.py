"""Numeric and date literals: their exact values, and how the VBE writes them.

Values follow MS-VBAL v20250520 sections 3.3.2 (numbers) and 3.3.3 (dates).
The VBE's own spelling of a literal was measured by pasting literals into a
module and exporting it (tests/oracle/vbe_rendering.json). What it does:

* An integer loses leading zeros and any type suffix its value makes
  redundant: `00012` is written `12`, `10%` is written `10`, but `10&` keeps
  its suffix because `10` alone would be an Integer.
* Hex and octal literals are written in upper case without leading zeros:
  `&h00ff` becomes `&HFF`, and the radix-less octal `&17` becomes `&O17`.
* A Double is written with at most 15 significant digits. An integral value
  gets a `#` so it stays a Double (`1.0`, `1e3` and `1.` become `1#`,
  `1000#` and `1#`), a fraction drops trailing zeros and gains a leading zero
  (`1.50` is `1.5`, `.5` is `0.5`), and a value that would need more than
  15 integer digits or more than 15 digits after the point is written in E
  notation with a signed two-digit exponent (`1E+15`, `1E-16`).
* A Single keeps its `!` and 7 significant digits, with the same shape at 7;
  a Currency keeps its `@`.
* Integer suffixes are dropped where the value makes them redundant, for hex
  and octal literals too: `32768&` is `32768`, `&HFF%` is `&HFF`.

Fifteen significant digits is fewer than a Double can hold, so the VBE
rewrites `3.141592653589793` as `3.14159265358979`, a different number. This
module never does that: a canonical spelling is offered only when it denotes
exactly the same value as the original, and `precision_loss` names the value
the VBE would substitute so the caller can warn about it.

Dates are written `#M/D/YYYY#`, `#h:mm:ss AM#`, or both, with a midnight
time dropped from a date. A literal whose value depends on the machine (a
two-digit year, or no year at all) is never rewritten.
"""

from __future__ import annotations

import datetime
import math
import re
import struct
from collections.abc import Iterator
from dataclasses import dataclass
from decimal import ROUND_HALF_EVEN, Decimal, DecimalException
from fractions import Fraction

from .chars import WSC_CHARS

__all__ = [
    "DateValue",
    "NumberValue",
    "canonical_date",
    "canonical_number",
    "number_value",
    "parse_date_body",
    "precision_loss",
    "resolve_date",
]


@dataclass(frozen=True)
class NumberValue:
    """A numeric literal's declared type and exact data value."""

    type: str  # Integer, Long, LongLong, Single, Double or Currency
    value: int | float | Decimal


# --- integers ------------------------------------------------------------

_INT_MAX = 32767
_LONG_MAX = 2147483647
_LONGLONG_MAX = 9223372036854775807


def _integer_value(text: str) -> NumberValue | None:
    """Declared type and signed value of an INTEGER token, per the spec table."""
    suffix = text[-1] if text[-1] in "%&^" and len(text) > 1 else ""
    body = text[: len(text) - len(suffix)] if suffix else text
    if body.startswith("&"):
        radix_char = body[1:2].lower()
        if radix_char == "h":
            digits, radix = body[2:], 16
        elif radix_char == "o":
            digits, radix = body[2:], 8
        else:
            digits, radix = body[1:], 8
        if not digits:
            return None
        n = int(digits, radix)
        return _radix_value(n, suffix)
    if not body.isdigit():
        return None
    significant = body.lstrip("0") or "0"
    if len(significant) > 310:
        # Past any Double; Python also refuses to read very long numbers.
        return None
    n = int(significant)
    if suffix == "%":
        return NumberValue("Integer", n) if n <= _INT_MAX else None
    if suffix == "&":
        return NumberValue("Long", n) if n <= _LONG_MAX else None
    if suffix == "^":
        return NumberValue("LongLong", n) if n <= _LONGLONG_MAX else None
    if n <= _INT_MAX:
        return NumberValue("Integer", n)
    if n <= _LONG_MAX:
        return NumberValue("Long", n)
    # Too big for a Long: the literal is a Double (MS-VBAL 3.3.2 note 1).
    return NumberValue("Double", float(n)) if n < 2**1024 else None


def _radix_value(n: int, suffix: str) -> NumberValue | None:
    """Hex and octal literals: the value wraps into the signed range."""
    if suffix == "^":
        if n <= 0xFFFFFFFFFFFFFFFF:
            return NumberValue("LongLong", n - (1 << 64) if n > _LONGLONG_MAX else n)
        return None
    if n <= 0xFFFF and suffix in ("", "%"):
        return NumberValue("Integer", n - 0x10000 if n > _INT_MAX else n)
    if suffix == "%":
        return None
    if n <= 0xFFFFFFFF:
        # An `&` suffix makes 0..&HFFFF a Long with its unsigned value.
        return NumberValue("Long", n - 0x100000000 if n > _LONG_MAX else n)
    return None


def _default_integer_type(n: int) -> str:
    if n <= _INT_MAX:
        return "Integer"
    if n <= _LONG_MAX:
        return "Long"
    return "Double"


# --- floats --------------------------------------------------------------

_FLOAT_RE = re.compile(
    r"^(?P<int>[0-9]*)(?:\.(?P<frac>[0-9]*))?(?:[DdEe](?P<exp>[+-]?[0-9]+))?(?P<suffix>[!#@]?)$"
)

_SINGLE_MAX = Fraction((2**24 - 1) * 2**104)
_CURRENCY_MAX = Decimal("922337203685477.5807")
# Past this no Currency value is near; quantizing it would overflow Decimal.
_CURRENCY_LIMIT = Decimal("1E16")


def _float_decimal(text: str) -> tuple[Decimal, str] | None:
    """The exact decimal value written by a FLOAT token, and its suffix."""
    m = _FLOAT_RE.match(text)
    if m is None or (not m.group("int") and not m.group("frac")):
        return None
    digits = (m.group("int") or "0") + "." + (m.group("frac") or "0")
    exp = m.group("exp") or "0"
    try:
        value = Decimal(digits).scaleb(int(exp))
    except (DecimalException, ValueError):
        # An exponent too large for any VBA type, or for Python to read.
        return None
    return value, m.group("suffix")


def _float_value(text: str) -> NumberValue | None:
    parsed = _float_decimal(text)
    if parsed is None:
        return None
    exact, suffix = parsed
    if suffix == "@":
        if abs(exact) >= _CURRENCY_LIMIT:
            return None
        rounded = exact.quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
        if abs(rounded) > _CURRENCY_MAX:
            return None
        return NumberValue("Currency", rounded)
    if suffix == "!":
        single = _round_to_single(Fraction(exact))
        return None if single is None else NumberValue("Single", single)
    try:
        value = float(exact)
    except OverflowError:
        return None
    if math.isinf(value):
        return None
    return NumberValue("Double", value)


def _round_to_single(x: Fraction) -> float | None:
    """Round an exact value to the nearest IEEE single, ties to even."""
    if x == 0:
        return 0.0
    sign = -1 if x < 0 else 1
    x = abs(x)
    # Find e with 2**e <= x < 2**(e+1).
    e = x.numerator.bit_length() - x.denominator.bit_length()
    if Fraction(2) ** e > x:
        e -= 1
    # Normal singles carry 24 significant bits; below 2**-126 the spacing is
    # fixed at 2**-149.
    quantum_exp = max(e - 23, -149)
    scaled = x / Fraction(2) ** quantum_exp
    q = scaled.numerator // scaled.denominator
    remainder = scaled - q
    if remainder > Fraction(1, 2) or (remainder == Fraction(1, 2) and q % 2 == 1):
        q += 1
    result = Fraction(q) * Fraction(2) ** quantum_exp
    if result > _SINGLE_MAX:
        return None
    return sign * float(result)


def number_value(text: str) -> NumberValue | None:
    """The declared type and exact value of a numeric literal, or None.

    None means the text is not a valid VBA numeric literal (out of range, or
    malformed); such a literal is never rewritten.
    """
    if not text:
        return None
    if text.startswith("&") or (text[-1] not in "!#@" and _is_integer_text(text)):
        return _integer_value(text)
    return _float_value(text)


def _is_integer_text(text: str) -> bool:
    body = text[:-1] if text[-1] in "%&^" else text
    return body.isdigit()


# --- VBE spelling --------------------------------------------------------

def _digits15(value: float, significant: int) -> tuple[str, int]:
    """Round to ``significant`` digits: (digits without trailing zeros, exponent).

    The exponent is the power of ten of the first digit, so 1500 gives
    ("15", 3) and 0.0015 gives ("15", -3).
    """
    if value == 0:
        return "0", 0
    formatted = f"{value:.{significant - 1}e}"
    mantissa, exp_text = formatted.split("e")
    digits = mantissa.replace(".", "").rstrip("0") or "0"
    return digits, int(exp_text)


def _fixed(digits: str, exp: int) -> str:
    """Digits with first-digit exponent ``exp``, written without an exponent."""
    if exp >= 0:
        whole = digits[: exp + 1].ljust(exp + 1, "0")
        frac = digits[exp + 1 :]
    else:
        whole = "0"
        frac = "0" * (-exp - 1) + digits
    return whole + ("." + frac if frac else "")


def _scientific(digits: str, exp: int) -> str:
    mantissa = digits[0] + ("." + digits[1:] if len(digits) > 1 else "")
    sign = "+" if exp >= 0 else "-"
    return f"{mantissa}E{sign}{abs(exp):02d}"


def _fits_fixed(digits: str, exp: int, precision: int) -> bool:
    """Whether the VBE writes this value without an exponent.

    Measured for Doubles: fixed notation while there are at most 15 integer
    digits and at most 15 digits after the point, so `1E-15` is written out
    and `1E-16` is not, and `0.000123456789012345` (18 decimals) switches to
    `1.23456789012345E-04`. Singles follow the same shape with 7
    (`1.677722E+07!`, `1E-10!`, `0.00001!`).
    """
    if exp >= precision:
        return False
    decimals = max(0, len(digits) - exp - 1)
    return decimals <= precision


def _vbe_double(value: float) -> str:
    digits, exp = _digits15(value, 15)
    if digits == "0":
        return "0#"
    if not _fits_fixed(digits, exp, 15):
        return _scientific(digits, exp)
    text = _fixed(digits, exp)
    return text if "." in text else text + "#"


def _vbe_single(value: float) -> str:
    digits, exp = _digits15(value, 7)
    if digits == "0":
        return "0!"
    if not _fits_fixed(digits, exp, 7):
        return _scientific(digits, exp) + "!"
    return _fixed(digits, exp) + "!"


def _vbe_currency(value: Decimal) -> str:
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    if text.startswith("-"):
        text = text[1:]
    return text + "@"


def canonical_number(text: str) -> str | None:
    """How the VBE writes a numeric literal, when that keeps its exact value.

    Returns None when the literal is invalid, when the VBE's spelling would
    change its value (see ``precision_loss``), or when the VBE's spelling of
    that kind of value has not been measured.
    """
    value = number_value(text)
    if value is None:
        return None
    if text.startswith("&"):
        return _canonical_radix(text)
    if isinstance(value.value, int):
        # An INTEGER token: drop leading zeros and any redundant suffix.
        suffix = text[-1] if text[-1] in "%&^" else ""
        digits = (text[:-1] if suffix else text).lstrip("0") or "0"
        if suffix and value.type == _default_integer_type(int(digits)):
            suffix = ""
        return digits + suffix
    if value.type == "Currency":
        assert isinstance(value.value, Decimal)
        candidate = _vbe_currency(value.value)
    elif value.type == "Single":
        assert isinstance(value.value, float)
        candidate = _vbe_single(value.value)
    else:
        assert isinstance(value.value, float)
        candidate = _vbe_double(value.value)
    if number_value(candidate) != value:
        return None
    return candidate


def _canonical_radix(text: str) -> str:
    """Upper case, no leading zeros, and no suffix the value makes redundant.

    `&h00ff` is `&HFF`, `&HFF%` is `&HFF` (an Integer either way), but
    `&HFFFF&` keeps its `&`: without it the literal is the Integer -1.
    """
    suffix = text[-1] if text[-1] in "%&^" else ""
    body = text[: len(text) - len(suffix)] if suffix else text
    if body[1:2] in ("h", "H"):
        prefix, digits = "&H", body[2:]
    elif body[1:2] in ("o", "O"):
        prefix, digits = "&O", body[2:]
    else:
        prefix, digits = "&O", body[1:]
    digits = digits.lstrip("0").upper() or "0"
    if suffix and number_value(prefix + digits) == number_value(text):
        suffix = ""
    return prefix + digits + suffix


def precision_loss(text: str) -> str | None:
    """The spelling the VBE substitutes when it would change the value.

    `3.141592653589793` has more significant digits than the VBE keeps; it
    stores and exports `3.14159265358979`, a different Double. Returns that
    spelling, or None when the VBE keeps the value.
    """
    value = number_value(text)
    if value is None or text.startswith("&") or value.type == "Currency":
        return None
    if isinstance(value.value, int):
        return None
    assert isinstance(value.value, float)
    spelled = _vbe_single(value.value) if value.type == "Single" else _vbe_double(value.value)
    return None if number_value(spelled) == value else spelled


# --- dates ---------------------------------------------------------------

_MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


@dataclass(frozen=True)
class _DatePart:
    """One of the three fields of a date-value: a number or a month name."""

    number: int | None
    month: int | None
    digits: int  # how many digits the number was written with


@dataclass(frozen=True)
class _TimeValue:
    hour: int
    minute: int
    second: int
    ampm: str  # "", "a" or "p"


@dataclass(frozen=True)
class DateBody:
    """A parsed date literal body: an optional date-value and time-value."""

    parts: tuple[_DatePart, ...]
    time: _TimeValue | None


@dataclass(frozen=True)
class DateValue:
    """A date literal's fully determined value."""

    year: int | None
    month: int | None
    day: int | None
    hour: int
    minute: int
    second: int


class _Cursor:
    def __init__(self, text: str) -> None:
        self.text = text
        self.pos = 0

    def skip_wsc(self) -> int:
        start = self.pos
        while self.pos < len(self.text) and self.text[self.pos] in WSC_CHARS:
            self.pos += 1
        return self.pos - start

    def peek(self) -> str:
        return self.text[self.pos] if self.pos < len(self.text) else ""

    def number(self) -> tuple[int, int] | None:
        start = self.pos
        while self.pos < len(self.text) and "0" <= self.text[self.pos] <= "9":
            self.pos += 1
        if self.pos == start:
            return None
        if self.pos - start > 18:
            # No part of a date has this many digits.
            self.pos = start
            return None
        return int(self.text[start : self.pos]), self.pos - start

    def word(self) -> str:
        start = self.pos
        while self.pos < len(self.text) and self.text[self.pos].isascii() and self.text[self.pos].isalpha():
            self.pos += 1
        return self.text[start : self.pos]

    @property
    def done(self) -> bool:
        return self.pos >= len(self.text)


def parse_date_body(body: str) -> DateBody | None:
    """Parse the text between a `#` pair (MS-VBAL 3.3.3), or return None.

    This is the syntactic grammar only; ``resolve_date`` applies the rules
    that make a date valid. The date grammar is ambiguous, so every reading
    is tried and the first that consumes the whole body wins.
    """
    stripped = body.strip("".join(WSC_CHARS))
    if not stripped:
        return DateBody((), None)
    for reading in _date_readings(body):
        return reading
    return None


def _date_readings(body: str) -> Iterator[DateBody]:
    """Every complete reading of a date literal body, most specific first.

    date-or-time = (date-value 1*WSC time-value) / date-value / time-value
    """
    cursor = _Cursor(body)
    cursor.skip_wsc()
    start = cursor.pos
    for parts, after in _date_values(body, start):
        cursor.pos = after
        spaces = cursor.skip_wsc()
        if cursor.done:
            yield DateBody(parts, None)
            continue
        if spaces:
            time = _time_value(body, cursor.pos)
            if time is not None:
                value, end = time
                cursor.pos = end
                cursor.skip_wsc()
                if cursor.done:
                    yield DateBody(parts, value)
    time = _time_value(body, start)
    if time is not None:
        value, end = time
        cursor.pos = end
        cursor.skip_wsc()
        if cursor.done:
            yield DateBody((), value)


def _date_part(text: str, pos: int) -> tuple[_DatePart, int] | None:
    cursor = _Cursor(text)
    cursor.pos = pos
    number = cursor.number()
    if number is not None:
        return _DatePart(number[0], None, number[1]), cursor.pos
    word = cursor.word().lower()
    if word in _MONTHS:
        return _DatePart(None, _MONTHS[word], 0), cursor.pos
    return None


def _date_separator(text: str, pos: int) -> int | None:
    """date-separator = 1*WSC / (*WSC ("/" / "-" / ",") *WSC)."""
    cursor = _Cursor(text)
    cursor.pos = pos
    spaces = cursor.skip_wsc()
    if cursor.peek() in ("/", "-", ","):
        cursor.pos += 1
        cursor.skip_wsc()
        return cursor.pos
    return cursor.pos if spaces > 0 else None


def _date_values(text: str, pos: int) -> Iterator[tuple[tuple[_DatePart, ...], int]]:
    left = _date_part(text, pos)
    if left is None:
        return
    sep1 = _date_separator(text, left[1])
    if sep1 is None:
        return
    middle = _date_part(text, sep1)
    if middle is None:
        return
    sep2 = _date_separator(text, middle[1])
    if sep2 is not None:
        right = _date_part(text, sep2)
        if right is not None:
            yield (left[0], middle[0], right[0]), right[1]
    yield (left[0], middle[0]), middle[1]


def _time_value(text: str, pos: int) -> tuple[_TimeValue, int] | None:
    """time-value = (hour ampm) / (hour sep minute [sep second] [ampm])."""
    cursor = _Cursor(text)
    cursor.pos = pos
    hour = cursor.number()
    if hour is None:
        return None
    after_hour = cursor.pos
    minute = second = 0
    has_minute = False
    cursor.skip_wsc()
    if cursor.peek() in (":", "."):
        cursor.pos += 1
        cursor.skip_wsc()
        found = cursor.number()
        if found is None:
            return None
        minute = found[0]
        has_minute = True
        mark = cursor.pos
        cursor.skip_wsc()
        if cursor.peek() in (":", "."):
            cursor.pos += 1
            cursor.skip_wsc()
            found = cursor.number()
            if found is None:
                return None
            second = found[0]
        else:
            cursor.pos = mark
    else:
        cursor.pos = after_hour
    mark = cursor.pos
    cursor.skip_wsc()
    word = cursor.word().lower()
    ampm = ""
    if word in ("am", "a", "pm", "p"):
        ampm = word[0]
    else:
        cursor.pos = mark
        if not has_minute:
            return None
    return _TimeValue(hour[0], minute, second, ampm), cursor.pos


class _MachineDependent(Exception):
    """The answer depends on the machine's two-digit-year window."""


def _year(field: _DatePart) -> int:
    """Year(x) from MS-VBAL 3.3.3, refusing the machine-dependent case.

    A year written with fewer than three digits goes through the machine's
    two-digit-year window (1999 on one machine, 2099 on another).
    """
    assert field.number is not None
    if field.digits < 3:
        raise _MachineDependent
    return field.number


def _legal_day(month: int, day: int, year_field: _DatePart) -> bool:
    """LegalDay(month, day, Year(year_field)).

    For a two-digit year the answer can still be certain: it only depends on
    the window for 29 February, so both centuries are tried.
    """
    if not 1 <= month <= 12 or day < 1:
        return False
    assert year_field.number is not None
    if year_field.digits >= 3:
        years = [year_field.number]
    else:
        years = [1900 + year_field.number, 2000 + year_field.number]
    verdicts = set()
    for year in years:
        try:
            datetime.date(year, month, day)
            verdicts.add(True)
        except ValueError:
            verdicts.add(False)
    if len(verdicts) > 1:
        raise _MachineDependent
    return verdicts.pop()


def resolve_date(body: DateBody) -> DateValue | None:
    """The value of a date literal, or None when invalid or machine-dependent.

    The fields are read with the rules of MS-VBAL 3.3.3 in their stated order.
    A literal whose value depends on the machine (a two-digit year, or a date
    with no year, which takes the current year) gets no value here.
    """
    if not body.parts and body.time is None:
        return None
    hour = minute = second = 0
    if body.time is not None:
        t = body.time
        if not (0 <= t.hour <= 23 and 0 <= t.minute <= 59 and 0 <= t.second <= 59):
            return None
        hour = t.hour
        if t.ampm == "p" and hour < 12:
            hour += 12
        elif t.ampm == "a" and hour == 12:
            hour = 0
        minute, second = t.minute, t.second
    if not body.parts:
        return DateValue(None, None, None, hour, minute, second)
    if len(body.parts) == 2:
        # Two fields leave the year to the current year, or make one field a
        # year with the day defaulting to 1 after a current-year check.
        return None
    try:
        found = _resolve_three(body.parts)
    except _MachineDependent:
        return None
    if found is None:
        return None
    year, month, day = found
    return DateValue(year, month, day, hour, minute, second)


def _resolve_three(parts: tuple[_DatePart, ...]) -> tuple[int, int, int] | None:
    left, middle, right = parts
    names = [part for part in parts if part.month is not None]
    if len(names) > 1:
        return None
    if not names:
        assert left.number is not None and middle.number is not None
        L, M = left.number, middle.number
        assert right.number is not None
        R = right.number
        if 1 <= L <= 12 and _legal_day(L, M, right):
            return _year(right), L, M
        if 1 <= M <= 12 and _legal_day(M, R, left):
            return _year(left), M, R
        if 1 <= M <= 12 and _legal_day(M, L, right):
            return _year(right), M, L
        return None
    month = names[0].month
    assert month is not None
    n1, n2 = [part for part in parts if part.number is not None]
    assert n1.number is not None and n2.number is not None
    if _legal_day(month, n1.number, n2):
        return _year(n2), month, n1.number
    if _legal_day(month, n2.number, n1):
        return _year(n1), month, n2.number
    return None


def _date_literal_value(text: str) -> DateValue | None:
    if len(text) < 2 or text[0] != "#" or text[-1] != "#":
        return None
    body = parse_date_body(text[1:-1])
    if body is None:
        return None
    return resolve_date(body)


def canonical_date(text: str) -> str | None:
    """How the VBE writes a date literal, or None when that is not safe.

    `#2020-01-15#` becomes `#1/15/2020#`, `#10:30 PM#` becomes
    `#10:30:00 PM#`, and `#1/1/2000 00:00:00#` becomes `#1/1/2000#`.
    """
    value = _date_literal_value(text)
    if value is None:
        return None
    pieces: list[str] = []
    if value.year is not None:
        pieces.append(f"{value.month}/{value.day}/{value.year}")
    has_time = value.hour or value.minute or value.second
    if has_time or value.year is None:
        hour12 = value.hour % 12 or 12
        ampm = "AM" if value.hour < 12 else "PM"
        pieces.append(f"{hour12}:{value.minute:02d}:{value.second:02d} {ampm}")
    return "#" + " ".join(pieces) + "#"


def date_identity(text: str) -> DateValue | str:
    """A date literal's value when known, otherwise its exact text."""
    value = _date_literal_value(text)
    return text if value is None else value


def single_bits(value: float) -> bytes:
    """The IEEE single encoding of ``value`` (used by tests)."""
    return struct.pack("<f", value)
