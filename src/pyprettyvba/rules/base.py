"""What a rule is: a named, configurable check that can also fix what it finds.

A rule reads a ``Document`` and yields ``Finding``s. A finding with a
replacement is a fix: the engine applies it unless a directive suppresses
the rule there. A finding without one is only reported. Rules never edit
text themselves, so the engine can map every finding back to the line the
user wrote, filter by suppression, and check that the fixes left the code
meaning what it meant.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass, field
from typing import Any, ClassVar

from ..document import Document

__all__ = ["Finding", "FormatContext", "Option", "Origin", "Rule"]


@dataclass(frozen=True)
class Option:
    """One setting a rule accepts in its configuration table."""

    name: str
    default: Any
    doc: str
    # Allowed values, for an option that is a choice.
    choices: tuple[Any, ...] | None = None
    # Accepted Python types; None accepts the default's type.
    types: tuple[type, ...] | None = None
    minimum: int | None = None
    maximum: int | None = None

    def validate(self, value: Any) -> str | None:
        """An error message for a bad value, None when it is acceptable."""
        if self.choices is not None:
            if value not in self.choices:
                allowed = ", ".join(repr(c) for c in self.choices)
                return f"must be one of {allowed}, not {value!r}"
            return None
        types = self.types or (type(self.default),)
        if isinstance(value, bool) and bool not in types:
            return f"must be {_type_names(types)}, not a boolean"
        if not isinstance(value, types):
            return f"must be {_type_names(types)}, not {type(value).__name__}"
        if isinstance(value, int) and not isinstance(value, bool):
            if self.minimum is not None and value < self.minimum:
                return f"must be at least {self.minimum}"
            if self.maximum is not None and value > self.maximum:
                return f"must be at most {self.maximum}"
        return None


def _type_names(types: tuple[type, ...]) -> str:
    names = {bool: "a boolean", int: "an integer", str: "a string", list: "a list"}
    return " or ".join(names.get(t, t.__name__) for t in types)


@dataclass(frozen=True)
class Finding:
    """Something a rule found at [start, end) of the text it was given.

    ``replacement`` is the text to put there; None makes the finding a
    report the rule cannot fix. Findings that share a ``group`` are all
    applied but reported once, for a change that touches many places for
    one reason (converting every line ending of a file).
    """

    start: int
    end: int
    replacement: str | None
    message: str
    group: str | None = None


class Origin:
    """Where text a rule sees came from in the module as the user wrote it."""

    def __init__(self, text: str, back: Callable[[int], int] | None = None) -> None:
        self.text = text
        self._back = back

    def offset(self, current: int) -> int:
        """The original offset of ``current`` (an offset in the current text)."""
        return current if self._back is None else self._back(current)

    def column(self, current: int, tab_width: int) -> int:
        """The display column the character at ``current`` had originally."""
        original = self.offset(current)
        start = max(self.text.rfind("\n", 0, original), self.text.rfind("\r", 0, original)) + 1
        col = 0
        for ch in self.text[start:original]:
            col += tab_width - (col % tab_width) if ch == "\t" else 1
        return col


@dataclass
class FormatContext:
    """Settings shared by every rule, plus what is known about the file."""

    indent_width: int = 4
    indent_style: str = "space"  # "space" or "tab"
    tab_width: int = 4
    line_ending: str = "auto"
    filename: str | None = None
    # standard, class, form or document; None when unknown.
    module_kind: str | None = None
    # Host libraries whose global names identifier-case knows.
    hosts: tuple[str, ...] = ("excel",)
    # Names other modules of the project let this one use (their Public
    # declarations, class and form names), lower case -> spelling.
    project_names: Mapping[str, str] = field(default_factory=dict)
    # Every name any module of the project declares, public or not: the
    # VBE's project-wide name table, which spells a name the same way in
    # every module.
    project_declared: Mapping[str, str] = field(default_factory=dict)
    # Names the project's modules let code reach after a `.`.
    project_members: Mapping[str, str] = field(default_factory=dict)
    # Every rule code, for reading suppression directives.
    known_codes: frozenset[str] = frozenset()
    # The resolved options of every enabled rule, for rules that depend on
    # another rule's settings (blank-line indentation follows indent's).
    rule_settings: Mapping[str, Mapping[str, Any]] = field(default_factory=dict)
    # Set by the engine before each rule runs.
    origin: Origin = field(default_factory=lambda: Origin(""))
    # The line terminator new lines are written with.
    newline: str = "\r\n"

    def indent_text(self, columns: int) -> str:
        """Whitespace that indents to ``columns``."""
        if columns <= 0:
            return ""
        if self.indent_style == "tab":
            return "\t" * (columns // self.tab_width) + " " * (columns % self.tab_width)
        return " " * columns

    def width(self, whitespace: str) -> int:
        """Columns ``whitespace`` spans, tabs advancing to the next tab stop."""
        col = 0
        for ch in whitespace:
            if ch == "\t":
                col += self.tab_width - (col % self.tab_width)
            else:
                col += 1
        return col


class Rule:
    """Base class. Subclasses set the class attributes and implement ``run``."""

    code: ClassVar[str]
    summary: ClassVar[str]
    category: ClassVar[str]
    # Enabled when the configuration does not mention the rule.
    default_enabled: ClassVar[bool] = True
    fixable: ClassVar[bool] = True
    # True when the rule only does what the VBE itself does to the code on
    # import, so its output is the text the VBE would export.
    vbe_canonical: ClassVar[bool] = False
    options: ClassVar[tuple[Option, ...]] = ()

    def __init__(self, settings: Mapping[str, Any], context: FormatContext) -> None:
        values = {option.name: option.default for option in self.options}
        values.update(settings)
        self.settings: dict[str, Any] = values
        self.context = context

    def run(self, doc: Document) -> Iterable[Finding]:
        raise NotImplementedError

    @classmethod
    def option(cls, name: str) -> Option:
        for option in cls.options:
            if option.name == name:
                return option
        raise KeyError(name)
