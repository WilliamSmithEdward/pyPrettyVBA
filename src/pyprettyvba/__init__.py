"""pyPrettyVBA: a configurable formatter for VBA, with autofix and rule suppression.

Formatting is a pipeline of named rules. Each rule can be enabled, disabled
and configured, and suppressed for a line, a region or a module with a
`'@prettyvba-ignore` comment. The default rules reproduce what the VBE does
to code when it reads it in (keyword case, spacing, literal spelling), then
add the layout the VBE leaves alone (indentation, blank lines). Output is
checked before it is returned: formatting never changes what the code does.

    from pyprettyvba import format_source

    result = format_source("sub greet()\\nmsgbox \\"hi\\"\\nend sub\\n")
    result.output       # 'Sub greet()\\n    MsgBox "hi"\\nEnd Sub\\n'
    result.violations   # what changed, by rule, with line and column
"""

from __future__ import annotations

from .api import (
    FileResult,
    ProjectNames,
    check_source,
    collect_files,
    format_file,
    format_office_file,
    format_paths,
    format_source,
    module_kind,
    project_names,
)
from .config import PRESETS, Config, ConfigError, Settings, find_config_file
from .engine import FormatResult, UnstableFormattingError, Violation
from .rules import RULES, RULES_BY_CODE, Rule
from .safety import SafetyError

__version__ = "0.1.0"

__all__ = [
    "Config",
    "ConfigError",
    "FileResult",
    "FormatResult",
    "PRESETS",
    "ProjectNames",
    "RULES",
    "RULES_BY_CODE",
    "Rule",
    "SafetyError",
    "Settings",
    "UnstableFormattingError",
    "Violation",
    "__version__",
    "check_source",
    "collect_files",
    "find_config_file",
    "format_file",
    "format_office_file",
    "format_paths",
    "format_source",
    "module_kind",
    "project_names",
]
