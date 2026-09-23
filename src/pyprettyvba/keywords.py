"""The VBA keyword table: which words are reserved, and how the VBE spells them.

Reserved identifiers are the closed set of MS-VBAL v20250520 section 3.3.5.2,
spelled as the spec spells them, except the literal identifiers, which the
spec writes in lower case and the VBE renders capitalized (True, False,
Nothing, Empty, Null). The spellings were checked against the VBE itself:
tests/oracle/vbe_rendering.json holds every probe the VBE rendered.

Reserved-for-implementation-use names (Attribute, VB_Name, ...) are reserved
but never recased: they only occur on Attribute lines, which the formatter
leaves exactly as written.

Contextual keywords are ordinary identifiers that the VBE capitalizes only in
the statement that gives them meaning (`Explicit` in `Option Explicit`, `Lib`
in a Declare). The lexer reads them as identifiers; the keyword-case rule
cases them where their statement makes them keywords.
"""

from __future__ import annotations

STATEMENT_KEYWORDS: tuple[str, ...] = (
    "Call", "Case", "Close", "Const", "Declare", "DefBool", "DefByte", "DefCur",
    "DefDate", "DefDbl", "DefInt", "DefLng", "DefLngLng", "DefLngPtr", "DefObj",
    "DefSng", "DefStr", "DefVar", "Dim", "Do", "Else", "ElseIf", "End", "EndIf",
    "Enum", "Erase", "Event", "Exit", "For", "Friend", "Function", "Get",
    "Global", "GoSub", "GoTo", "If", "Implements", "Input", "Let", "Lock",
    "Loop", "LSet", "Next", "On", "Open", "Option", "Print", "Private", "Public",
    "Put", "RaiseEvent", "ReDim", "Resume", "Return", "RSet", "Seek", "Select",
    "Set", "Static", "Stop", "Sub", "Type", "Unlock", "Wend", "While", "With",
    "Write",
)

REM_KEYWORD = "Rem"

MARKER_KEYWORDS: tuple[str, ...] = (
    "Any", "As", "ByRef", "ByVal", "Case", "Each", "Else", "In", "New", "Shared",
    "Until", "WithEvents", "Write", "Optional", "ParamArray", "Preserve", "Spc",
    "Tab", "Then", "To",
)

OPERATOR_IDENTIFIERS: tuple[str, ...] = (
    "AddressOf", "And", "Eqv", "Imp", "Is", "Like", "New", "Mod", "Not", "Or",
    "TypeOf", "Xor",
)

RESERVED_NAMES: tuple[str, ...] = (
    "Abs", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng",
    "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "CVErr", "Date", "Debug",
    "DoEvents", "Fix", "Int", "Len", "LenB", "Me", "PSet", "Scale", "Sgn",
    "String",
)

SPECIAL_FORMS: tuple[str, ...] = (
    "Array", "Circle", "Input", "InputB", "LBound", "Scale", "UBound",
)

RESERVED_TYPE_IDENTIFIERS: tuple[str, ...] = (
    "Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long",
    "LongLong", "LongPtr", "Single", "String", "Variant",
)

LITERAL_IDENTIFIERS: tuple[str, ...] = ("True", "False", "Nothing", "Empty", "Null")

FUTURE_RESERVED: tuple[str, ...] = ("CDecl", "Decimal", "DefDec")

RESERVED_FOR_IMPLEMENTATION_USE: tuple[str, ...] = (
    "Attribute", "LINEINPUT", "VB_Base", "VB_Control", "VB_Creatable",
    "VB_Customizable", "VB_Description", "VB_Exposed", "VB_Ext_KEY",
    "VB_GlobalNameSpace", "VB_HelpID", "VB_Invoke_Func", "VB_Invoke_Property",
    "VB_Invoke_PropertyPut", "VB_Invoke_PropertyPutRef", "VB_MemberFlags",
    "VB_Name", "VB_PredeclaredId", "VB_ProcData", "VB_TemplateDerived",
    "VB_UserMemId", "VB_VarDescription", "VB_VarHelpID", "VB_VarMemberFlags",
    "VB_VarProcData", "VB_VarUserMemId",
)

# Contextual keywords, keyed by lower case. The keyword-case rule decides from
# the statement whether a given occurrence is the keyword or a name.
CONTEXTUAL_KEYWORDS: dict[str, str] = {
    word.lower(): word
    for word in (
        # Option Explicit / Base / Compare {Binary | Text | Database} / Private Module
        "Explicit", "Base", "Compare", "Binary", "Text", "Database", "Module",
        # Declare [PtrSafe] ... Lib "..." [Alias "..."]
        "PtrSafe", "Lib", "Alias",
        # Property Get / Let / Set, End Property, Exit Property
        "Property",
        # For ... Step
        "Step",
        # On Error, Error n, Resume Next
        "Error",
        # Open ... For {Append | Binary | Input | Output | Random}
        #   [Access {Read | Write | Read Write}] [Shared | Lock Read ...]
        "Append", "Output", "Random", "Access", "Read",
        # Line Input #
        "Line",
        # Width #n, width
        "Width",
        # Name old As new
        "Name",
        # As Object
        "Object",
    )
}

_CASED_LISTS: tuple[tuple[str, ...], ...] = (
    STATEMENT_KEYWORDS,
    (REM_KEYWORD,),
    MARKER_KEYWORDS,
    OPERATOR_IDENTIFIERS,
    RESERVED_NAMES,
    SPECIAL_FORMS,
    RESERVED_TYPE_IDENTIFIERS,
    LITERAL_IDENTIFIERS,
    FUTURE_RESERVED,
)


def _build_canonical() -> dict[str, str]:
    table: dict[str, str] = {}
    for words in _CASED_LISTS:
        for word in words:
            table.setdefault(word.lower(), word)
    return table


# Lower-case reserved word -> the VBE's spelling (reserved-for-implementation
# names excluded, see the module docstring).
CANONICAL: dict[str, str] = _build_canonical()

RESERVED: frozenset[str] = frozenset(CANONICAL) | frozenset(
    word.lower() for word in RESERVED_FOR_IMPLEMENTATION_USE
)

# Words that take an operand on both sides.
BINARY_OPERATOR_WORDS: frozenset[str] = frozenset(
    ("and", "or", "xor", "eqv", "imp", "mod", "like", "is")
)

# Words that take one operand on their right.
UNARY_OPERATOR_WORDS: frozenset[str] = frozenset(("not", "addressof", "typeof", "new"))

# Reserved words that can end an operand, so an operator after one is binary.
# Everything else in RESERVED is structural (statement and marker keywords,
# operator words), and an operator after it starts a new operand.
OPERAND_WORDS: frozenset[str] = frozenset(
    word.lower()
    for word in (
        *RESERVED_NAMES,
        *SPECIAL_FORMS,
        *RESERVED_TYPE_IDENTIFIERS,
        *LITERAL_IDENTIFIERS,
        "Decimal",
    )
)

# Words that name a function when a parenthesis follows them, so `Len(x)`
# keeps the parenthesis glued the way the VBE writes it.
FUNCTION_WORDS: frozenset[str] = frozenset(
    word.lower()
    for word in (
        "Abs", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng",
        "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "CVErr", "Fix", "Int",
        "Len", "LenB", "Sgn", "Array", "Input", "InputB", "LBound", "UBound",
        "Date", "String", "Spc", "Tab", "DoEvents",
    )
)

DEFTYPE_KEYWORDS: frozenset[str] = frozenset(
    word.lower() for word in STATEMENT_KEYWORDS if word.startswith("Def")
)


def canonical_keyword(word: str) -> str | None:
    """The VBE spelling of a reserved word, or None for anything else."""
    return CANONICAL.get(word.lower())


def is_reserved(word: str) -> bool:
    """True when ``word`` can never be a user-defined name (MS-VBAL 3.3.5.2)."""
    return word.lower() in RESERVED
