# What the VBE does, measured

Several rules claim to do what the VBE does to code it reads in: keyword and
identifier casing, spacing, literal spelling, `Call Foo()` and `Else x`,
labels in column 1, trailing whitespace. Those claims are checked against
the VBE itself, and the checks are kept as data the test suite replays
without Office.

## How it was measured

`tools/vbe_oracle.py record` drives a real Excel through
[pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness) and
writes `tests/oracle/vbe_rendering.json`. The current recording was made
with Excel 16.0, build 20326, on 2026-09-23.

- **Rendering.** 265 probes, each a few lines of deliberately untidy VBA,
  are pasted into modules with `CodeModule.AddFromString` and the modules
  exported again with `VBComponent.Export`. The export is what the VBE
  made of each line.
- **Whole modules.** Eleven questions about how names interact, and about
  blank lines, are asked in a workbook of their own each, because the VBE
  keeps one spelling per name for a whole project. Two of them import a
  module from a file with `VBComponents.Import`, the way an exported module
  goes back in.
- **Compile checks.** 26 small modules are compiled with the VBE's Debug >
  Compile, for the grammar questions a rule's safety depends on.
- **Run-time checks.** 9 functions return a number that answers a
  question, such as whether a statement after a comment ending in ` _`
  runs.

Each re-recording so far has reproduced every earlier result byte for byte.

`tests/test_oracle.py` formats every probe with the `vbe` preset and
compares the result with the VBE's, line for line, and checks the language
facts below. `tests/test_live.py` (run with `pytest -m live`) goes further,
on the fixture inputs:

- It formats each with the `vbe` preset, imports the result from a file
  into a new workbook of its own, exports it again, and requires the
  export to match line for line. On 2026-09-23 all 67 applicable inputs
  came back unchanged. The other 14 are the exceptions in the table at the
  end: text outside the code page, an empty module, a two-digit year,
  value-changing literals, and suppressed lines.
- It compiles each module before and after formatting with its own
  configuration, and requires the same outcome and message. All 30
  modules that formatting changes passed.

## What the VBE does

Casing:

- Keywords take one fixed spelling. `EndIf` is written `End If`, `DefLng a-c`
  is written `DefLng A-C`, and `Debug.Print` is capitalized.
- Contextual keywords are capitalized only in the statement that makes
  them keywords: `Option Compare Text` keeps `Text`, while a variable
  named `text` keeps its own spelling.
- The VBE keeps one spelling per name for the whole project. A name takes
  the spelling of its last declaration, in any procedure or module; a
  `Dim zqCount` in one procedure and a `Dim ZQCOUNT` in the next leave every
  use reading `ZQCOUNT`. A Type member is a declaration: a member named
  `second` makes the VBE write the library function `second(Now)`. So is a
  `Declare`'s parameter: a parameter `HWND` makes it write
  `Application.HWND`.
- The referenced libraries feed the same table, with every name they
  define, members included. An undeclared `sql` is written `Sql` (Excel has
  a `QueryTable.Sql`), and `ws.cells(1, 1).value` becomes
  `ws.Cells(1, 1).Value`. Where two libraries spell a name differently,
  the VBA library wins, then the host (`ID` and `Filename` are Excel's,
  over Office's `Id` and `FileName`). Parameter names are not in the
  table: `savechanges:=` and `prompt:=` stay as written. The library
  spellings come from pyVBAanalysis's object models, corrected where the
  VBE was measured to differ: it writes `Err.LastDllError`, where the
  models have `LastDLLError`.
- A name nothing declares takes the spelling of its last use.
- Compiler constants are spelled `VBA6`, `VBA7`, `Win16`, `Win32`, `Win64`
  and `Mac`. `twinbasic` is not one, and is left alone.

Spacing:

- Runs of spaces collapse to one, except indentation, the gap before an
  end-of-line comment that was wider than one space, the column of `As` in
  a declaration when it was widened, the statement after a `:` when it was
  widened, and the code after a label or line number.
- A statement that calls a procedure without `Call` gets a space before a
  parenthesized first argument, `Foo (x)`, and before a comma that follows
  the name directly, `Foo , 2`. Unary signs stay glued: `Foo -1`, `Step -1`.
- `=>`, `=<` and `><` are written `>=`, `<=` and `<>`.
- Tabs become spaces, and whitespace at the end of a line goes. A line
  holding nothing but whitespace is kept.
- A line the VBE cannot parse is stored exactly as written.

Literals:

- A Double keeps 15 significant digits and a Single 7: `1.0` becomes `1#`,
  `.5` becomes `0.5`, and a value too long for fixed notation gets an
  exponent (`1E+15`). A suffix the value makes redundant goes (`10%`,
  `32768&`, `&HFF%`). Hex is upper case without leading zeros, and `&17`
  becomes `&O17`.
- Dates are written `#M/D/YYYY h:mm:ss AM#`, with a midnight time dropped.

Statements and lines:

- `Call Foo()` becomes `Call Foo`, and a block `Else x` becomes `Else: x`.
- A line label or line number moves to column 1, and the statement after
  it keeps its column.
- `AddFromString` drops blank lines above the first line of code, but
  `Import` keeps blank lines at both ends of a module.

The language facts the formatter's safety depends on, from the compile and
run-time checks:

- `If x Then:` with nothing after the colon is a single-line If, not a
  block (`End If` after it does not compile), and every statement after its
  `Then`, colons included, belongs to it.
- A comment ending in ` _` continues onto the next line, for `'`, `Rem`,
  end-of-line and directive comments alike: the statement on the next
  line does not run.
- A line continuation followed by spaces still continues.
- `s&t` is a syntax error (the `&` is a type character), and so is a bang
  with spaces around it.
- `Foo, 2` and `Foo , 2` compile the same (both are rejected when the first
  argument is required) and run the same.

## Where the `vbe` preset differs, on purpose

| Case | What the VBE does | What the preset does, and why |
| --- | --- | --- |
| A Double or Single with more digits than the VBE keeps | Rounds it to 15 (or 7) digits, a different value | Leaves it and reports it: the rewrite would change the program. |
| A date with a two-digit year | Reads the year through the machine's date window | Leaves it: the result depends on the machine. |
| The name of a `Declare` with no `Alias` | Respells it like any name | Leaves it: it is also the DLL entry point, looked up case-sensitively. |
| A line suppressed with `'@prettyvba-ignore` | Formats it | Leaves it, as asked. |
| A keyword used as a member name (`obj.Print`) | Spelled it, in the one probe that asks, the way the word was last typed elsewhere in the module, keyword positions included | Spells it from the libraries. |
| Names declared in other modules | Spells them project-wide | Spells them only when the project is formatted together. |
| Blank lines at the top of a module read with `AddFromString` | Drops them | Keeps them, as `Import` does. |
| Text outside the Windows code page | Replaces it with `?` | Keeps it. |

The oracle replay lists the probes these affect, with the reason for each,
in `DEVIATIONS` and `ISOLATED_DEVIATIONS` in `tests/test_oracle.py`. A
deviation that starts matching the VBE fails the suite, so the list cannot
go stale.
