# pyPrettyVBA

A formatter for VBA modules (`.bas`, `.cls`, `.frm`), in the spirit of
Prettier but configurable: every change it makes belongs to a named rule
that can be switched off, tuned, or suppressed for one line, a region or a
whole module. It spells code the way the VBE spells it, lays it out the way
you configure, and refuses to return output that would run differently from
the input.

Before:

```vba
public function RestockList(ws as worksheet) as collection
dim r as long,qty as long,result as new collection
on error goto fail
for r=2 to ws.cells(ws.rows.count,1).end(xlup).row
qty=ws.cells(r,3).value
if qty<low_stock then
result.add ws.cells(r,1).value
elseif qty=0 then result.add ws.cells(r,1).value & " (out)"
end if
next r
```

After `pyprettyvba format`:

```vba
Public Function RestockList(ws As Worksheet) As Collection
    Dim r As Long, qty As Long, result As New Collection
    On Error GoTo fail
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        qty = ws.Cells(r, 3).Value
        If qty < LOW_STOCK Then
            result.Add ws.Cells(r, 1).Value
        ElseIf qty = 0 Then result.Add ws.Cells(r, 1).Value & " (out)"
        End If
    Next r
```

The whole module, before and after, is in
[tests/fixtures/showcase](tests/fixtures/showcase/before-and-after/).

## Install

pyPrettyVBA needs Python 3.11 or later and nothing else: it uses only the
standard library. It is not on PyPI yet; install it from a checkout:

```bash
pip install -e .
```

## Use

```bash
pyprettyvba format src/            # format every module under src/ in place
pyprettyvba format --check src/    # exit 1 if anything would change
pyprettyvba format --diff src/     # show the changes, write nothing
pyprettyvba check src/             # list what each rule found, by line
pyprettyvba check --fix src/       # fix what can be fixed, list the rest
pyprettyvba rules                  # every rule, on or off, fixable or not
pyprettyvba rules indent           # one rule's options
pyprettyvba config src/Module1.bas # the settings that apply to a file
pyprettyvba init                   # write a starter pyprettyvba.toml
```

`check` prints one violation per line, `file:line:column: rule message`, with
`[*]` on the ones formatting fixes. `--output-format` also offers `grouped`,
`json` and `github` (workflow annotations), and `--statistics` counts them by
rule. A path of `-` reads standard input. Exit codes: 0 clean, 1 findings or
changes, 2 an error.

From Python:

```python
from pyprettyvba import Config, format_source

result = format_source(source_text, Config({"preset": "vbe"}))
result.output       # the formatted text
result.violations   # what each rule found, with line and column in the input
```

`format_file` and `format_paths` read and write files; `project_names`
gathers what a project's modules declare, so that a name declared in one
module is spelled the same way in the others.

## Presets

| Preset | What it does |
| --- | --- |
| `default` | What the VBE does to each line (casing, spacing, literal spelling), plus indentation, blank lines and line endings. |
| `vbe` | Only what the VBE itself does to code it reads in, so the VBE has nothing left to change (the exceptions are listed in [docs/vbe-evidence.md](docs/vbe-evidence.md)). |
| `xlide` | What XLIDE's Format Document does: indentation, casing and inserted spaces, never removing a space or a line. |
| `strict` | The defaults plus every style rule: one statement per line, `'` comments with a space, no `Let`, aligned declarations. |
| `minimal` | Whitespace only: trailing whitespace, the end of the file, line endings. |
| `none` | Nothing but checking suppression directives; a base for enabling rules one by one. |

## Rules

Twenty rules, each documented with its options in
[docs/rules.md](docs/rules.md): keyword and identifier casing, spacing,
numeric and date literal spelling, statement forms, indentation (block
structure, `Case` arms, `#If` blocks, labels, continuation lines), end-of-line
comments, blank lines, trailing whitespace, line endings, and a handful of
optional style rules.

## Configuration

A `pyprettyvba.toml` (or a `[tool.pyprettyvba]` table in `pyproject.toml`)
next to the code, or in any directory above it:

```toml
preset = "default"
indent-width = 4
line-ending = "crlf"

[rules]
blank-lines = { max-consecutive = 1 }
comment-space = true
numeric-literals = false

[[overrides]]
files = ["legacy/**"]
rules = { indent = false }
```

See [docs/configuration.md](docs/configuration.md) for every key.

## Suppression

```vba
x   =  1   '@prettyvba-ignore: spacing -- aligned on purpose
'@prettyvba-ignore-next-line: indent
'@prettyvba-ignore-start
    table(0) = Array("id",   "name")
'@prettyvba-ignore-end
```

See [docs/suppression.md](docs/suppression.md).

## Safety

Whitespace and letter case are not always meaningless in VBA. `s&t` is a
syntax error where `s & t` concatenates; `Foo .Bar` passes a With member
where `Foo.Bar` calls one; a comment ending in ` _` swallows the next line;
a `Declare` without an `Alias` looks its entry point up case-sensitively.
Before it returns anything, the formatter reduces the input and the output
to the statements they run and compares them, and it refuses output that
differs. Property-based tests, a corpus of 422 real modules, and the VBE
itself (below) check that this never has to happen.

## What the VBE does, measured

The rules that claim to do what the VBE does are checked against the VBE:
probe modules pasted into Excel and exported again, compile checks, run-time
checks, and whole modules imported from files. The recording is replayed by
the test suite without Office. See [docs/vbe-evidence.md](docs/vbe-evidence.md).

## Development

```bash
pip install -e .[dev]
python -m pytest                 # unit tests, fixtures, properties, oracle replay
python -m pytest -m live         # against a real VBE (Windows, Excel, pyVBAharness)
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the fixture workflow, the corpus
and property tests, and how to add a rule.

## License

MIT. See [LICENSE](LICENSE).
