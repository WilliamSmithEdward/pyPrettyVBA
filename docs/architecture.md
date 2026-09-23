# Architecture

pyPrettyVBA is a pipeline of rules over a lossless model of a module. This
page follows a module through it; the module docstrings carry the detail.

## Reading: `textio`, `lexer`, `document`

`textio.decode` reads a file's bytes: UTF-8 when they are valid UTF-8 (with
or without a byte order mark), otherwise the Windows ANSI code page, with
`surrogateescape` so that no byte is ever lost. `encode` writes the text
back the same way.

`document.split_header` sets aside what the VBE writes above the code of an
exported module: `VERSION`, `Object =`, the `BEGIN ... END` block of a class
or form, and the leading `Attribute VB_` lines. The header is never
formatted, apart from its line endings.

`lexer.tokenize` turns the body into tokens, following the MS-VBAL lexical
grammar the way the VBE reads it. Every character belongs to exactly one
token, whitespace, line continuations and line breaks included, so joining
the tokens gives the text back. Where the VBE departs from the
specification, the lexer follows the VBE: a comment ending in ` _` runs
onto the next line, and ` _` followed by spaces is still a continuation.
Text the lexer has no rule for becomes an `UNKNOWN` token, which the rules
leave alone.

`document.Document` groups the tokens into logical lines (a line and its
continuations), finds each line's label, comment and statements, and
classifies the statements (`If` block or single-line `If`, `Select Case`,
`Case`, loop, procedure header, declaration, directive and so on).

`structure.analyze_structure` walks the statements and gives every line its
block depth. It understands `#If` blocks, including the common one whose
arms each open the same procedure, and it reports what it cannot match (an
`End If` with no `If`); the indentation rules leave such a region as
written.

`scopes.collect_declarations` finds what the module declares and where:
each procedure's parameters, locals and labels, and the module's variables,
constants, procedures, Types, Enums and their members.

## Rules: `rules/`

A rule reads a `Document` and yields `Finding`s: a span of the text and the
text to put there, or no replacement for something it can only report.
Rules never edit text themselves. Each has a code, a summary, a category,
options, and flags for whether it is on by default and whether it belongs
to what the VBE itself does (the `vbe` preset). `rules/__init__.py` lists
them in the order they run.

## Running: `engine`

`engine.run_pipeline` runs the rules one after another, each on the text
the rules before it produced:

1. The rule's findings are filtered through the suppression directives of
   the current text (`suppression.scan_suppressions`).
2. Its fixes are applied together, dropping any that overlap one kept
   earlier, and the text is read again.
3. Every finding is recorded with the stage it was found at. Each stage
   records its edits, so a position in any later text maps back to the
   original: a violation always names the line and column the user wrote.

The whole pipeline repeats until a pass changes nothing, so the output is a
fixed point: formatting it again changes nothing. The rules are ordered so
that one pass is enough and the second is the proof; a pipeline that still
changes the text after four passes is an error.

## The safety check: `safety`

Before `run_pipeline` returns, `safety.first_difference` compares the input
and the output. Each is reduced to its statements, and each statement to
normalized tokens: names by their case-folded key (except the name of a
`Declare` with no `Alias`, compared exactly), numbers and dates by type and
value, comments by their text, `:` and a line break as the same separator
(except inside a single-line `If`). Where whitespace decides the parse, the
signature records it: whether `&`, `!`, `#`, `^` and `.` touch the tokens
around them. Output whose signature differs from the input's is refused
with `SafetyError`, and the input is left untouched.

## The API and the command line: `api`, `cli`, `config`

`config.Config` holds a configuration: a preset, global settings, rule
settings and per-file overrides, validated as they are read, and
`settings_for(path)` resolves them for one file. `api.format_source`
formats text, `format_file` a file, and `format_paths` directories, as
projects: the files that share a configuration are formatted together, so
that `project_names` can give each module the spellings the others declare.
`cli.main` is the `pyprettyvba` command.

## Tests

- `tests/test_*.py` unit tests, one file per module.
- `tests/fixtures/`: a hundred cases, each an input, a configuration and
  the expected output and report, grouped by what they show (rules, presets,
  suppression, configuration, file handling, safety, the command line).
  `tools/fixtures.py bless` rewrites the expected files from current
  behavior, for review in the diff.
- `tests/test_properties.py`: Hypothesis properties over random token soup
  (every preset is safe and settles; the lexer loses nothing) and over
  generated modules scrambled in every way the default preset repairs,
  which must format back to their canonical form byte for byte.
- `tests/test_oracle.py`: the recorded VBE behavior, replayed
  (see [vbe-evidence.md](vbe-evidence.md)).
- `tests/test_corpus.py`: every preset over a directory of real modules
  (`PYPRETTYVBA_CORPUS`), with pyVBAanalysis's diagnostics compared before
  and after.
- `tests/test_live.py`: the `vbe` preset's output imported into a real VBE,
  and compile outcomes before and after formatting (`pytest -m live`).
- `tests/test_docs.py`: `docs/rules.md` matches the rules.
