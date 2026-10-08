# Changelog

## Unreleased

- `indent` and `continuation-indent` indented a line that held only a
  bare `_`, which the lexer has no rule for; indented, ` _` at the end of
  the line became a continuation, and the safety check refused the
  module. A line, or a continuation line, that starts with such text now
  keeps its indentation. Found by the fuzz run.
- `align-declarations` padded the space before `As` on a line holding text
  the lexer has no rule for, moving a byte that touched the `As`, and the
  safety check refused the module. Like every other rule, it now leaves
  such a line as written. Found by the fuzz run.
- An `Attribute` line ending in a comment with a space after it could not
  be formatted: `trailing-whitespace` trimmed the space, and the safety
  check compared the comment on an Attribute line by its exact text, where
  a comment anywhere else compares by what it says. It now compares the
  same way. Found by the fuzz run.
- A new rule, `wrap-lines`, off by default: a line over `max` columns is
  broken after a comma, a spaced operator or the `Then` of a single-line
  If, at the rightmost point that fits, with the continuation one level in.
  A line with no break point is left for `max-line-length` to report. With
  `pad-blocks`, a statement that spans lines gets a blank line on each side.
- `blank-lines` gains `pad-blocks`, off by default: one blank line inside
  every block and on both sides of every block within a procedure, so each
  body stands apart. Types and Enums are set apart like procedures, and
  the Option statements get a blank line below. An `#If` that splices a
  line, such as two versions of a procedure header, is left alone.
- A new rule, `one-declaration-per-line`, off by default: `Dim a As Long,
  b As Long` becomes a `Dim` for each. Each item keeps its own type, so
  `Dim x, y As Long` splits into `Dim x` and `Dim y As Long`, which is what
  it declared. The safety check knows the two forms mean the same thing.
- `date-literals` rewrote a year under 100 written with extra digits, such
  as `#1/15/0022#`, to `#1/15/22#`, and the safety check refused the
  module, since the two do not read back the same way. The VBE windows any
  year under 100 however it is written (`#1/15/0022#` is 2022, measured),
  so such a literal is now left alone, like `#1/15/22#`. A year of 100 or
  more is still written as the VBE writes it. Found by the fuzz run.
- The Malware scan builds the wheel and sdist with the hash-locked build
  tools, as a release builds them, and ClamAV and YARA-X scan them beside
  the committed files.
- The Security workflow audits the runtime dependency, pyOpenVBA, with
  pip-audit: any known vulnerability fails it. The audit reads a hash-locked
  `.github/requirements/runtime.txt` that Dependabot moves daily.
- `comment-space` left a comment on an `Attribute` line alone only in the
  module header. On a member attribute inside a procedure, such as
  `Attribute Item.VB_UserMemId = 0 'default member`, it added a space, and
  the safety check refused the whole module, so `strict` could not format
  it. No rule edits an Attribute line, and now this one does not either.
  Found by fuzzing.
- `numeric-literals` could join numbers into a date literal. Between two `#`
  signs a type suffix can be all that keeps numbers apart: `x = #0% 0#` is
  `#`, `0%` and `0#`, but without the redundant `%` it reads as the date
  `#0 0#`. The rule checked only the literal's neighbours, so the safety
  check refused the module. It now relexes the whole logical line with every
  respelling of the pass, and keeps the `%` where removing it would make a
  date. Found by the daily fuzz run.
- Coverage-guided fuzzing with Atheris of the lexer and of the formatter
  under every preset (`fuzz/fuzz_formatter.py`), daily and on every change
  to main. The test suite replays the seed corpus on every run, pull
  requests included.
- Releases carry signed build provenance: the signed bundle from GitHub's
  artifact attestations goes on the GitHub release as
  `pyprettyvba-<version>.sigstore.json`. `SECURITY.md` has the steps to
  verify a download.
- OpenSSF Scorecard rates the repository's security practices on every
  change to main and weekly, and the README shows its badge.
- CI and the release's test step install the test tools and pyOpenVBA from a
  hash-locked lock (`.github/requirements/test.txt`) instead of the dev extra.

## 0.1.0 (2026-09-29)

The first version.

- Twenty rules: keyword and identifier casing, spacing, numeric and date
  literals, statement forms, statement splitting, indentation, continuation
  lines, declaration alignment, end-of-line comments, comment style, blank
  lines, trailing whitespace, end of file, line endings, line length, and
  directive checking. Each can be enabled, disabled and configured.
- Presets: `default`, `vbe` (only what the VBE does), `xlide` (XLIDE's
  Format Document), `strict`, `minimal` and `none`.
- Suppression with `'@prettyvba-ignore` comments, for a line, the next line,
  a region or a module, per rule.
- Configuration in `pyprettyvba.toml` or `pyproject.toml`, with per-file
  overrides and `extend`.
- A command line (`format`, `check`, `rules`, `config`, `init`) with text,
  grouped, JSON and GitHub output, and a Python API.
- Projects formatted together share their modules' spellings.
- The VBA inside Excel, Word, PowerPoint and Access files, checked and
  formatted in place through pyOpenVBA.
- A digitally signed VBA project is left alone unless `--remove-signatures`
  is given; then it is formatted and its signature removed.
- A safety check that refuses output which runs differently from the input.
- Rule behavior measured against Excel's VBE and replayed by the tests.
- Python 3.11 or later, and one dependency: pyOpenVBA, which is pure Python
  with no dependencies of its own.
