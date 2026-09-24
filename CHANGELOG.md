# Changelog

## 0.1.0 (unreleased)

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
