# Contributing

## Set up

```bash
python -m venv .venv
.venv/Scripts/activate            # Windows; source .venv/bin/activate elsewhere
pip install -e .[dev]             # pytest, hypothesis, ruff, mypy
pip install -e .[verify]          # pyVBAanalysis, and pyVBAharness on Windows
```

## Check

```bash
ruff check src tests tools
mypy                              # strict, on src/pyprettyvba
python -m pytest                  # everything that needs no Office
```

The suite runs in a few seconds. Longer and heavier runs:

```bash
HYPOTHESIS_PROFILE=thorough python -m pytest tests/test_properties.py
PYPRETTYVBA_CORPUS=path/to/modules python -m pytest -m corpus
python -m pytest -m live          # Windows, desktop Excel, pyVBAharness
```

The property tests generate random token soup for every preset, and whole
modules scrambled in every way the default preset repairs; `thorough` runs
3000 examples of each. The corpus tests format every module under the given
directories with every preset and compare pyVBAanalysis's diagnostics before
and after. The live tests import the `vbe` preset's output into a real VBE
and compile fixtures before and after formatting.

## Fixtures

Each directory under `tests/fixtures/` with a `case.toml` is a case: an
input, a configuration, and the expected output and report (the layout is
described in `tests/fixture_support.py`). To add one, write `case.toml` and
the input, then let the formatter write the expected files and read what it
wrote:

```bash
python tools/fixtures.py bless rules/indent/my-case
git diff tests/fixtures
```

`bless` with no argument rewrites every case; the diff is the review.

## Adding a rule

1. Write the rule class in the `rules/` module for its area. Its docstring is
   its documentation: say what it changes, what it leaves alone and why,
   and what the VBE does, with the evidence.
2. Add it to `RULES` in `rules/__init__.py` at the point in the pipeline
   where it belongs, and to the presets in `config.py` that should run it.
3. Add fixture cases under `tests/fixtures/rules/<code>/`: a basic case, and
   one for each option.
4. Run `python tools/build_docs.py` to regenerate `docs/rules.md`.

A rule must never change what code does. Leave the gaps around an
`UNKNOWN` token alone; never add or remove the space between a name and a
glued `&`, `!`, `#`, `^` or `.`; and check what a changed line reads as at
the start of a line (a label? a line number? a directive?). The engine's
safety check refuses output that runs differently, and the property tests
go looking for such output; a `SafetyError` in a test is a bug in the
rule, never in the check.

## Recording what the VBE does

`tools/vbe_oracle.py record` re-records `tests/oracle/vbe_rendering.json`
from a real Excel. Record on purpose, read the diff, and keep the probe
names' `zq` prefix: the VBE keeps one spelling per name for a whole
project, and the prefix keeps probes from changing each other's casing.
Questions about how names interact go in `ISOLATED_PROBES`, which run in a
workbook each.

## Other tools

- `tools/build_names.py` rebuilds `src/pyprettyvba/data/names.json.gz`, the
  library names identifier-case spells, from pyVBAanalysis's object models.
- `tools/xlide_differential.py PATH` formats modules with the `xlide`
  preset and with XLIDE's own Format Document, and lists where they differ.
- `tools/build_gallery.py` renders every fixture case as one
  self-contained page (`artifacts/gallery/index.html`): input beside output
  with the changes marked, the configuration, and the report.
