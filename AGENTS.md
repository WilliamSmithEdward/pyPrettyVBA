# Notes for coding agents

pyPrettyVBA is a VBA formatter: a pipeline of named rules over a lossless
token model, with a check that refuses output which runs differently from
the input. Read `docs/architecture.md` first.

## Commands

```bash
pip install -e .[dev]
ruff check src tests tools
mypy
python -m pytest                             # fast; no Office needed
python tools/fixtures.py bless [PREFIX]      # rewrite expected fixture files
python tools/build_docs.py                   # regenerate docs/rules.md
```

## Rules of the road

- Formatting must never change what code does. Do not loosen
  `safety.first_difference` to make a test pass; a `SafetyError` means a
  rule is wrong.
- Every rule claim about what the VBE does needs evidence in
  `tests/oracle/vbe_rendering.json`. Recording it needs Windows, Excel and
  pyVBAharness (`tools/vbe_oracle.py record`); say so rather than guessing.
- Bless fixtures only after reading the diff of the expected files, and
  explain every changed output.
- A rule's docstring is its user documentation (`docs/rules.md` is built
  from it, and `tests/test_docs.py` fails when it is stale).
- Keep text in files plain ASCII, apart from test data that is about
  non-ASCII text.
