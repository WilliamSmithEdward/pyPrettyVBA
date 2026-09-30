# Notes for coding agents

pyPrettyVBA is a VBA formatter: a pipeline of named rules over a lossless
token model, with a check that refuses output which runs differently from
the input. Read `docs/architecture.md` first.

## Commands

```bash
pip install -e .[dev]
ruff check src tests tools scripts
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
- The package imports the standard library and pyOpenVBA, nothing else
  (`tests/test_packaging.py` checks). Other third-party packages are for
  development only.
- Keep text in files plain ASCII, apart from test data that is about
  non-ASCII text.

<!-- repo-standards:begin. Copied from WilliamSmithEdward/repo-standards, templates/agents/AGENTS-block.md. Change it there; the weekly rescan fails a copy that differs. -->
## Releases, CI and security

These rules are the same in every WilliamSmithEdward repository.

- **How a release happens here:** pushing a `vX.Y.Z` tag runs Publish, which uploads to PyPI and creates the GitHub release with its security and malware reports.
- **Starting a workflow by hand never releases anything.** Publish and every
  release report are dry runs when started with `gh workflow run` or the Run
  workflow button. They build, scan and assemble the release files exactly
  as a release would, and upload them as the `release-preview` artifact
  instead. Run one after changing anything on the release path:
  `gh workflow run <file> --ref main`, then
  `gh run download <run-id> -n release-preview`.
- **Do not create, publish, edit or delete a release or a `v*` tag** unless
  the owner asks for it. A `v*` tag cannot be moved or deleted once pushed.
- **Every change to `main` goes through a pull request** that passes CI
  passed, Security passed and Malware scan passed. No one can push to `main`
  directly or skip the checks, admins included. Push a branch, open a pull
  request, and let it merge itself: `gh pr merge --auto --squash <number>`.
- **Pins.** Actions by full commit SHA with the version as a comment. Images
  by digest, in `.github/security/<tool>/Dockerfile`. Python tools from the
  hash-locked `.github/requirements/<purpose>.txt`, compiled from the `.in`
  beside it with
  `uv pip compile <purpose>.in --universal --generate-hashes --python-version 3.12 -o <purpose>.txt`.
  Runners are named releases, never `-latest`.
- **Updates merge themselves.** Dependabot and the Update YARA rules workflow
  open pull requests that merge once the three checks pass, except a
  third-party major version, which waits for the owner. Leave them alone
  unless asked.
- **A scanner finding is fixed or accepted with a written reason** in the
  repository's accepted list. Never silence a scanner without one.
<!-- repo-standards:end -->
