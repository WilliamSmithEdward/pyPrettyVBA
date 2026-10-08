# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/pyPrettyVBA/security/advisories/new).
Only the maintainer sees it. Include the pyPrettyVBA version, the Python
version and the operating system, and the smallest module, Office file or
steps that show it, with credentials and private data removed.

A confirmed vulnerability is fixed in a release on PyPI, and the advisory is
published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on PyPI receives security fixes. Older
releases are not maintained separately; update when a fix ships.

## Scope

pyPrettyVBA reads VBA modules and Office files that may come from anyone,
and writes them back in place, so the things to worry about are an input
file and a write. These count as vulnerabilities:

- a module or Office file that makes the formatter hang, run out of
  memory or crash the interpreter;
- a write that touches any file other than the one being formatted, or
  leaves that file damaged when the write fails;
- output that runs differently from the input and gets past the safety
  check;
- a digitally signed VBA project written without `--remove-signatures`.

pyPrettyVBA never runs the code it formats. It does not start Office,
execute VBA, open a network connection or run a command. Reading and
writing Office files goes through its one runtime dependency,
[pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA), which has
its own policy.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security also runs weekly, so new queries and rules reach code
that has not changed, and Malware scan runs daily, so new signatures and
rules reach files that have not changed.

- **Code:** CodeQL with GitHub's security-extended queries, for Python and
  GitHub Actions, and Semgrep with the default, Python, security-audit,
  secrets and GitHub Actions rule sets. Both scan the package, the
  workflows that build and publish it, and the scripts in
  `scripts/security` that judge the scans. A `nosemgrep` comment cannot
  hide a finding. Results go to the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** pip-audit checks the one runtime dependency,
  pyOpenVBA, as a fresh install resolves it today
  (`.github/requirements/runtime.txt`, hash-locked and moved daily by
  Dependabot). Any known vulnerability fails Security. pyOpenVBA has no
  dependencies of its own and is scanned in its own repository. The tools
  the workflows install come from hash-locked files (see Pinning and
  updates).
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every file the commit holds, test fixtures included,
  and the wheel and sdist built from it with the hash-locked build tools,
  as a release builds them. YARA-X runs the full
  [YARA Forge](https://github.com/YARAHQ/yara-forge) rule collection. A
  scan error fails the report as a match does.
- **Fuzzing:** Atheris drives two targets in `fuzz/fuzz_formatter.py` with
  generated text: the lexer, which must lose nothing, and the formatter
  under every preset, which must not raise, must format its own output to
  itself, must change no character outside ASCII and whitespace, and under
  the `none` preset must change nothing. Both start from the seeds in
  `tests/fuzz_corpus/vba`. The Fuzz workflow runs on every change to main
  that touches the package, the fuzz targets or the corpus, and daily. It
  does not run on pull requests, where its random search would fail for
  code the pull request did not touch; a finding becomes a regression
  test with its fix, a seed that `tests/test_fuzz_corpus.py` replays on
  every CI run, pull requests included.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Its Code-Review and Contributors checks assume more than one
  maintainer, such as a second person approving every change, so a
  single-maintainer project cannot score full marks on them.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[.github/security/accepted.toml](.github/security/accepted.toml) for
CodeQL and Semgrep, or
[.github/security/malware-accepted.toml](.github/security/malware-accepted.toml)
for ClamAV and YARA-X. An entry matches on the tool, the rule and the
file, and in accepted.toml also the text of the flagged line, so an edited
line needs another review, and an entry that no longer matches fails the
report. zizmor keeps its exceptions in `.github/zizmor.yml` or inline
beside the line they excuse, each with its reason.

The current entries:

- Semgrep `python37-compatibility-importlib2`, in
  `src/pyprettyvba/names.py`: a compatibility warning about
  `importlib.resources`, which every Python the package supports (3.11 and
  later) has.
- malware-accepted.toml has none.
- zizmor's `self-repository` and `superfluous-actions` rules are turned
  off in `.github/zizmor.yml`, each with its reason and when it comes back.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, the development tools to exact versions in
`pyproject.toml`, and the YARA-X engine and YARA Forge rules to a release
and its SHA-256. The runtime dependency on pyOpenVBA stays a version
range for the package's users; the test and fuzz locks pin it exactly.
ClamAV's signatures change too often to pin, so freshclam fetches and
verifies them on every run.

Dependabot proposes updates to the GitHub Actions, the Semgrep and ClamAV
images, the development tools in `pyproject.toml` and the hash-locked
tools in `.github/requirements` once a version is a week old, and at once
for a security advisory. A new pyOpenVBA release skips the week's wait,
since it is the same owner's package. The Update YARA rules workflow
proposes new YARA pins in `.github/security/yara.json` each week. A minor
or patch update, and the YARA pull request, merges itself once CI,
Security and Malware scan pass; a third-party major version waits for
review.

## Releases

A pushed `v*.*.*` tag checks that the tag matches `__version__` in
`src/pyprettyvba/__init__.py`, runs the test suite, builds the sdist and
wheel, and runs Security and Malware scan on the tagged commit. Nothing is
published unless all of them pass. The distributions go to PyPI through
Trusted Publishing, so no upload token exists to leak. The GitHub release
carries the distributions, `pyprettyvba-<version>-security-report.md` and
`pyprettyvba-<version>-malware-report.md` beside the scan results they
were made from, and the provenance bundle. Started by hand, the Publish
workflow is always a dry run and publishes nothing.

### Verifying a download

Every file on PyPI carries PyPI's own provenance, which names this
repository's `publish.yml` as the publisher; the file's page on PyPI shows it.
Releases published after 2026-09-30 also carry a GitHub build provenance
attestation, which you can check against any copy of the file, from PyPI or
from the GitHub release:

```
pip download pyprettyvba --no-deps -d check
gh attestation verify check/<file> --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file. The
signed bundle is also attached to the GitHub release as
`pyprettyvba-<version>.sigstore.json`, so the check works without asking
GitHub for it: add `--bundle pyprettyvba-<version>.sigstore.json`.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
