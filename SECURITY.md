# Security policy

## Reporting a vulnerability

Please report a vulnerability privately, not in a public issue. Use
[Report a vulnerability](https://github.com/WilliamSmithEdward/pyPrettyVBA/security/advisories/new)
on the repository's Security tab. It opens a draft advisory that only
you and the maintainer can see.

A useful report names the pyPrettyVBA version, the Python version and the
operating system, and includes the module or Office file that shows the
problem, cut down as far as it will go.

## Supported versions

Only the latest release on PyPI is supported. A fix ships in a new
release, and earlier versions do not get one.

## What to report

pyPrettyVBA reads VBA modules and Office files that may come from anyone,
and writes them back in place, so the things to worry about are an input
file and a write. For example:

- a module or Office file that makes the formatter hang, run out of
  memory or crash the interpreter;
- a write that touches any file other than the one being formatted, or
  leaves that file damaged when the write fails;
- output that runs differently from the input and gets past the safety
  check;
- a digitally signed VBA project written without `--remove-signatures`.

pyPrettyVBA never runs the code it formats. It does not start Office,
execute VBA, open a network connection or run a command. Reading and
writing Office files goes through its one dependency,
[pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA), which has
its own policy.

## How the code is checked

Every push to main, every pull request and every release is checked by
two workflows, and each fails on anything it does not expect.

- **Security:** CodeQL (Python and GitHub Actions, security-extended
  queries) and Semgrep (the default, Python, security-audit, secrets and
  GitHub Actions rule sets) scan the package, the workflows and the scan
  scripts. They also run weekly, so new rules reach code that has not
  changed. A finding fails the scan unless
  [.github/security/accepted.toml](.github/security/accepted.toml) lists it
  with the reason it is accepted.
- **Malware scan:** ClamAV, with signatures fetched fresh on every run,
  and YARA-X, with the full [YARA Forge](https://github.com/YARAHQ/yara-forge)
  rule collection, scan every file the commit holds, test fixtures
  included. They also run daily, so new signatures and rules reach files
  that have not changed. A match fails the scan unless
  [.github/security/malware-accepted.toml](.github/security/malware-accepted.toml)
  lists it with the reason it is accepted.

In both lists an entry that no longer matches fails the scan too, so the
lists cannot outlive what they excuse.

- **Fuzz:** Atheris feeds generated text to the lexer and to the formatter
  under every preset, in [fuzz.yml](.github/workflows/fuzz.yml), daily and
  on every change to the package. The lexer must lose nothing, and the
  formatter must not raise, must settle in one pass, and must change no
  character outside ASCII and whitespace. It is not a gate: a finding fails
  that workflow and becomes a regression seed in `tests/fuzz_corpus/vba`,
  which the test suite replays.

A release is published only after its commit passes both, and it carries
the reports as `pyprettyvba-<version>-security-report.md` and
`pyprettyvba-<version>-malware-report.md`, beside the SARIF the security
report was made from. PyPI receives the release through Trusted
Publishing, so no upload token exists to leak.

## Pinned tools

Every action the workflows use is pinned to a commit, the Semgrep and
ClamAV images to a digest, the development tools to exact versions, and
the test and build tools to hash-locked files in .github/requirements. Dependabot proposes updates to all of them, each a week after
its release. The YARA-X engine and the YARA Forge rules are pinned by
release and SHA-256 in
[.github/security/yara.json](.github/security/yara.json); a weekly workflow
proposes new pins in a pull request, and the scans check that pull
request with the new rules before it can be merged.

[OpenSSF Scorecard](https://scorecard.dev/viewer/?uri=github.com/WilliamSmithEdward/pyPrettyVBA)
rates these practices on every change to main and weekly, and publishes
the result the README badge shows. Some of its checks assume more than one
maintainer, such as a second person approving every change, so a
single-maintainer project cannot score full marks on them.

## Verifying a download

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
