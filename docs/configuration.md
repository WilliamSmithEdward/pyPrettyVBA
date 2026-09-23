# Configuration

pyPrettyVBA reads its settings from a TOML table. It looks for the first of
these files in the directory of the file being formatted, then in each
directory above it:

| File | What is read |
| --- | --- |
| `pyprettyvba.toml` | the whole file |
| `.pyprettyvba.toml` | the whole file |
| `pyproject.toml` | its `[tool.pyprettyvba]` table, if it has one |

With no file, the `default` preset applies. `pyprettyvba config PATH` prints
the settings that apply to a file and where they came from, and
`pyprettyvba init` writes a commented starter file.

Every key is optional. An unknown key, rule or option, or a value of the
wrong type, is an error that names the closest valid spelling.

## A complete example

```toml
preset = "default"            # default | vbe | xlide | strict | minimal | none
extend = "../shared.toml"     # start from another configuration file

indent-width = 4
indent-style = "space"        # space | tab
tab-width = 4                 # the columns a tab advances to, when reading
line-ending = "auto"          # auto | crlf | lf
encoding = "auto"             # auto, or a codec name such as "cp1252"
hosts = ["excel"]             # the libraries identifier-case spells names from

include = ["*.bas", "*.cls", "*.frm", "*.doccls"]
exclude = [".git", "node_modules"]
extend-exclude = ["vendor/"]

[rules]
blank-lines = { max-consecutive = 1 }
comment-space = true
numeric-literals = false
indent = { case = false, directives = "column-zero" }

[[overrides]]
files = ["legacy/**"]
exclude-files = ["legacy/keep/**"]
indent-width = 2
rules = { identifier-case = false }
```

## Presets

A preset decides which rules run and with which options before the
`[rules]` table changes anything.

| Preset | Rules |
| --- | --- |
| `default` | Every rule that is on by default: casing, spacing, literals, statement forms, indentation, continuation lines, end-of-line comments, blank lines, trailing whitespace, end of file, line endings. |
| `vbe` | Only the rules that do what the VBE does when it reads code in, set the way the VBE behaves: indentation kept as written (labels still move to column 1), one spelling per name for the whole module, whitespace-only lines and blank lines at either end kept. |
| `xlide` | XLIDE's Format Document: indentation, continuation lines, keyword and identifier casing (members and named arguments left alone), spaces inserted but never removed, blank lines indented to the code around them. |
| `strict` | `default` plus `comment-space`, `rem-comments`, `let-keyword`, `split-statements` and `align-declarations`; at most one blank line in a row; lines over 120 columns reported. |
| `minimal` | `trailing-whitespace`, `end-of-file` and `line-endings`. |
| `none` | Only `suppression-directive`, which checks the directives themselves. |

## Global settings

| Key | Default | Meaning |
| --- | --- | --- |
| `indent-width` | `4` | Columns per indentation level (1 to 16). |
| `indent-style` | `"space"` | Indent with `space`s or `tab`s. |
| `tab-width` | `4` | How far a tab advances, when measuring existing indentation (1 to 16). |
| `line-ending` | `"auto"` | `auto` keeps the ending a file mostly uses; `crlf` or `lf` forces one. The VBE writes CRLF, and `crlf` is the safe choice for code that goes back into Office. |
| `encoding` | `"auto"` | How files are read and written. `auto` reads UTF-8 (with or without a byte order mark) when the bytes are valid UTF-8 and the Windows ANSI code page otherwise, and writes a file back the way it was read. |
| `hosts` | `["excel"]` | The Office libraries whose names identifier-case knows: any of `excel`, `word`, `powerpoint`, `access`. The VBA library is always known. |

## Choosing files

When a directory is formatted, a file is taken when its path, relative to
the configuration file's directory, matches an `include` pattern and no
`exclude` or `extend-exclude` pattern. Patterns are gitignore-style globs,
matched without regard to case: a pattern without a `/` matches a file or
directory name at any depth, one with a `/` is anchored to the
configuration's directory, `**` crosses directories, and a pattern that
matches a directory covers everything under it. `extend-exclude` adds to
the default `exclude` list instead of replacing it. Files named on the
command line are always formatted.

## Rules

The `[rules]` table sets each rule by its name (see [rules.md](rules.md)):

```toml
[rules]
comment-space = true                        # enable, with its default options
numeric-literals = false                    # disable
blank-lines = { max-consecutive = 1 }       # enable and set options
indent = { enabled = false, case = false }  # set options but keep it off
```

Options not given keep the preset's value, and the preset's value is the
rule's default unless the preset says otherwise.

## Overrides

Each `[[overrides]]` table applies to the files its `files` patterns match
and its `exclude-files` patterns do not, matched like `include`. It may set
the global settings above and a `rules` table, which layer over the file's
own: a rule set to a table there keeps the options it had and changes the
ones listed. When several overrides match a file, they apply in order.

## Extending another file

`extend` names a configuration file, relative to the one that names it,
whose settings apply first. The extending file's keys replace the base's,
except `rules`, which is merged rule by rule. Chains are followed; a cycle is
an error.

## The command line

The command-line options change the configuration found for each file:

| Option | Effect |
| --- | --- |
| `--config FILE` | Use this configuration file instead of looking for one. |
| `--isolated` | Use no configuration file at all. |
| `--preset NAME` | Start from this preset. |
| `--enable RULE`, `--disable RULE` | Switch a rule on or off (repeatable). |
| `--set RULE.OPTION=VALUE` | Set an option (repeatable). The value is read as TOML, and taken as a string when it is not valid TOML: `--set indent.case=false`, `--set trailing-comments.position=align`. |
| `--indent-width`, `--indent-style`, `--line-ending` | Override those settings. |
| `--stdin-filename NAME` | The name to treat standard input as, for overrides and the module kind. |
| `--no-project-casing` | Format each file on its own, without the names the other modules declare. |

When a directory is formatted, the files that share a configuration are one
project: a name one module declares takes that module's spelling in the
others, the way the VBE keeps one spelling per name in a project.
