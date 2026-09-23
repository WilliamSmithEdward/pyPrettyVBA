# Suppression

A comment that starts with `'@prettyvba-ignore` switches rules off for part
of a module. Nothing a suppressed rule would change there is changed, and
nothing it finds is reported.

```vba
'@prettyvba-ignore-file: identifier-case -- the whole module

Sub Demo()
    x   =  1   '@prettyvba-ignore: spacing -- this line
    '@prettyvba-ignore-next-line: indent -- the next line
          y = 2
    '@prettyvba-ignore-start: indent, spacing -- from here...
      z  =  3
    '@prettyvba-ignore-end -- ...to here
End Sub
```

| Directive | Where it applies |
| --- | --- |
| `'@prettyvba-ignore` | The line the comment ends. A line continued with ` _` is one line. |
| `'@prettyvba-ignore-next-line` | The next line that is neither blank nor another directive, so directives for different rules can be stacked above one line. |
| `'@prettyvba-ignore-file` | The whole module. It must come before the first line of code. |
| `'@prettyvba-ignore-start` ... `'@prettyvba-ignore-end` | Every line between them, both included. Regions nest. |

After the directive, a colon and a comma-separated list of rule names
limits it to those rules; without a list, or with `all`, it applies to every
rule. Text after `--` is a reason, for the reader:

```vba
table(0) = Array("id",   "name")   '@prettyvba-ignore: spacing -- columns line up
```

Only apostrophe comments are directives: a `Rem` comment or a `'''`
documentation comment never is. Letter case does not matter in the
directive, and the syntax matches pyVBAanalysis's `'@pyvba-ignore` in a
namespace of its own, so each tool leaves the other's directives alone.

## Mistakes are reported

A directive that cannot be read is reported under the
`suppression-directive` rule, which no directive can switch off:

- an unknown directive (`'@prettyvba-ignore-line`) or rule name, which
  suppresses nothing (a list that names known rules too still applies to
  those);
- `-file` below the first line of code, which suppresses nothing;
- `-end` with no `-start`, which closes nothing;
- `-start` with no `-end`, which suppresses to the end of the module;
- `-next-line` with no line after it.

The `none` preset runs this check alone, so `pyprettyvba check --preset none`
lists the broken directives of a project and nothing else.
