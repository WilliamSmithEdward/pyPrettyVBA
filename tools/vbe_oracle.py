"""Record how the VBE itself renders and compiles VBA, as test evidence.

The formatter's canonical rules (keyword casing, the spaces the VBE inserts,
literal normalization) claim to reproduce what the VBE does to a line. This
tool asks the VBE: it injects probe modules into a real Excel through
pyVBAharness (``CodeModule.AddFromString``), exports them again (VBIDE
``Export``), and records each probe's text before and after. A second phase
runs compile checks for the grammar questions a rule's safety depends on,
such as whether ``If x Then:`` opens a block.

The recording is written to ``tests/oracle/vbe_rendering.json`` and replayed
by the unit suite, so the evidence is checked on every run without Office.
Re-record only on purpose, and review the diff:

    python tools/vbe_oracle.py record

Needs Windows, desktop Excel with "Trust access to the VBA project object
model" enabled, and pyvbaharness installed.
"""

from __future__ import annotations

import argparse
import datetime
import json
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DEFAULT_OUT = REPO / "tests" / "oracle" / "vbe_rendering.json"

MARKER_RE = re.compile(r"^'@@case (\S+)\s*$")


@dataclass(frozen=True)
class Probe:
    """One rendering probe.

    ``scope`` is where the lines go: ``module`` (the declarations section),
    ``body`` (inside a generated procedure of their own), or ``proc`` (the
    lines are a whole procedure).
    """

    id: str
    scope: str
    lines: tuple[str, ...]


@dataclass(frozen=True)
class CompileProbe:
    """One compile check: a whole module body and the question it answers."""

    id: str
    question: str
    source: str


def _p(id: str, scope: str, *lines: str) -> Probe:
    return Probe(id, scope, lines)


# Identifiers carry a "zq" prefix so the project-wide name table the VBE
# keeps (one casing per name) cannot collide with the harness's own modules.
RENDER_PROBES: tuple[Probe, ...] = (
    # --- declarations section -------------------------------------------
    _p("m-option-compare", "module", "option compare text"),
    _p("m-option-base", "module", "option base 1"),
    _p("m-option-private", "module", "option private module"),
    _p("m-declare-alias", "module",
       'private declare ptrsafe function zqgettick lib "kernel32" alias "GetTickCount" () as long'),
    _p("m-declare-noalias", "module",
       'private declare ptrsafe function GetTickCount lib "kernel32" () as long'),
    _p("m-declare-sub", "module",
       'private declare ptrsafe sub zqsleep lib "kernel32" alias "Sleep" (byval ms as long)'),
    _p("m-type-block", "module",
       "private type zqtpoint", "x as long", "    y    as long", "end type"),
    _p("m-enum-block", "module",
       "public enum zqecolor", "zqred=1", "zqgreen   =   2", "zqblue", "end enum"),
    _p("m-dim-list", "module", "private zqmvalue as long,zqmother as string"),
    _p("m-dim-aligned", "module", "dim zqa1     as long", "dim zqlongername as string"),
    _p("m-const", "module", "const zqkmax as long=10"),
    _p("m-const-typed", "module", 'public const zqkname$="abc"'),
    _p("m-public-array", "module", "public zqpub(1 to 3) as long"),
    _p("m-global", "module", "global zqglob as long"),
    _p("m-deftype", "module", "defint i-j"),
    _p("m-directives", "module",
       "#const zqdebug=1",
       "#if zqdebug then",
       "private const zqk1 as long = 1",
       "#elseif vba7 then",
       "private const zqk1 as long = 2",
       "#else",
       "private const zqk1 as long = 3",
       "#end if",
       "    #if win64 then",
       "    #endif"),
    _p("m-comment", "module", "'module comment", "   ' indented module comment"),
    _p("m-blank-lines", "module", "dim zqb1 as long", "", "", "", "dim zqb2 as long"),
    # --- spacing inside statements ---------------------------------------
    _p("p-assign-nospace", "body", "zqx=1+2"),
    _p("p-assign-multispace", "body", "zqx    =    3"),
    _p("p-arith", "body", "zqx=zqx*2-zqx/2"),
    _p("p-intdiv", "body", "zqx = zqx\\2"),
    _p("p-mod", "body", "zqx = zqx mod 2"),
    _p("p-pow", "body", "zqx = zqx^2"),
    _p("p-neg", "body", "zqx=-zqx"),
    _p("p-neg-spaced", "body", "zqx = - zqx"),
    _p("p-negneg", "body", "zqx=zqx- -1"),
    _p("p-paren", "body", "zqx=(zqx+1)*2"),
    _p("p-paren-inner-space", "body", "zqx = ( zqx + 1 ) * 2"),
    _p("p-concat", "body", 'zqs="a"&"b"'),
    _p("p-concat-glued-vars", "body", "zqs=zqs&zqt"),
    _p("p-concat-glued-string", "body", 'zqs=zqs&"x"'),
    _p("p-concat-spaced", "body", "zqs = zqs & zqt"),
    _p("p-typed-assign", "body", "zqn&=5"),
    _p("p-typed-concat", "body", 'zqs = zqt$&"x"'),
    _p("p-array-call", "body", "zqv=array(1,2,3)"),
    _p("p-array-spaces", "body", "zqv = Array( 1 , 2 , 3 )"),
    _p("p-cmp", "body", "if zqx<>1 then zqx=2"),
    _p("p-cmp-logic", "body", "if zqx<=1 and zqx>=0 or not zqx=5 then zqx=3"),
    _p("p-if-else-single", "body", "if zqx = 1 then zqx = 2 else zqx = 3"),
    _p("p-if-colon-single", "body", "if zqx=1 then zqx=2:zqy=3 else zqx=4"),
    _p("p-print-list", "body", "debug.print zqx;zqy,zqs"),
    _p("p-print-list-spaced", "body", "debug.print zqx ; zqy , zqs"),
    _p("p-print-neg", "body", "debug.print -1"),
    _p("p-print-paren", "body", "debug.print(zqx)"),
    _p("p-call-kw", "body", "call zqfoo(zqx,zqy)"),
    _p("p-call-kw-spaces", "body", "call zqfoo( zqx , zqy )"),
    _p("p-call-nokw", "body", "zqfoo zqx,zqy"),
    _p("p-call-nokw-spaces", "body", "zqfoo  zqx,  zqy"),
    _p("p-call-neg-arg", "body", "zqfoo -1, 2"),
    _p("p-call-minus-spaced", "body", "zqfoo - 1"),
    _p("p-call-paren-arg", "body", "zqfoo (zqx), zqy"),
    _p("p-call-paren-single", "body", "zqbar(zqx)"),
    _p("p-func-call", "body", "zqx = zqfunc(zqx,zqy)"),
    _p("p-func-call-space", "body", "zqx = zqfunc (zqx, zqy)"),
    _p("p-index-space", "body", "zqx = zqarr (1)"),
    _p("p-named-args", "body", 'msgbox prompt:="hi",buttons:=vbokonly'),
    _p("p-named-args-spaced", "body", 'msgbox prompt := "hi"'),
    _p("p-msgbox", "body", 'msgbox "hi",vbokonly,"t"'),
    _p("p-set-nothing", "body", "set zqo=nothing"),
    _p("p-set-new", "body", "set zqo = new collection"),
    _p("p-member-call", "body", 'zqo.add 1,"k"'),
    _p("p-member-chain", "body", "zqx = zqo.item(1)"),
    _p("p-member-space-before-dot", "body", "zqx = zqo .item(1)"),
    _p("p-bang", "body", "zqx = zqo!zqkey"),
    _p("p-bang-spaced", "body", "zqx = zqo ! zqkey"),
    _p("p-redim", "body", "redim preserve zqarr(1 to 20)"),
    _p("p-erase", "body", "erase zqarr"),
    _p("p-for-step", "body", "for zqi=1 to 10 step -1", "next zqi"),
    _p("p-for-step-glued", "body", "for zqi = 10 to 1 step-1", "next"),
    _p("p-for-each", "body", "for each zqv in zqo", "next zqv"),
    _p("p-next-multi", "body", "for zqi = 1 to 2", "for zqj = 1 to 2", "next zqj,zqi"),
    _p("p-do-while", "body", "do while zqx<5", "zqx=zqx+1", "loop"),
    _p("p-do-until", "body", "do", "loop until zqx>5"),
    _p("p-while-wend", "body", "while zqx<10", "wend"),
    _p("p-select", "body",
       "select case zqx", "case 1,2", "case 3 to 5", "case is>10", "case else",
       "end select"),
    _p("p-select-colon", "body", "select case zqx", "case 1: zqy = 1", "end select"),
    _p("p-with", "body", "with zqo", ".add 1", "zqfoo .count", "end with"),
    _p("p-on-error", "body", "on error goto zqerr", "on error resume next", "on error goto 0"),
    _p("p-label-indented", "body", "zqx = 1", "    zqlbl:", "zqx = 2"),
    _p("p-label-stmt", "body", "zqlbl2: zqx = 1"),
    _p("p-line-number", "body", "   10 zqx = 1", "20: zqx = 2", "30    zqx = 3"),
    _p("p-goto", "body", "goto zqlbl3", "zqlbl3:"),
    _p("p-gosub", "body", "gosub zqlbl4", "exit sub", "zqlbl4:", "return"),
    _p("p-stop-end", "body", "stop", "end"),
    _p("p-typeof", "body", "if typeof zqo is collection then zqx = 1"),
    _p("p-is-nothing", "body", "if not zqo is nothing then zqx = 1"),
    _p("p-addressof", "body", "zqfoo addressof zqfoo"),
    _p("p-iif", "body", "zqx = iif(zqx>1,1,2)"),
    _p("p-not-paren", "body", "zqb = not(zqb)"),
    _p("p-string-fns", "body", "zqs = ucase$(left$(zqs,2))"),
    _p("p-mid-stmt", "body", 'mid$(zqs,1,1)="x"'),
    _p("p-lset", "body", 'lset zqs = "a"'),
    _p("p-file-open", "body", 'open "c:\\x.txt" for input as #1'),
    _p("p-file-open-access", "body", 'open "c:\\x.txt" for binary access read write lock read as #2'),
    _p("p-file-open-random", "body", 'open "c:\\x.txt" for random as 3 len=10'),
    _p("p-file-line-input", "body", "line input #1,zqs"),
    _p("p-file-print", "body", 'print #1,"a";"b"'),
    _p("p-file-write", "body", "write #1,1,2"),
    _p("p-file-get", "body", "get #1,,zqs"),
    _p("p-file-put", "body", "put #1,1,zqs"),
    _p("p-file-close", "body", "close #1,#2"),
    _p("p-file-close-all", "body", "close"),
    _p("p-file-seek", "body", "seek #1,1"),
    _p("p-file-input", "body", "input #1,zqx,zqy"),
    _p("p-file-lock", "body", "lock #1", "unlock #1"),
    _p("p-error-fn", "body", "zqs = error(5)", "zqs = error$(5)"),
    _p("p-err-raise", "body", 'err.raise 5,"src","desc"'),
    _p("p-let", "body", "let zqx = 1"),
    _p("p-rem", "body", "rem remark here"),
    _p("p-rem-after-colon", "body", "zqx = 1: rem after"),
    _p("p-comment-nospace", "body", "zqx = 1'no space"),
    _p("p-comment-spaced", "body", "zqx = 1      ' spaced"),
    _p("p-comment-tab", "body", "zqx = 1\t' tab before"),
    _p("p-trailing-spaces", "body", "zqx = 1   "),
    _p("p-leading-tab", "body", "\tzqx = 1"),
    _p("p-internal-tab", "body", "zqx\t=\t1"),
    _p("p-continuation", "body", "zqx = 1 + _", "    2"),
    _p("p-continuation-flush", "body", "zqx = 1 + _", "2"),
    _p("p-continuation-multi", "body", 'zqs = "a" & _', '      "b" & _', '      "c"'),
    _p("p-continuation-2spaces", "body", "zqx = 1 +  _", "2"),
    _p("p-continuation-glued", "body", "zqfoo zqx, _", "zqy"),
    # --- literals -------------------------------------------------------
    _p("n-float-point-zero", "body", "zqd = 1.0"),
    _p("n-float-trailing-zero", "body", "zqd = 1.50"),
    _p("n-float-leading-dot", "body", "zqd = .5"),
    _p("n-float-exp", "body", "zqd = 1e3"),
    _p("n-float-exp-sign", "body", "zqd = 1.5e+3"),
    _p("n-float-exp-neg", "body", "zqd = 1.5E-3"),
    _p("n-float-d-exp", "body", "zqd = 1d2"),
    _p("n-float-d-exp-sign", "body", "zqd = 1D+2"),
    _p("n-float-big", "body", "zqd = 100000000000000000000"),
    _p("n-float-big2", "body", "zqd = 123456789012345678"),
    _p("n-float-15", "body", "zqd = 123456789012345"),
    _p("n-float-16", "body", "zqd = 1234567890123456"),
    _p("n-float-e15", "body", "zqd = 1E15"),
    _p("n-float-e16", "body", "zqd = 1E16"),
    _p("n-float-pi15", "body", "zqd = 3.14159265358979"),
    _p("n-float-pi17", "body", "zqd = 3.141592653589793"),
    _p("n-float-point-one", "body", "zqd = 0.1"),
    _p("n-float-hundred", "body", "zqd = 100.0"),
    _p("n-float-zero", "body", "zqd = 0.0"),
    _p("n-float-tiny", "body", "zqd = 0.000001"),
    _p("n-float-tinier", "body", "zqd = 0.0000001"),
    _p("n-float-e-20", "body", "zqd = 1E-20"),
    _p("n-float-e300", "body", "zqd = 1.23456789012345E+300"),
    _p("n-float-dot-only", "body", "zqd = 1."),
    _p("n-float-exp-zero", "body", "zqd = 5E+0"),
    _p("n-suffix-double", "body", "zqd = 0.1#", "zqd = 1#", "zqd = 12345678901234567890#"),
    _p("n-suffix-single", "body", "zqd = 5!", "zqd = 1.0!", "zqd = 1.5!", "zqd = 1.5E3!", "zqd = 16777217!"),
    _p("n-suffix-currency", "body", "zqd = 2.5@", "zqd = 0.5@", "zqd = 922337203685477.5807@"),
    _p("n-hex", "body", "zqd = &hff", "zqd = &HFF&", "zqd = &h00ff", "zqd = &h0", "zqd = &hFFFF",
       "zqd = &hFFFFFFFF", "zqd = &h7FFFFFFF", "zqd = &H8000", "zqd = &H8000&"),
    _p("n-octal", "body", "zqd = &o17", "zqd = &O17&", "zqd = &O100000", "zqd = &17"),
    _p("n-integer", "body", "zqd = 32768", "zqd = 2147483648", "zqd = 10&", "zqd = 10%",
       "zqd = 00012", "zqd = 32767%"),
    _p("n-negative", "body", "zqd = -1", "zqd = - 1", "zqd = -1.0", "zqd = -0.0"),
    _p("d-iso", "body", "zqdt = #2020-01-15#"),
    _p("d-us", "body", "zqdt = #1/15/2020#"),
    _p("d-swapped", "body", "zqdt = #15/1/2020#"),
    _p("d-datetime", "body", "zqdt = #1/1/2000 10:30#"),
    _p("d-time", "body", "zqdt = #10:30:00#"),
    _p("d-time-pm", "body", "zqdt = #10:30 PM#"),
    _p("d-month-name", "body", "zqdt = #January 15, 2020#"),
    _p("d-slash-ymd", "body", "zqdt = #2020/1/15#"),
    _p("d-midnight", "body", "zqdt = #12:00:00 AM#"),
    _p("d-abbrev", "body", "zqdt = #1-Jan-2020#"),
    _p("d-padded", "body", "zqdt = #  1/1/2000  #"),
    _p("d-two-digit-year", "body", "zqdt = #1/1/99#"),
    _p("d-zero-time", "body", "zqdt = #00:00#"),
    _p("d-full", "body", "zqdt = #1/1/2000 00:00:00#", "zqdt = #12/31/2020 23:59:59#"),
    # --- keywords and casing --------------------------------------------
    _p("k-types", "body",
       "dim zqa as integer, zqb as boolean, zqc as variant, zqdd as currency, zqe as byte",
       "dim zqf as single, zqg as double, zqh as date, zqii as string, zqj as object",
       "dim zqk as longptr, zql as longlong"),
    _p("k-static-local", "body", "static zqst as long"),
    _p("k-const-local", "body", 'const zqlocal as string = "x"'),
    _p("k-dim-new", "body", "dim zqc2 as new collection"),
    _p("k-fixed-string", "body", "dim zqfs as string * 10"),
    _p("k-dim-arrays", "body", "dim zqarr2() as long, zqarr3(0 to 2, 1 to 3) as string"),
    _p("k-literals", "body", "zqb = true: zqb = false: zqv = empty: zqv = null: set zqo = nothing"),
    _p("k-member-keywords", "body", "zqs = zqo.text", "zqs = zqo.name", "zqs = zqo.value",
       "zqs = zqo.Print", "zqs = zqo.input"),
    _p("k-casing-user", "body", "dim zqmyvar as long", "ZQMYVAR = 1", "ZqMyVar = 2"),
    _p("k-casing-builtin", "body", 'msgbox "x"', 'MSGBOX "y"', "zqs = ucase$(zqs)", "zqs = vbcrlf",
       "zqs = vbCRLF", "zqs = VBA.left$(zqs, 1)", "debug.print now"),
    _p("k-endif", "body", "if zqx then", "zqx = 1", "endif"),
    _p("k-end-if-spaces", "body", "if zqx then", "zqx = 1", "end   if"),
    _p("k-elseif-inline", "body", "if zqx then", "zqx = 1", "elseif zqy then zqx = 2",
       "else zqx = 3", "end if"),
    _p("k-if-then-colon", "body", "if zqx then:", "zqx = 1", "end if"),
    _p("k-if-then-rem", "body", "if zqx then rem hello"),
    _p("k-comment-continued", "body", "' comment _", "continued text here"),
    _p("k-exit", "body", "exit for", "exit do"),
    _p("k-debug-assert", "body", "debug.assert zqx > 0"),
    _p("k-doevents", "body", "doevents", "randomize", "randomize 5"),
    _p("k-date-fns", "body", "zqd = date", "zqs = date$", "zqd = now()"),
    # --- which gaps survive -----------------------------------------------
    _p("g-pow-literals", "body", "zqx = 2 ^ 3", "zqx = (zqx)^2", "zqx = zqx ^ 2", "zqx = zqx ^2"),
    _p("g-as-const", "body", "const zqc1    as long = 1", "const zqc22     = 5"),
    _p("g-as-after-comma", "body", "dim zqg1 as long,    zqg2     as long"),
    _p("g-as-static", "body", "static zqg3     as long"),
    _p("g-as-redim", "body", "redim zqg4(1 to 2)     as long"),
    _p("g-dim-gap", "body", "dim    zqg5 as long"),
    _p("g-colon-gaps", "body", "zqx = 1   : zqy = 2", "zqx = 1:    zqy = 2"),
    _p("g-label-gap", "body", "    zqlbl5:   zqx = 1"),
    _p("g-then-gap", "body", "if zqx then     zqy = 1"),
    _p("g-keyword-gaps", "body", "end    select", "exit    sub"),
    _p("g-tab-comment", "body", "zqlongname = 1\t' after tab", "\t' tab indented comment"),
    _p("g-whitespace-only", "body", "zqx = 1", "      ", "zqx = 2"),
    _p("g-print-trailing-semicolon", "body", 'debug.print "a";', "debug.print spc(5);\"a\";tab(3);\"b\""),
    _p("g-unary-after-op", "body", "zqx = 1 +-2", "zqx = zqa -zqb", "zqx = zqa*-zqb"),
    _p("g-call-kw-space-paren", "body", "call zqfoo (1)"),
    _p("g-plus-strings", "body", 'zqs = "a"+"b"'),
    _p("g-like-xor", "body", 'if zqs like "a*" xor zqb imp zqc eqv zqd then zqx = 1'),
    _p("g-hex-suffix-expr", "body", "zqx = &hff&+1"),
    _p("g-date-expr", "body", "zqdt = #1/1/2000#+1"),
    _p("g-case-is-gap", "body", "select case zqx", "case is >10", "case is=5", "end select"),
    _p("g-with-gaps", "body", "with   zqo", "  .add    1", "end with"),
    _p("g-directive-gaps", "module", "#if    zqdebug    then", "#end if"),
    _p("g-line-number-indented", "body", "    40 zqx = 4", "50 zqx = 5"),
    _p("g-negative-literal-arg", "body", "zqfoo zqx, -1", "zqfoo(-1)"),
    _p("g-nested-parens", "body", "zqx = ((zqa + zqb) * (zqc - zqd))"),
    _p("g-string-concat-continued", "body", 'zqs = "a" & _', '"b"'),
    _p("g-casing-first-use", "proc", "sub zqusefirst()", "ZqLater = 1", "end sub",
       "sub zqdeclarelater()", "dim zqlater as long", "end sub"),
    _p("g-casing-module-var", "proc", "sub zqusemodvar()", "ZQMODVAR = 1", "end sub"),
    _p("g-param-as", "proc", "sub zqparams(byval zqp1     as long, zqp2 as string)", "end sub"),
    _p("g-function-as", "proc", "function zqret()     as long", "end function"),
    _p("g-declare-param-as", "module",
       'private declare ptrsafe function zqbeep lib "kernel32" alias "Beep" (byval f     as long, byval d as long) as long'),
    _p("g-module-var-as", "module", "private zqmv1     as long", "public zqmv2      as string"),
    _p("g-type-member-array", "module", "private type zqtarr", "    items(1 to 3)    as long", "end type"),
    _p("g-enum-aligned", "module", "private enum zqealign", "    zqfirst   = 1", "    zqsecondx = 2", "end enum"),
    _p("g-module-casing-decl", "module", "private zqmodvar as long"),
    # --- round three: the questions implementation raised -----------------
    _p("r-call-omitted-first-spaced", "body", "zqfoo , 2"),
    _p("r-call-omitted-first-glued", "body", "zqfoo, 2"),
    _p("r-float-small", "body", "zqd = 1E-8", "zqd = 1E-10", "zqd = 1E-14", "zqd = 1E-15",
       "zqd = 1E-16", "zqd = 1E-19", "zqd = 1.5E-10", "zqd = 1.23E-5", "zqd = 0.000123456789012345",
       "zqd = 0.00000001", "zqd = 1.2345E-8"),
    _p("r-float-small-single", "body", "zqd = 0.00001!", "zqd = 1E-10!", "zqd = 0.001!", "zqd = 0.0001!"),
    _p("r-integer-suffixes", "body", "zqd = 32768&", "zqd = &HFF%", "zqd = &HFFFF&", "zqd = 70000&",
       "zqd = 10^", "zqd = 40000^"),
    _p("r-comment-after-collapse", "body", "zqx    =    1 ' c1", "zqx    =    1    ' c2"),
    _p("r-colon-after-collapse", "body", "zqx    =    1:    zqy = 2"),
    _p("r-ge-variants", "body", "if zqa => zqb then zqx = 1", "if zqa =< zqb then zqx = 1",
       "if zqa >< zqb then zqx = 1"),
    _p("r-undeclared-casing", "proc", "sub zqundecl1()", "QqUndecl = 1", "end sub", "sub zqundecl2()",
       "qqundecl = 2", "end sub"),
    _p("r-compiler-consts", "module", "#if vba6 or win16 or win32 or mac or twinbasic then", "#end if"),
    _p("r-single-if-call-paren", "body", "if zqx then zqfoo(1)"),
    _p("r-label-comment", "body", "zqlbl6: ' c", "zqlbl7: rem c"),
    _p("r-print-trailing", "body", "print #1, zqx; zqy;", "debug.print zqx,"),
    _p("r-bang-expr", "body", "zqx = zqo!zqk + 1"),
    _p("r-named-arg-spaced", "body", "zqfoo zqx:=1, zqy:= 2"),
    _p("r-call-empty-parens", "body", "call zqfoo()", "zqfoo()"),
    _p("r-spc-glued", "body", 'debug.print spc(2);"x"'),
    _p("r-literal-comma", "body", "zqfoo 1 , 2"),
    _p("r-mod-paren", "body", "zqx = zqy mod(2)"),
    _p("r-dim-gaps-collapse", "body", "dim    zqg6    as long", "dim zqg7 as long,zqg8     as long"),
    _p("r-type-member-array", "module", "private type zqt3", "zqarr() as byte", "end type"),
    _p("r-function-returns-array", "proc", "function zqbytes() as byte()", "end function"),
    _p("r-with-paren", "body", "with zqo", "zqfoo (.count)", ".item (1) = 2", "end with"),
    _p("r-negative-paren", "body", "zqx = -(zqy)", "if not(zqb) then zqx = 1"),
    _p("r-glued-names", "body", "zqs = zqa+zqb", "zqx = zqa ^zqb"),
    _p("r-line-number-long", "body", "12345 zqx = 1"),
    _p("r-line-input-casing", "body", "LINE INPUT #1, zqs"),
    _p("r-trim-space-paren", "body", "zqs = trim$ (zqs)"),
    _p("r-member-call-paren", "body", "zqo.add (1)", "set zqo = zqfunc (1)"),
    _p("r-deftype-multi", "module", "deflng a-c, x-z"),
    # --- whole procedures -----------------------------------------------
    _p("s-static-sub", "proc", "public static sub zqstat()", "end sub"),
    _p("s-typed-function", "proc", "private function zqfn$(byval s as string)", "zqfn$ = s",
       "end function"),
    _p("s-optional-params", "proc",
       "sub zqopt(optional byval a as long = 5, optional b as variant, paramarray c() as variant)",
       "end sub"),
    _p("s-one-line", "proc", "sub zqone(): zqx1 = 1: end sub"),
    _p("s-end-comment", "proc", "sub zqtail()", "end sub ' tail comment"),
    _p("s-friend-in-std", "proc", "private sub zqpriv(byref a as long, byval b as string)", "end sub"),
)

CLASS_PROBES: tuple[Probe, ...] = (
    _p("c-event", "module", "public event zqchanged(byval x as long)"),
    _p("c-member", "module", "private mzqval as long"),
    _p("c-property-get", "proc", "public property get zqval() as long", "zqval = mzqval",
       "end property"),
    _p("c-property-let", "proc", "public property let zqval(byval v as long)", "mzqval = v",
       "raiseevent zqchanged(v)", "end property"),
    _p("c-friend", "proc", "friend sub zqf()", "set me.zqobj = nothing", "end sub"),
    _p("c-initialize", "proc", "private sub class_initialize()", "mzqval = 0", "end sub"),
)

COMPILE_PROBES: tuple[CompileProbe, ...] = (
    CompileProbe(
        "if-then-colon-opens-block",
        "Does `If x Then:` with nothing after the colon open a block If?",
        "Sub ZqC()\n    Dim zqx As Boolean\n    If zqx Then:\n        zqx = False\n    End If\nEnd Sub\n"),
    CompileProbe(
        "if-then-colon-statement-single",
        "Is `If x Then: stmt` a single-line If (no End If needed)?",
        "Sub ZqC()\n    Dim zqx As Boolean\n    If zqx Then: zqx = False\nEnd Sub\n"),
    CompileProbe(
        "comment-continues",
        "Does ` _` at the end of a comment continue the comment onto the next line?",
        "Sub ZqC()\n    ' a comment that continues _\n    this line is not valid code !!!\nEnd Sub\n"),
    CompileProbe(
        "comment-not-continued-control",
        "Control: the same invalid line without a continued comment is rejected.",
        "Sub ZqC()\n    ' a comment that ends here\n    this line is not valid code !!!\nEnd Sub\n"),
    CompileProbe(
        "continuation-trailing-space",
        "Is ` _` followed by trailing spaces still a line continuation?",
        "Sub ZqC()\n    Dim zqx As Long\n    zqx = 1 + _   \n        2\nEnd Sub\n"),
    CompileProbe(
        "rem-after-then",
        "Is `If x Then Rem text` accepted?",
        "Sub ZqC()\n    Dim zqx As Boolean\n    If zqx Then Rem hello world\nEnd Sub\n"),
    CompileProbe(
        "elseif-inline-statement",
        "Is `ElseIf c Then stmt` on one line accepted inside a block If?",
        "Sub ZqC()\n    Dim zqx As Boolean, zqy As Long\n    If zqx Then\n        zqy = 1\n"
        "    ElseIf zqy = 2 Then zqy = 3\n    End If\nEnd Sub\n"),
    CompileProbe(
        "else-inline-statement",
        "Is `Else stmt` on one line accepted inside a block If?",
        "Sub ZqC()\n    Dim zqx As Boolean, zqy As Long\n    If zqx Then\n        zqy = 1\n"
        "    Else zqy = 3\n    End If\nEnd Sub\n"),
    CompileProbe(
        "one-line-procedure",
        "Is a whole procedure on one line (colon separated) accepted?",
        "Sub ZqC(): Dim zqx As Long: zqx = 1: End Sub\n"),
    CompileProbe(
        "concat-glued-identifiers",
        "Is `s = s&t` (no spaces around &) accepted for two String variables?",
        "Sub ZqC()\n    Dim zqs As String, zqt As String\n    zqs = zqs&zqt\nEnd Sub\n"),
    CompileProbe(
        "endif-single-word",
        "Is `EndIf` accepted as `End If`?",
        "Sub ZqC()\n    Dim zqx As Boolean\n    If zqx Then\n        zqx = False\n    EndIf\nEnd Sub\n"),
    CompileProbe(
        "indented-label",
        "Is an indented label accepted?",
        "Sub ZqC()\n    Dim zqx As Long\n    GoTo ZqLbl\n        ZqLbl:\n    zqx = 1\nEnd Sub\n"),
    CompileProbe(
        "line-number-colon",
        "Is a line number followed by a colon accepted?",
        "Sub ZqC()\n    Dim zqx As Long\n10: zqx = 1\n20 zqx = 2\nEnd Sub\n"),
    CompileProbe(
        "trailing-static",
        "Is `Sub Name() Static` accepted?",
        "Sub ZqC() Static\n    Dim zqx As Long\n    zqx = zqx + 1\nEnd Sub\n"),
    CompileProbe(
        "longlong-suffix",
        "Is the `^` LongLong type suffix accepted on this Office?",
        "Sub ZqC()\n    Dim zqll As LongLong\n    zqll = 10^\nEnd Sub\n"),
    CompileProbe(
        "bang-with-spaces",
        "Is `obj ! key` (spaces around the bang) accepted?",
        "Sub ZqC()\n    Dim zqo As New Collection, zqv As Variant\n    zqv = zqo ! zqkey\nEnd Sub\n"),
    CompileProbe(
        "bang-glued-control",
        "Control: `obj!key` is accepted.",
        "Sub ZqC()\n    Dim zqo As New Collection, zqv As Variant\n    zqv = zqo!zqkey\nEnd Sub\n"),
    CompileProbe(
        "with-member-argument",
        "Is `Foo .Member` inside a With block a call with a with-member argument?",
        "Sub ZqC()\n    Dim zqo As New Collection\n    With zqo\n        ZqTake .Count\n"
        "    End With\nEnd Sub\nSub ZqTake(ByVal n As Long)\nEnd Sub\n"),
    CompileProbe(
        "float-trailing-dot",
        "Is `1.` (no fractional digits) accepted as a literal?",
        "Sub ZqC()\n    Dim zqd As Double\n    zqd = 1.\nEnd Sub\n"),
    CompileProbe(
        "label-then-statement-same-line",
        "Is `Label: statement` accepted?",
        "Sub ZqC()\n    Dim zqx As Long\nZqLbl: zqx = 1\n    GoTo ZqLbl\nEnd Sub\n"),
    CompileProbe(
        "rem-continues",
        "Does a Rem comment ending in ` _` continue onto the next line?",
        "Sub ZqC()\n    Rem a remark that continues _\n    this line is not valid code !!!\nEnd Sub\n"),
    CompileProbe(
        "directive-comment-continues",
        "Does a comment on a #If line continue with ` _`?",
        "#If True Then ' directive comment _\n    this line is not valid code !!!\n#End If\n"
        "Sub ZqC()\nEnd Sub\n"),
    CompileProbe(
        "type-member-indented-spaces",
        "Control: a Type with extra spaces inside member lines is accepted.",
        "Private Type ZqT\n    a    As Long\nEnd Type\nSub ZqC()\nEnd Sub\n"),
    CompileProbe(
        "call-omitted-first-spaced",
        "Is `Foo , 2` a call with the first argument omitted (rejected when it is required)?",
        "Private Sub ZqNeeds(ByVal x As Long, ByVal y As Long)\nEnd Sub\n"
        "Sub ZqC()\n    ZqNeeds , 2\nEnd Sub\n"),
    CompileProbe(
        "call-omitted-first-glued",
        "Is `Foo, 2` the same call (rejected the same way when the argument is required)?",
        "Private Sub ZqNeeds(ByVal x As Long, ByVal y As Long)\nEnd Sub\n"
        "Sub ZqC()\n    ZqNeeds, 2\nEnd Sub\n"),
    CompileProbe(
        "greater-equal-reversed",
        "Is `=>` accepted as a comparison operator?",
        "Sub ZqC()\n    Dim zqa As Long, zqb As Long\n    If zqa => zqb Then zqa = 1\nEnd Sub\n"),
)


@dataclass(frozen=True)
class IsolatedProbe:
    """A whole module, rendered in a workbook of its own.

    The VBE keeps one spelling per name for a whole project, so a probe that
    asks how names interact gets a fresh project, where no other probe (and
    no zq prefix) can take part.
    """

    id: str
    question: str
    lines: tuple[str, ...]


ISOLATED_PROBES: tuple[IsolatedProbe, ...] = (
    IsolatedProbe(
        "i-contextual-keyword-as-name",
        "Do variables named text and binary keep their own spelling, and do they change "
        "the Text of Option Compare Text?",
        ("option compare text", "sub zqctx()", "dim text as string, binary as long", 'text = "a"',
         "binary = 1", "debug.print TEXT; BINARY", "end sub"),
    ),
    IsolatedProbe(
        "i-parameter-named-text",
        "Does a parameter named text keep its spelling in a module that uses nothing else "
        "called text?",
        ("function zqlen(byval text as string) as long", "zqlen = len(TEXT)", "end function"),
    ),
    IsolatedProbe(
        "i-type-member-named-like-library-function",
        "Type members named second and left: do they keep their spelling, and do the library "
        "functions Second and Left elsewhere in the module take it?",
        ("private type zqpair", "second as long", "left as long", "end type", "sub zqtm()",
         "dim n as long, s as string", "n = Second(now)", 's = Left("ab", 1)', "end sub"),
    ),
    IsolatedProbe(
        "i-two-declarations",
        "Two procedures declare one local name with different casing: which spelling wins?",
        ("sub zqr1()", "dim zqCount as long", "zqcount = 1", "end sub", "sub zqr2()",
         "dim ZQCOUNT as long", "ZqCount = 2", "end sub"),
    ),
    IsolatedProbe(
        "i-parameter-named-like-library-function",
        "A parameter named LEFT: which spelling do it and a call of the Left function in "
        "another procedure get?",
        ("sub zqpl(byval LEFT as long)", "debug.print left", "end sub", "sub zqpl2()",
         'debug.print Left("ab", 1)', "end sub"),
    ),
    # --- round five: the library name table, and blank lines at the top ---
    IsolatedProbe(
        "i-library-names",
        "Undeclared names, members and named arguments that Excel's or Office's libraries "
        "define: which spelling does each get, including names two libraries spell differently?",
        ("sub zqlib()", "debug.print sql, ready, state, pi, total, z",
         "debug.print id, filename, ddb, checkboxes, strikethrough, url, gridlines, xlconstants",
         'activesheet.range("a1").value = activesheet.cells(1, 1).END(xlup).row',
         "activeworkbook.close savechanges:=false",
         'msgbox prompt:="p", title:="t", buttons:=vbokonly', "end sub"),
    ),
    IsolatedProbe(
        "i-leading-blank-lines",
        "Blank lines before the first line of a module: does AddFromString keep them?",
        ("", "", "option explicit", "", "sub zqlead()", "end sub"),
    ),
    # --- round six: library spellings the data and a VBE export disagree on,
    # and a Declare's parameters ------------------------------------------
    IsolatedProbe(
        "i-vba-library-members",
        "How does the VBE spell the members of Err, Debug and Collection, and "
        "Application.Hwnd and .Version?",
        ("sub zqmem()", "debug.print err.lastdllerror, err.helpcontext, err.helpfile, err.source",
         "dim c as new collection", "c.add 1: debug.print c.count, c.item(1)",
         "debug.print application.hwnd, application.version", "debug.assert true", "end sub"),
    ),
    IsolatedProbe(
        "i-declare-parameters",
        "Are a Declare's parameter names declarations: does `HWND` there respell "
        "application.hwnd?",
        ('private declare ptrsafe function zqwintext lib "user32" alias "GetWindowTextA" '
         "(byval HWND as longptr, byval lpString as string, byval cch as long) as long",
         "sub zqdp()", "debug.print application.hwnd", "end sub"),
    ),
)

# Modules written to a file and read in with VBComponents.Import, the way an
# exported module comes back, rather than with AddFromString.
IMPORT_PROBES: tuple[IsolatedProbe, ...] = (
    IsolatedProbe(
        "f-import-leading-blank-lines",
        "Blank lines between the Attribute lines and the first line of code: does Import keep them?",
        ("", "", "option explicit", "", "sub zqlead()", "end sub"),
    ),
    IsolatedProbe(
        "f-import-trailing-blank-lines",
        "Blank lines after the last line of code: does Import keep them?",
        ("option explicit", "sub zqtail()", "end sub", "", ""),
    ),
)

IMPORTER = (
    "Public Function ZqImport(ByVal zqFile As String) As String\r\n"
    "    ZqImport = ThisWorkbook.VBProject.VBComponents.Import(zqFile).Name\r\n"
    "End Function\r\n"
)


@dataclass(frozen=True)
class RuntimeProbe:
    """One runtime check: a Function whose return value answers a question."""

    id: str
    question: str
    source: str
    proc: str


RUNTIME_PROBES: tuple[RuntimeProbe, ...] = (
    RuntimeProbe(
        "comment-continuation-hides-statement",
        "Does a statement on the line after `' comment _` run? (1 = hidden, 2 = ran)",
        "Public Function ZqR() As Long\n    Dim n As Long\n    n = 1\n    ' comment _\n"
        "    n = 2\n    ZqR = n\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "rem-continuation-hides-statement",
        "Does a statement on the line after `Rem comment _` run? (1 = hidden, 2 = ran)",
        "Public Function ZqR() As Long\n    Dim n As Long\n    n = 1\n    Rem comment _\n"
        "    n = 2\n    ZqR = n\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "end-of-line-comment-continuation-hides-statement",
        "After `n = 1 ' note _`, does the next line run? (1 = hidden, 2 = ran)",
        "Public Function ZqR() As Long\n    Dim n As Long\n    n = 1 ' note _\n"
        "    n = 2\n    ZqR = n\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "continuation-with-trailing-space",
        "Does `1 + _   ` continue onto the next line? (3 = continued)",
        "Public Function ZqR() As Long\n    ZqR = 1 + _   \n        2\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "if-then-colon-is-single-line",
        "After `If False Then:`, does the next line run? (5 = ran, so the If was single-line)",
        "Public Function ZqR() As Long\n    Dim n As Long\n    If False Then:\n    n = 5\n"
        "    ZqR = n\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "single-line-if-colon-scope",
        "In `If False Then n = 1: n = 2`, is the second statement conditional? (0 = yes)",
        "Public Function ZqR() As Long\n    Dim n As Long\n    If False Then n = 1: n = 2\n"
        "    ZqR = n\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "single-line-else-colon-scope",
        "In `If True Then n = 1 Else n = 2: n = 3`, is `n = 3` part of the Else? (1 = yes)",
        "Public Function ZqR() As Long\n    Dim n As Long\n    If True Then n = 1 Else n = 2: n = 3\n"
        "    ZqR = n\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "call-omitted-first-glued-runs",
        "Does `Foo, 2` call Foo with the first argument omitted? (102 = yes)",
        "Private zqLast As Long\n"
        "Private Sub ZqTwo(Optional a As Variant, Optional b As Variant)\n"
        "    zqLast = IIf(IsMissing(a), 100, 0) + IIf(IsMissing(b), 10, b)\nEnd Sub\n"
        "Public Function ZqR() As Long\n    ZqTwo, 2\n    ZqR = zqLast\nEnd Function\n",
        "ZqR"),
    RuntimeProbe(
        "call-omitted-first-spaced-runs",
        "Control: `Foo , 2` calls Foo with the first argument omitted. (102 = yes)",
        "Private zqLast As Long\n"
        "Private Sub ZqTwo(Optional a As Variant, Optional b As Variant)\n"
        "    zqLast = IIf(IsMissing(a), 100, 0) + IIf(IsMissing(b), 10, b)\nEnd Sub\n"
        "Public Function ZqR() As Long\n    ZqTwo , 2\n    ZqR = zqLast\nEnd Function\n",
        "ZqR"),
)


def build_module_text(probes: tuple[Probe, ...], proc_prefix: str) -> str:
    """Assemble probes into one module body with a marker before each."""
    head: list[str] = []
    tail: list[str] = []
    for index, probe in enumerate(probes, start=1):
        marker = f"'@@case {probe.id}"
        if probe.scope == "module":
            head.append(marker)
            head.extend(probe.lines)
        elif probe.scope == "body":
            tail.append(marker)
            tail.append(f"Sub {proc_prefix}{index:03d}()")
            tail.extend(probe.lines)
            tail.append("End Sub")
        elif probe.scope == "proc":
            tail.append(marker)
            tail.extend(probe.lines)
        else:
            raise ValueError(f"unknown scope {probe.scope!r}")
    tail.append("'@@case @@end")
    return "\r\n".join(head + tail) + "\r\n"


def split_rendered(text: str, probes: tuple[Probe, ...]) -> dict[str, list[str]]:
    """Cut exported module text back into per-probe line lists."""
    lines = text.splitlines()
    chunks: dict[str, list[str]] = {}
    current: str | None = None
    for line in lines:
        match = MARKER_RE.match(line)
        if match:
            current = match.group(1)
            chunks[current] = []
            continue
        if current is not None:
            chunks[current].append(line)
    by_id = {probe.id: probe for probe in probes}
    result: dict[str, list[str]] = {}
    for probe_id, chunk in chunks.items():
        probe = by_id.get(probe_id)
        if probe is None:
            continue
        while chunk and chunk[-1] == "" and probe.lines[-1] != "":
            chunk.pop()
        if probe.scope == "body":
            # Drop the generated Sub/End Sub wrapper.
            chunk = chunk[1:-1]
        result[probe_id] = chunk
    return result


def record(out: Path) -> dict[str, object]:
    from pyvbaharness import ExcelSession

    rendering: list[dict[str, object]] = []
    compiles: list[dict[str, object]] = []
    runtime: list[dict[str, object]] = []
    with ExcelSession() as excel:
        excel.new_workbook()
        version = str(excel.eval("Application.Version"))
        build = str(excel.eval("Application.Build"))
        std_text = build_module_text(RENDER_PROBES, "ZqP")
        cls_text = build_module_text(CLASS_PROBES, "ZqQ")
        excel.add_module("ZqProbe", std_text, kind="standard")
        excel.add_module("ZqClassProbe", cls_text, kind="class")
        with tempfile.TemporaryDirectory() as tmp:
            files = excel.export_modules(tmp)
            exported = {Path(f).stem: Path(f).read_bytes() for f in files}
        for name, probes in (("ZqProbe", RENDER_PROBES), ("ZqClassProbe", CLASS_PROBES)):
            raw = exported[name]
            text = raw.decode("cp1252")
            rendered = split_rendered(text, probes)
            for probe in probes:
                output = rendered.get(probe.id)
                rendering.append({
                    "id": probe.id,
                    "module": "class" if name == "ZqClassProbe" else "standard",
                    "scope": probe.scope,
                    "input": list(probe.lines),
                    "output": output,
                })
            # Keep the raw export head so the header the VBE writes is on record.
            rendering.append({
                "id": f"@export-head-{name}",
                "module": "class" if name == "ZqClassProbe" else "standard",
                "scope": "export",
                "input": [],
                "output": text.splitlines()[:12],
            })
        excel.remove_module("ZqProbe")
        excel.remove_module("ZqClassProbe")

        # Runtime checks run before any compile check: showing the VBE makes
        # every later COM call slower, and a rejected compile leaves nothing
        # to run.
        for rprobe in RUNTIME_PROBES:
            result = excel.run_vba(rprobe.source, proc=rprobe.proc, line_numbers=False,
                                   timeout=30)
            runtime.append({
                "id": rprobe.id,
                "question": rprobe.question,
                "source": rprobe.source,
                "outcome": result.outcome,
                "value": result.value,
                "error": None if result.error is None else str(result.error.description),
            })
            print(f"runtime {rprobe.id}: {result.outcome} {result.value!r}", file=sys.stderr)

        for probe in COMPILE_PROBES:
            excel.add_module("ZqCompile", probe.source, kind="standard")
            result = excel.compile_project(watch_seconds=15)
            compiles.append({
                "id": probe.id,
                "question": probe.question,
                "source": probe.source,
                "outcome": result.outcome,
                "message": result.message,
            })
            excel.remove_module("ZqCompile")
            print(f"compile {probe.id}: {result.outcome} {result.message!r}", file=sys.stderr)

        isolated: list[dict[str, object]] = []
        for iprobe in ISOLATED_PROBES:
            excel.new_workbook()
            excel.add_module("ZqIso", "\r\n".join(iprobe.lines) + "\r\n", kind="standard")
            with tempfile.TemporaryDirectory() as tmp:
                files = excel.export_modules(tmp)
                raw = {Path(f).stem: Path(f).read_bytes() for f in files}["ZqIso"]
            lines = [line for line in raw.decode("cp1252").splitlines() if not line.startswith("Attribute VB_")]
            isolated.append({
                "id": iprobe.id,
                "question": iprobe.question,
                "input": list(iprobe.lines),
                "output": lines,
            })
            excel.remove_module("ZqIso")
            print(f"isolated {iprobe.id}: {lines!r}", file=sys.stderr)

        for iprobe in IMPORT_PROBES:
            excel.new_workbook()
            with tempfile.TemporaryDirectory() as tmp:
                source = Path(tmp) / "in" / "ZqImp.bas"
                source.parent.mkdir()
                text = 'Attribute VB_Name = "ZqImp"\r\n' + "\r\n".join(iprobe.lines) + "\r\n"
                source.write_bytes(text.encode("cp1252"))
                run = excel.run_vba(IMPORTER, proc="ZqImport", args=(str(source),), line_numbers=False,
                                    timeout=30)
                files = excel.export_modules(Path(tmp) / "out")
                raw = {Path(f).stem: Path(f).read_bytes() for f in files}.get("ZqImp")
            lines = [] if raw is None else [
                line for line in raw.decode("cp1252").splitlines() if not line.startswith("Attribute VB_")
            ]
            isolated.append({
                "id": iprobe.id,
                "question": iprobe.question,
                "input": list(iprobe.lines),
                "output": lines,
                "method": "VBComponents.Import",
            })
            if raw is not None:
                excel.remove_module("ZqImp")
            print(f"import {iprobe.id}: {run.outcome} {lines!r}", file=sys.stderr)

    evidence: dict[str, object] = {
        "recorded": datetime.date.today().isoformat(),
        "excel_version": version,
        "excel_build": build,
        "method": (
            "Rendering: CodeModule.AddFromString into a new module, then "
            "VBComponent.Export, read as cp1252. Compile: VBE Debug > Compile "
            "through pyVBAharness compile_project."
        ),
        "rendering": rendering,
        "compile": compiles,
        "runtime": runtime,
        "isolated": isolated,
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(evidence, indent=1, ensure_ascii=True) + "\n", encoding="utf-8")
    return evidence


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="command", required=True)
    rec = sub.add_parser("record", help="record rendering and compile evidence from Excel")
    rec.add_argument("--out", type=Path, default=DEFAULT_OUT)
    show = sub.add_parser("show", help="print a recording as input -> output pairs")
    show.add_argument("--file", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args(argv)
    if args.command == "record":
        evidence = record(args.out)
        print(f"wrote {args.out} ({len(evidence['rendering'])} rendering, "  # type: ignore[arg-type]
              f"{len(evidence['compile'])} compile)")  # type: ignore[arg-type]
        return 0
    data = json.loads(args.file.read_text(encoding="utf-8"))
    for case in data["rendering"]:
        print(f"== {case['id']}")
        for line in case["input"]:
            print(f"  in : {line!r}")
        for line in case["output"] or ["<missing>"]:
            print(f"  out: {line!r}")
    for case in data["compile"]:
        print(f"== compile {case['id']}: {case['outcome']} {case['message']!r}")
    for case in data.get("runtime", []):
        print(f"== runtime {case['id']}: {case['outcome']} value={case['value']!r}")
    for case in data.get("isolated", []):
        print(f"== isolated {case['id']}")
        for line in case["input"]:
            print(f"  in : {line!r}")
        for line in case["output"]:
            print(f"  out: {line!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
