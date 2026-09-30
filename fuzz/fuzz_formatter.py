"""Coverage-guided fuzzing of the lexer and the formatter on arbitrary text.

The properties tests/test_properties.py checks with Hypothesis, run here
under Atheris's coverage guidance:

  lexer   tokenize() loses nothing: its tokens join back to the input.
  format  format_source() never raises (SafetyError or
          UnstableFormattingError is a formatter bug, not a refusal), its
          output formats to itself again, it changes no character outside
          ASCII and whitespace, and the "none" preset changes nothing. The
          first input byte picks the preset; the rest is the source, as
          UTF-8.

    python fuzz/fuzz_formatter.py <target> [libFuzzer options] [corpus dirs]
    python fuzz/fuzz_formatter.py format -max_total_time=60 tests/fuzz_corpus/vba

The .github/workflows/fuzz.yml workflow runs each target from
tests/fuzz_corpus/vba. A finding becomes a seed there, which
tests/test_fuzz_corpus.py replays on every CI run.
"""

import sys

import atheris

with atheris.instrument_imports():
    from pyprettyvba import PRESETS, Config, format_source
    from pyprettyvba.chars import is_wsc
    from pyprettyvba.lexer import tokenize

CONFIGS = [Config({"preset": preset}) for preset in PRESETS]
NONE = PRESETS.index("none")


def _non_ascii(text):
    return [c for c in text if ord(c) > 127 and not is_wsc(c)]


def fuzz_lexer(data):
    text = data.decode("utf-8", errors="replace")
    if "".join(token.text for token in tokenize(text)) != text:
        raise AssertionError("the tokens do not join back to the input")


def fuzz_format(data):
    if not data:
        return
    index = data[0] % len(CONFIGS)
    text = data[1:].decode("utf-8", errors="replace")
    config = CONFIGS[index]
    result = format_source(text, config)
    if index == NONE:
        if result.output != text:
            raise AssertionError('the "none" preset changed the text')
        return
    if format_source(result.output, config).output != result.output:
        raise AssertionError(f"{PRESETS[index]}: formatting the output changes it again")
    if _non_ascii(result.output) != _non_ascii(text):
        raise AssertionError(f"{PRESETS[index]}: formatting changed a character outside ASCII")


TARGETS = {"lexer": fuzz_lexer, "format": fuzz_format}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"usage: fuzz_formatter.py <{'|'.join(TARGETS)}> [libFuzzer options]")
    atheris.Setup([sys.argv[0], *sys.argv[2:]], TARGETS[sys.argv[1]])
    atheris.Fuzz()


if __name__ == "__main__":
    main()
