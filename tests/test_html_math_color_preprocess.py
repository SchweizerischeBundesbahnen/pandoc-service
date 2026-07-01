"""Unit tests for ``app.HtmlMathColorPreProcess``.

These verify the *encode* half of the math-color shim: that ``\\color`` /
``\\textcolor`` inside ``<script type="math/tex">`` blocks are
rewritten into ``\\text{@@PMC:RRGGBB@@}...\\text{@@PMCEND@@}`` markers, that color
values (named, ``#hex``, bare hex, ``[HTML]{hex}``) resolve correctly, that nested
colors and braced content are handled, and that non-color/non-math input is left
untouched. The companion decoder is exercised in ``test_docx_math_color_postprocess.py``
and end-to-end in ``test_math_color_integration.py``.
"""

from __future__ import annotations

import re

from app import HtmlMathColorPreProcess
from app.HtmlMathColorPreProcess import MARKER_END, MARKER_PREFIX, MARKER_SUFFIX

_SCRIPT_OPEN = '<html><body><p><script type="math/tex; mode=display">'
_SCRIPT_CLOSE = "</script></p></body></html>"


def _encode_body(latex: str) -> str:
    """Run one math formula through preprocess() and return the rewritten script body."""
    html = (_SCRIPT_OPEN + latex + _SCRIPT_CLOSE).encode("utf-8")
    out = HtmlMathColorPreProcess.preprocess(html).decode("utf-8")
    match = re.search(r'mode=display">(.*)</script>', out, re.DOTALL)
    assert match, f"script body not found in output: {out!r}"
    return match.group(1)


def _start(hex_color: str) -> str:
    return "\\text{" + MARKER_PREFIX + hex_color + MARKER_SUFFIX + "}"


_END = "\\text{" + MARKER_END + "}"


def test_named_color_textcolor() -> None:
    assert _encode_body("\\textcolor{red}{x^2}") == _start("FF0000") + "x^2" + _END


def test_named_color_is_case_insensitive() -> None:
    # Polarion emits dvips-style capitalized names such as \color{Red}.
    assert _encode_body("\\color{Red}{b^2-4ac}") == _start("FF0000") + "b^2-4ac" + _END


def test_bare_hex_color() -> None:
    assert _encode_body("\\color{FFA500}{x}") == _start("FFA500") + "x" + _END


def test_hash_hex_color() -> None:
    assert _encode_body("\\color{#00ff00}{x}") == _start("00FF00") + "x" + _END


def test_three_digit_hex_expands() -> None:
    assert _encode_body("\\color{#0f0}{x}") == _start("00FF00") + "x" + _END


def test_html_model_color() -> None:
    assert _encode_body("\\textcolor[HTML]{EA1B2C}{x}") == _start("EA1B2C") + "x" + _END


def test_mathcolor_is_left_untouched() -> None:
    # \mathcolor is a KaTeX command, not a MathJax one, so Polarion never emits it;
    # we do not transform it (it is not in the color-command set).
    assert _encode_body("\\mathcolor{blue}{y}") == "\\mathcolor{blue}{y}"


def test_braced_content_is_preserved() -> None:
    assert _encode_body("\\textcolor{red}{\\frac{a}{b}}") == _start("FF0000") + "\\frac{a}{b}" + _END


def test_nested_colors() -> None:
    result = _encode_body("\\color{red}{a\\color{blue}{b}c}")
    assert result == _start("FF0000") + "a" + _start("0000FF") + "b" + _END + "c" + _END


def test_color_inside_larger_formula() -> None:
    result = _encode_body("x=\\frac{-b\\pm\\sqrt{\\color{Red}{b^2-4ac}}}{2a}")
    assert result == "x=\\frac{-b\\pm\\sqrt{" + _start("FF0000") + "b^2-4ac" + _END + "}}{2a}"


def test_unknown_color_name_unwraps_content() -> None:
    # An unresolvable color drops the color but keeps the content, so \textcolor
    # (which texmath cannot parse) does not leak the whole formula.
    assert _encode_body("\\textcolor{notacolor}{x^2}") == "x^2"


def test_unsupported_color_model_unwraps_content() -> None:
    assert _encode_body("\\textcolor[rgb]{1,0,0}{x}") == "x"


def test_switch_form_left_untouched() -> None:
    # One-argument \color (color applies to the rest of the group) has no second
    # brace group; we leave it as-is (texmath keeps the content, drops the color).
    assert _encode_body("\\color{red} x + y") == "\\color{red} x + y"


def test_non_color_command_untouched() -> None:
    assert _encode_body("\\frac{a}{b} + \\sqrt{c}") == "\\frac{a}{b} + \\sqrt{c}"


def test_control_symbol_is_copied_verbatim() -> None:
    # A control symbol (backslash + non-letter, e.g. "\,") is copied as its two
    # characters; the surrounding color command still rewrites.
    assert _encode_body("\\,\\color{red}{x}") == "\\," + _start("FF0000") + "x" + _END


def test_unclosed_model_bracket_leaves_command_untouched() -> None:
    # A "[model" with no closing "]" is malformed; the command is copied verbatim.
    latex = "\\textcolor[HTML{FF0000}{x}"
    assert _encode_body(latex) == latex


def test_color_command_without_brace_group_untouched() -> None:
    # A color command not followed by a "{color}" brace group is copied verbatim.
    latex = "\\textcolor abc"
    assert _encode_body(latex) == latex


def test_unbalanced_content_brace_leaves_command_untouched() -> None:
    # The content group "{x" is never closed; the command is copied verbatim.
    latex = "\\color{red}{x"
    assert _encode_body(latex) == latex


def test_color_outside_math_script_is_untouched() -> None:
    # \textcolor in ordinary HTML text (not a math script) must not be rewritten.
    html = b"<html><body><p>literal \\textcolor{red}{x} here</p></body></html>"
    assert HtmlMathColorPreProcess.preprocess(html) == html


def test_no_color_returns_input_unchanged() -> None:
    html = (_SCRIPT_OPEN + "\\frac{a}{b}" + _SCRIPT_CLOSE).encode("utf-8")
    assert HtmlMathColorPreProcess.preprocess(html) is html or HtmlMathColorPreProcess.preprocess(html) == html


def test_multiple_scripts_each_rewritten() -> None:
    html = b'<script type="math/tex">\\color{red}{a}</script>X<script type="math/tex; mode=display">\\textcolor{blue}{b}</script>'
    out = HtmlMathColorPreProcess.preprocess(html).decode("utf-8")
    assert _start("FF0000") + "a" + _END in out
    assert _start("0000FF") + "b" + _END in out
    assert "X" in out


def test_non_utf8_input_passes_through() -> None:
    garbage = b"\xff\xfe\\textcolor{red}{x}"
    assert HtmlMathColorPreProcess.preprocess(garbage) == garbage


def test_single_quoted_script_type() -> None:
    html = b"<script type='math/tex'>\\color{red}{a}</script>"
    out = HtmlMathColorPreProcess.preprocess(html).decode("utf-8")
    assert _start("FF0000") + "a" + _END in out
