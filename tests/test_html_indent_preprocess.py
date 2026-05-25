"""Unit tests for ``app.HtmlIndentPreProcess``.

These tests verify the Python side of the paragraph-indent pipeline: that
``<p style="margin-left:...">`` is rewritten into a marker ``<div>`` carrying
``class="pandoc-indent"`` and ``data-indent-twips="N"``, and that the unit
conversion from CSS lengths to Word twips is correct.

The companion Lua filter (filters/inline_styles.lua) is exercised separately
in ``test_inline_styles_filter_integration.py``.
"""

from __future__ import annotations

import pytest

from app import HtmlIndentPreProcess


def _twips_after_preprocess(src: bytes) -> int | None:
    """Helper: return the data-indent-twips value as int, or None if not wrapped."""
    out = HtmlIndentPreProcess.preprocess(src)
    if b"data-indent-twips=" not in out:
        return None
    # Look for data-indent-twips="N" in the output.
    import re

    m = re.search(rb'data-indent-twips="(\d+)"', out)
    assert m, f"data-indent-twips marker present but unparseable: {out!r}"
    return int(m.group(1))


# --- happy path -----------------------------------------------------------


def test_px_margin_left_is_wrapped_with_twips():
    """40px at the CSS reference DPI (96) = 1/2.4 inch = 600 twips."""
    src = b'<p style="margin-left: 40px;">hi</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b'class="pandoc-indent"' in out
    assert b'data-indent-twips="600"' in out
    assert b"<p" in out and b"hi" in out


def test_user_provided_polarion_snippet():
    """Two paragraphs at different indents each get wrapped independently."""
    src = b'<p id="polarion_1" style="margin-left: 40px;">Indentation</p><p id="polarion_2" style="margin-left: 80px;">2 levels</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert out.count(b'class="pandoc-indent"') == 2
    assert b'data-indent-twips="600"' in out
    assert b'data-indent-twips="1200"' in out


# --- unit conversion ------------------------------------------------------


@pytest.mark.parametrize(
    ("value", "expected_twips"),
    [
        ("40px", 600),  # 1 px = 15 twips
        ("80px", 1200),
        ("12pt", 240),  # 1 pt = 20 twips
        ("0.5in", 720),  # 1 in = 1440 twips
        ("1in", 1440),
        ("1cm", 567),  # 1 cm ≈ 566.93 twips → rounded
        ("10mm", 567),  # 10 mm = 1 cm
        ("1pc", 240),  # 1 pc = 12 pt
        ("1em", 240),  # default 12pt body size
        ("2rem", 480),
        ("40", 600),  # bare number — interpreted as px
    ],
)
def test_css_length_conversion(value: str, expected_twips: int):
    src = f'<p style="margin-left: {value}">x</p>'.encode()
    assert _twips_after_preprocess(src) == expected_twips


@pytest.mark.parametrize(
    "value",
    [
        "0",  # zero indent → no wrapping
        "0px",
        "-40px",  # negative → no wrapping (Word indents are >= 0)
        "50%",  # percentages can't resolve without container width
        "auto",  # keyword we don't understand
        "calc(40px + 10px)",  # function — too complex to evaluate
    ],
)
def test_invalid_or_zero_lengths_do_not_wrap(value: str):
    src = f'<p style="margin-left: {value}">x</p>'.encode()
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"pandoc-indent" not in out


# --- style parsing --------------------------------------------------------


def test_margin_left_among_other_declarations_is_found():
    """The regex must locate margin-left when surrounded by other props."""
    src = b'<p style="color: red; margin-left: 40px; font-size: 12pt;">x</p>'
    assert _twips_after_preprocess(src) == 600


def test_margin_left_with_extra_whitespace():
    src = b'<p style="  margin-left  :   40px   ;">x</p>'
    assert _twips_after_preprocess(src) == 600


def test_margin_shorthand_is_not_treated_as_margin_left():
    """`margin: 0 0 0 40px` is a shorthand we deliberately don't parse."""
    src = b'<p style="margin: 0 0 0 40px">x</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"pandoc-indent" not in out


def test_inline_style_left_alone_inside_indented_p():
    """The original <p> keeps its style attribute — the Lua filter doesn't
    need it stripped, and we don't want to alter unrelated CSS."""
    src = b'<p style="margin-left: 40px; color: red">x</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"color: red" in out


# --- pass-through ---------------------------------------------------------


def test_p_without_style_is_unchanged():
    src = b"<p>plain</p>"
    assert HtmlIndentPreProcess.preprocess(src) == src


def test_p_with_style_but_no_margin_left_is_unchanged():
    src = b'<p style="color: red">plain</p>'
    assert HtmlIndentPreProcess.preprocess(src) == src


def test_empty_input_is_returned_unchanged():
    assert HtmlIndentPreProcess.preprocess(b"") == b""


def test_non_p_elements_are_ignored():
    """Only <p> is targeted; <div style="margin-left:..."> is left alone for
    now (would need separate handling — see the limitations note)."""
    src = b'<div style="margin-left: 40px">x</div>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"pandoc-indent" not in out


# --- structure / context --------------------------------------------------


def test_top_level_p_is_wrapped():
    """Regression for the synthetic-root fix: a <p> that is a top-level
    fragment (no parent in the fragment list) must still be wrapped — we
    re-parent fragments under a synthetic root before walking."""
    src = b'<p style="margin-left: 40px">top-level</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"pandoc-indent" in out


def test_nested_p_is_wrapped():
    """A <p> deep inside other elements is still found and wrapped."""
    src = b'<div><section><p style="margin-left: 40px">deep</p></section></div>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"pandoc-indent" in out


def test_multiple_paragraphs_with_mixed_indentation():
    src = b'<p style="margin-left: 40px">A</p><p>plain</p><p style="margin-left: 80px">B</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert out.count(b'class="pandoc-indent"') == 2
    assert b">plain</p>" in out  # un-wrapped one still exists as-is
