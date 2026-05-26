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


# --- parse failure --------------------------------------------------------
#
# fragments_fromstring is intentionally forgiving — it accepts garbage that
# would normally raise (lxml's HTML parser tolerates bad markup). To exercise
# the except branch we patch the symbol the preprocessor calls.


def test_parse_failure_passes_input_through(mocker):
    """If lxml's parser raises, we MUST return the input bytes unchanged.

    A conversion request must never blow up because of malformed HTML — the
    contract is "best-effort, never break the pipeline". A regression where
    a stricter parser or future lxml version starts raising on input that
    used to pass would otherwise propagate up to the caller as a 500.
    """
    # ValueError is one of the caught types in the except tuple and needs no
    # extra imports here; the assertion is "any caught exception returns the
    # original source untouched".
    mocker.patch(
        "app.HtmlIndentPreProcess.html.fragments_fromstring",
        side_effect=ValueError("synthetic parse failure"),
    )
    src = b'<p style="margin-left: 40px">x</p>'
    assert HtmlIndentPreProcess.preprocess(src) is src


def test_parse_failure_logs_warning(mocker, caplog):
    """The parse-failure path must log so we can diagnose runtime hits in
    production. Lock down the log message text so the operator's grep / alert
    rule on the service logs keeps firing if someone refactors the branch.
    """
    import logging  # noqa: PLC0415

    mocker.patch(
        "app.HtmlIndentPreProcess.html.fragments_fromstring",
        side_effect=ValueError("boom"),
    )
    caplog.set_level(logging.WARNING, logger="app.HtmlIndentPreProcess")
    HtmlIndentPreProcess.preprocess(b"<p>x</p>")
    assert any("HtmlIndentPreProcess" in rec.message for rec in caplog.records), f"expected a WARNING from HtmlIndentPreProcess; got: {[r.message for r in caplog.records]!r}"


# --- leading text ---------------------------------------------------------
#
# fragments_fromstring may emit a leading str (any text before the first
# element). That string carries no parent and must be re-emitted verbatim
# in front of the wrapped element on the output side.


def test_leading_text_before_indented_p_is_preserved_when_rewriting():
    """Text before the indented <p> must come first in the output.

    Verifies both the capture branch (leading_text = frag) and the emit
    branch (parts.append(leading_text.encode(...))) — that pair only fires
    together when there's text AND at least one paragraph gets wrapped.
    """
    src = b'hello there <p style="margin-left: 40px">indented</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert out.startswith(b"hello there"), f"leading text not preserved: {out!r}"
    assert b'class="pandoc-indent"' in out


def test_leading_text_alone_does_not_force_rewrite():
    """Leading text without any wrap-target must return the source unchanged.

    Otherwise we'd burn the round-trip-through-lxml cost on inputs we have
    no work to do on — and the byte-identical pass-through guarantee is what
    lets us promise "valid HTML is never modified".
    """
    src = b"hello there <p>no indent</p>"
    assert HtmlIndentPreProcess.preprocess(src) is src


# --- decimal / rounding behavior ------------------------------------------


@pytest.mark.parametrize(
    ("value", "expected_twips"),
    [
        ("0.5px", 8),  # 0.5 * 15 = 7.5 → 8 (round-half-to-even, then up)
        ("1.5px", 22),  # 1.5 * 15 = 22.5 → 22 (banker's rounding: even)
        ("0.5pt", 10),  # 0.5 * 20 = 10
        ("2.54cm", 1440),  # exactly 1 inch
        ("25.4mm", 1440),  # exactly 1 inch
        ("0.75em", 180),  # 0.75 * 240
    ],
)
def test_subpixel_and_decimal_lengths(value: str, expected_twips: int):
    """Decimal CSS lengths must round to the nearest integer twip.

    Python's round() uses banker's rounding (half-to-even). The expected
    values bake that in: 7.5 → 8, 22.5 → 22. If someone swaps to int() or
    math.floor() one day, this test will catch it.
    """
    src = f'<p style="margin-left: {value}">x</p>'.encode()
    assert _twips_after_preprocess(src) == expected_twips


def test_very_large_indent_does_not_overflow():
    """Sanity: huge values still work — no overflow, no exception."""
    src = b'<p style="margin-left: 10000px">x</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b'data-indent-twips="150000"' in out


# --- additional style-attribute shapes ------------------------------------


@pytest.mark.parametrize(
    "style",
    [
        b"MARGIN-LEFT: 40px",  # uppercase property
        b"Margin-Left: 40Px",  # mixed case property + unit
        b"margin-left:40px",  # no spaces around colon
        b"margin-left :40px",  # space before colon
        b"margin-left: 40PX;",  # uppercase unit with semicolon
    ],
)
def test_property_and_unit_are_case_insensitive(style: bytes):
    src = b'<p style="' + style + b'">x</p>'
    assert _twips_after_preprocess(src) == 600


def test_last_margin_left_wins_when_declared_twice():
    """CSS says later declarations override earlier; for now we deliberately
    match the FIRST occurrence (regex.search is leftmost). Document that as
    the contract so any future improvement to honor "last-wins" is
    intentional — not an accidental regression.

    Real-world Polarion never emits duplicate margin-left, so the simpler
    implementation is acceptable.
    """
    src = b'<p style="margin-left: 40px; margin-left: 80px">x</p>'
    # Currently 40px (first) wins; flip the assertion if/when last-wins is added.
    assert _twips_after_preprocess(src) == 600


def test_empty_p_with_margin_left_is_still_wrapped():
    """An empty paragraph that happens to be indented is still wrapped —
    visual whitespace alone is valid content in a document."""
    src = b'<p style="margin-left: 40px"></p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b'class="pandoc-indent"' in out
    assert b'data-indent-twips="600"' in out


def test_indented_p_with_nested_inline_children_keeps_structure():
    """Nested inline elements inside the indented <p> must not be flattened
    or reordered — the wrap is purely additive."""
    src = b'<p style="margin-left: 40px">a <strong>b</strong> <em>c</em> d</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"<strong>b</strong>" in out
    assert b"<em>c</em>" in out
    # The strong appears before the em in the output (order preserved).
    assert out.index(b"<strong") < out.index(b"<em")


def test_scientific_notation_lengths_are_rejected():
    """`1e10px` must not be silently misinterpreted — the numeric regex
    intentionally doesn't accept exponents.
    """
    src = b'<p style="margin-left: 1e10px">x</p>'
    out = HtmlIndentPreProcess.preprocess(src)
    assert b"pandoc-indent" not in out


# --- module-level contract ------------------------------------------------


def test_module_constants_match_documented_contract():
    """The class name and attribute key are part of the contract with the
    Lua filter — pin them so a casual rename here doesn't silently break
    filters/inline_styles.lua, which hard-codes the same strings."""
    assert HtmlIndentPreProcess.INDENT_CLASS == "pandoc-indent"
    assert HtmlIndentPreProcess.INDENT_ATTR == "data-indent-twips"


# --- helper-function targeted tests ---------------------------------------
#
# Targeted at the private helpers so failures point straight at the
# arithmetic, not at the HTML wrapping layer above it.


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("40px", 600),
        ("0", None),
        ("-1px", None),
        ("50%", None),
        ("auto", None),
        ("", None),
        ("abc", None),
        ("40xyz", None),  # unknown unit
        ("+40px", 600),  # explicit + sign
    ],
)
def test_css_length_to_twips_direct(value: str, expected: int | None):
    assert HtmlIndentPreProcess._css_length_to_twips(value) == expected


@pytest.mark.parametrize(
    ("style", "expected"),
    [
        ("margin-left: 40px", 600),
        ("color: red", None),
        ("", None),
        ("margin-right: 40px", None),  # different property
        ("padding-left: 40px", None),  # padding-left is not handled
        ("margin-left: 40px !important", None),  # !important not handled — document the limitation
    ],
)
def test_extract_margin_left_twips_direct(style: str, expected: int | None):
    assert HtmlIndentPreProcess._extract_margin_left_twips(style) == expected
