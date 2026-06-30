"""Integration tests for ``filters/docx_text_decorations.lua``.

Runs the real ``pandoc`` binary on a native AST (what the docx reader produces)
and checks how Underline/Strikeout are rendered to LaTeX. Pandoc-only — no DOCX
fixture or tectonic needed.

The contract: plain-text underline/strikeout is left to pandoc's soul
``\\ul``/``\\st`` (so it still hyphenates), while underline/strikeout that
carries other formatting is rerouted to ulem ``\\uline``/``\\sout`` (which
compose with nested macros, where soul aborts with "Reconstruction failed").
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

PANDOC = shutil.which("pandoc")
FILTERS_DIR = Path(__file__).resolve().parents[1] / "filters"
DECO = FILTERS_DIR / "docx_text_decorations.lua"
COLORS = FILTERS_DIR / "docx_colors_to_latex.lua"

pytestmark = pytest.mark.skipif(
    PANDOC is None or not DECO.exists(),
    reason="pandoc binary or filters/docx_text_decorations.lua not available",
)


def _native_to_latex(native: str, *filters: Path) -> str:
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.native"
        src.write_text(native, encoding="utf-8")
        cmd = [PANDOC, "-f", "native", "-t", "latex", str(src)]
        for f in filters:
            cmd.append(f"--lua-filter={f}")
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    return " ".join(result.stdout.split())


def test_plain_underline_left_to_soul():
    """Plain text -> pandoc's soul \\ul (which hyphenates); ulem not used."""
    latex = _native_to_latex('[ Para [ Underline [ Str "plain" ] ] ]', DECO)
    assert "\\ul{plain}" in latex, latex
    assert "\\uline" not in latex, latex


def test_underline_with_formatting_uses_ulem():
    """Underline carrying other formatting (here bold) can't go through soul,
    so it is rendered with ulem \\uline."""
    latex = _native_to_latex('[ Para [ Underline [ Strong [ Str "x" ] ] ] ]', DECO)
    assert "\\uline{" in latex, latex
    assert "\\ul{" not in latex, latex


def test_combined_underline_strikeout_both_ulem():
    """When underline and strikeout combine, BOTH must be ulem — a soul \\ul/\\st
    nested inside a ulem command breaks ulem's leaders."""
    latex = _native_to_latex('[ Para [ Underline [ Strikeout [ Str "x" ] ] ] ]', DECO)
    assert "\\uline{" in latex and "\\sout{" in latex, latex
    assert "\\ul{" not in latex and "\\st{" not in latex, latex


def test_linebreak_is_hoisted_out_of_ulem():
    """A forced line break must not sit inside \\uline/\\sout (the leaders can't
    span it); the run is split into one ulem command per line."""
    latex = _native_to_latex('[ Para [ Strikeout [ Str "a", LineBreak, Str "b" ] ] ]', DECO)
    # Two separate \sout commands, and no break inside either.
    assert latex.count("\\sout{") == 2, latex
    assert "\\sout{a}" in latex and "\\sout{b}" in latex, latex


def test_highlight_stripped_inside_decoration():
    """A highlighted span inside underline/strikeout would emit soul \\hl, which
    breaks ulem's leaders. The decoration filter strips the highlight (keeping
    the text) before the colour filter runs, so no \\hl is emitted there."""
    native = '[ Para [ Underline [ Span ("",[],[("custom-style","PandocColor__BG_FFFF00")]) [ Str "x" ] ] ] ]'
    latex = _native_to_latex(native, DECO, COLORS)
    assert "\\uline{" in latex, latex
    assert "\\hl" not in latex, latex


def test_combined_decoration_always_nests_underline_outside_strikeout():
    """ulem draws a \\uline nested inside \\sout at the wrong height (looks like a
    second strike). Regardless of the AST nesting order, the combined decoration
    must come out as \\uline{\\sout{...}} (underline outermost)."""
    # AST has strikeout OUTSIDE underline — the filter must flip it.
    latex = _native_to_latex('[ Para [ Strikeout [ Underline [ Str "x" ] ] ] ]', DECO)
    assert "\\uline{\\sout{x}}" in latex, latex
    assert "\\sout{\\uline{" not in latex, latex


def test_highlight_inside_decoration_becomes_colorbox():
    """A background highlight trapped inside a decoration is rendered with
    \\colorbox (box-safe inside ulem) rather than stripped or soul \\hl."""
    native = '[ Para [ Underline [ Span ("",[],[("custom-style","PandocColor__BG_CC33CC")]) [ Str "x" ] ] ] ]'
    latex = _native_to_latex(native, DECO, COLORS)
    assert "\\colorbox[HTML]{CC33CC}" in latex, latex
    assert "\\hl" not in latex, latex


def test_foreground_color_preserved_inside_decoration():
    """Stripping highlight must keep foreground colour (only __BG_/__HL_ go)."""
    native = '[ Para [ Underline [ Span ("",[],[("custom-style","PandocColor__FG_FF0000__BG_FFFF00")]) [ Str "x" ] ] ] ]'
    latex = _native_to_latex(native, DECO, COLORS)
    assert "\\uline{" in latex and "\\textcolor[HTML]{FF0000}" in latex, latex
    assert "\\hl" not in latex, latex


def test_inter_span_space_inherits_shared_highlight():
    """Two adjacent runs sharing a background leave the Space pandoc lifts out
    between them un-highlighted — a visible gap in the band. The space must be
    wrapped in the shared highlight so the band stays continuous (rendered as a
    boxed space abutting its neighbours)."""
    native = '[ Para [ Span ("",[],[("custom-style","PandocColor__BG_99FFFF")]) [ Str "injected" ] , Space , Span ("",[],[("custom-style","PandocColor__BG_99FFFF")]) [ Strong [ Str "humour" ] ] ] ]'
    latex = _native_to_latex(native, DECO, COLORS)
    # Three highlight boxes: word, the bridged space, word — no un-highlighted gap.
    assert latex.count("\\colorbox[HTML]{99FFFF}") == 3, latex
    # The bridge box must keep its space: \strut{} (not "\strut ") so the space
    # isn't swallowed as the control word's terminator (would jam the words).
    assert "\\strut{} }" in latex, f"bridge space was lost (words would jam): {latex}"
    assert "\\strut{} injected" not in latex  # the word box has no spurious leading space


def test_inter_span_space_not_filled_when_backgrounds_differ():
    """Only a SHARED background bridges the space; differing backgrounds (or a
    plain neighbour) leave the space alone."""
    native = '[ Para [ Span ("",[],[("custom-style","PandocColor__BG_99FFFF")]) [ Str "a" ] , Space , Span ("",[],[("custom-style","PandocColor__BG_FF0000")]) [ Str "b" ] ] ]'
    latex = _native_to_latex(native, DECO, COLORS)
    # Two highlight boxes (a, b); the differing-background space is left plain.
    assert latex.count("\\colorbox[HTML]") == 2, latex
