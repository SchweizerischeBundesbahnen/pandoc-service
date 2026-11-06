"""Integration tests for ``filters/docx/docx_text_decorations.lua``.

Runs the real ``pandoc`` binary inside the pandoc-service container on a
native AST (what the docx reader produces) and checks how Underline/Strikeout
are rendered to LaTeX. Pandoc-only — no DOCX fixture or tectonic needed.

The contract: plain-text underline/strikeout is left to pandoc's soul
``\\ul``/``\\st`` (so it still hyphenates), while underline/strikeout that
carries other formatting is rerouted to ulem ``\\uline``/``\\sout`` (which
compose with nested macros, where soul aborts with "Reconstruction failed").
"""

from __future__ import annotations

from docker.models.containers import Container

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTERS_DIR = "/usr/local/share/pandoc/filters/docx"
DECO = f"{FILTERS_DIR}/docx_text_decorations.lua"
COLORS = f"{FILTERS_DIR}/docx_colors_to_latex.lua"


def _native_to_latex(container: Container, native: str, *filters: str) -> str:
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    cmd = f"{PANDOC_PATH} -f native -t latex /tmp/test/in.native"
    for f in filters:
        cmd += f" --lua-filter={f}"
    exit_code, output = container.exec_run(["sh", "-c", cmd])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return " ".join(output.decode("utf-8").split())


def test_plain_underline_left_to_soul(test_parameters: TestParameters):
    """Plain text -> pandoc's soul \\ul (which hyphenates); ulem not used."""
    latex = _native_to_latex(test_parameters.container, '[ Para [ Underline [ Str "plain" ] ] ]', DECO)
    assert "\\ul{plain}" in latex, latex
    assert "\\uline" not in latex, latex


def test_underline_with_formatting_uses_ulem(test_parameters: TestParameters):
    """Underline carrying other formatting (here bold) can't go through soul,
    so it is rendered with ulem \\uline."""
    latex = _native_to_latex(test_parameters.container, '[ Para [ Underline [ Strong [ Str "x" ] ] ] ]', DECO)
    assert "\\uline{" in latex, latex
    assert "\\ul{" not in latex, latex


def test_combined_underline_strikeout_both_ulem(test_parameters: TestParameters):
    """When underline and strikeout combine, one ulem command draws both: ulem breaks a line only at
    the spaces of its outermost command, and a soul \\ul/\\st inside ulem breaks its leaders."""
    latex = _native_to_latex(test_parameters.container, '[ Para [ Underline [ Strikeout [ Str "x", Space, Str "y" ] ] ] ]', DECO)
    assert "\\pdcUlineSout{x y}" in latex, latex
    assert "\\ul{" not in latex and "\\st{" not in latex, latex


def test_linebreak_is_hoisted_out_of_ulem(test_parameters: TestParameters):
    """A forced line break must not sit inside \\uline/\\sout (the leaders can't
    span it); the run is split into one ulem command per line."""
    latex = _native_to_latex(test_parameters.container, '[ Para [ Strikeout [ Str "a", LineBreak, Str "b" ] ] ]', DECO)
    # Two separate \sout commands, and no break inside either.
    assert latex.count("\\sout{") == 2, latex
    assert "\\sout{a}" in latex and "\\sout{b}" in latex, latex


def test_highlight_stripped_inside_decoration(test_parameters: TestParameters):
    """A highlighted span inside underline/strikeout would emit soul \\hl, which
    breaks ulem's leaders. The decoration filter strips the highlight (keeping
    the text) before the colour filter runs, so no \\hl is emitted there."""
    native = '[ Para [ Underline [ Span ("",[],[("custom-style","PandocColor__BG_FFFF00")]) [ Str "x" ] ] ] ]'
    latex = _native_to_latex(test_parameters.container, native, DECO, COLORS)
    assert "\\uline{" in latex, latex
    assert "\\hl" not in latex, latex


def test_combined_decoration_always_nests_underline_outside_strikeout(test_parameters: TestParameters):
    """Regardless of the AST nesting order, the combined decoration comes out as the one command
    which draws both lines."""
    latex = _native_to_latex(test_parameters.container, '[ Para [ Strikeout [ Underline [ Str "x" ] ] ] ]', DECO)
    assert "\\pdcUlineSout{x}" in latex, latex
    assert "\\sout{\\uline{" not in latex, latex


def test_highlight_inside_decoration_becomes_colorbox(test_parameters: TestParameters):
    """A background highlight trapped inside a decoration is rendered with
    \\colorbox (box-safe inside ulem) rather than stripped or soul \\hl."""
    native = '[ Para [ Underline [ Span ("",[],[("custom-style","PandocColor__BG_CC33CC")]) [ Str "x" ] ] ] ]'
    latex = _native_to_latex(test_parameters.container, native, DECO, COLORS)
    assert "\\colorbox[HTML]{CC33CC}" in latex, latex
    assert "\\hl" not in latex, latex


def test_foreground_color_preserved_inside_decoration(test_parameters: TestParameters):
    """Stripping highlight must keep foreground colour (only __BG_/__HL_ go)."""
    native = '[ Para [ Underline [ Span ("",[],[("custom-style","PandocColor__FG_FF0000__BG_FFFF00")]) [ Str "x" ] ] ] ]'
    latex = _native_to_latex(test_parameters.container, native, DECO, COLORS)
    assert "\\uline{" in latex and "\\textcolor[HTML]{FF0000}" in latex, latex
    assert "\\hl" not in latex, latex


def test_inter_span_space_inherits_shared_highlight(test_parameters: TestParameters):
    """Two adjacent runs sharing a background leave the Space pandoc lifts out
    between them un-highlighted — a visible gap in the band. The space gets the
    shared highlight, as a space in its colour where a line may still end."""
    native = '[ Para [ Span ("",[],[("custom-style","PandocColor__BG_99FFFF")]) [ Str "injected" ] , Space , Span ("",[],[("custom-style","PandocColor__BG_99FFFF")]) [ Strong [ Str "humour" ] ] ] ]'
    latex = _native_to_latex(test_parameters.container, native, DECO, COLORS)
    # Two highlight boxes, and between them a space in their colour: no un-highlighted gap
    assert latex.count("\\colorbox[HTML]{99FFFF}") == 2, latex
    assert "\\strut{}injected}}{\\color[HTML]{99FFFF}\\leaders\\vrule" in latex, latex


def test_inter_span_space_not_filled_when_backgrounds_differ(test_parameters: TestParameters):
    """Only a SHARED background bridges the space; differing backgrounds (or a
    plain neighbour) leave the space alone."""
    native = '[ Para [ Span ("",[],[("custom-style","PandocColor__BG_99FFFF")]) [ Str "a" ] , Space , Span ("",[],[("custom-style","PandocColor__BG_FF0000")]) [ Str "b" ] ] ]'
    latex = _native_to_latex(test_parameters.container, native, DECO, COLORS)
    # Two highlight boxes (a, b); the differing-background space is left plain.
    assert latex.count("\\colorbox[HTML]") == 2, latex
