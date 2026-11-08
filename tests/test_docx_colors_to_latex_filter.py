"""End-to-end integration test for ``filters/docx/docx_colors_to_latex.lua``.

Companion to ``test_inline_styles_filter_integration.py``: runs pandoc inside
the pandoc-service container with the Lua filter loaded and asserts that the
LaTeX output contains the expected ``\\textcolor`` / ``\\colorbox`` raw commands
for runs carrying the synthetic ``custom-style`` attribute the
``docx_color_pre_process`` preprocessor produces.

The test runs the full preprocessor + pandoc + filter pipeline on the
checked-in ``tests/data/colored.docx`` fixture, which contains:

* one paragraph of red foreground text (``<w:color w:val="FF0000"/>``)
* one paragraph of green shading (``<w:shd w:fill="00FF00"/>``)
* one paragraph with the Word "yellow" highlight (``<w:highlight w:val="yellow"/>``)

We use a LaTeX target rather than PDF so the test does not depend on
``tectonic`` being on PATH on dev machines.
"""

from __future__ import annotations

import ctypes
import io
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c
from docker.models.containers import Container
from docx import Document
from docx.shared import Pt

from app import docx_color_pre_process
from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_colors_to_latex.lua"
FIXTURE_PATH = Path(__file__).resolve().parents[1] / "tests" / "data" / "colored.docx"


def _run_pandoc_in_container(container: Container, cmd: str, input_bytes: bytes | None = None, input_file: str | None = None) -> str:
    """Run a pandoc command inside the container, return stdout as string."""
    if input_bytes is not None:
        # Write input bytes to a file inside the container
        import base64

        b64 = base64.b64encode(input_bytes).decode()
        container.exec_run(["sh", "-c", f"echo '{b64}' | base64 -d > {input_file}"])
    exit_code, output = container.exec_run(["sh", "-c", cmd])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode("utf-8")


def _convert_docx_to_latex(container: Container, docx_bytes: bytes) -> str:
    """Run the full preprocess + pandoc + filter pipeline, return LaTeX."""
    preprocessed = docx_color_pre_process.preprocess(docx_bytes)
    _run_pandoc_in_container(container, "mkdir -p /tmp/test", input_bytes=None, input_file=None)
    _run_pandoc_in_container(container, "true", input_bytes=preprocessed, input_file="/tmp/test/in.docx")
    return _run_pandoc_in_container(
        container,
        f"{PANDOC_PATH} -f docx+styles -t latex --lua-filter={FILTER_PATH} /tmp/test/in.docx",
    )


def _md_to_latex(container: Container, md: str) -> str:
    """Convert a markdown snippet to LaTeX through the filter, return stdout."""
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.md << 'HEREDOC_EOF'\n{md}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/in.md"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return _flatten_whitespace(output.decode("utf-8"))


def _md_to_latex_standalone(container: Container, md_or_docx_bytes: bytes, src_format: str = "docx+styles") -> str:
    """Run pandoc with --standalone to get header-includes."""
    import base64

    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    b64 = base64.b64encode(md_or_docx_bytes).decode()
    ext = "docx" if "docx" in src_format else "md"
    container.exec_run(["sh", "-c", f"echo '{b64}' | base64 -d > /tmp/test/in.{ext}"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f {src_format} -t latex --standalone --lua-filter={FILTER_PATH} /tmp/test/in.{ext}"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode("utf-8")


def _flatten_whitespace(latex: str) -> str:
    """Pandoc wraps long output lines at column ~72, which can split our
    raw-LaTeX wrappers across a newline. Collapse whitespace so the assertions
    work regardless of where pandoc decides to wrap."""
    return " ".join(latex.split())


def test_fg_color_emits_textcolor(test_parameters: TestParameters):
    """A red foreground run becomes \\textcolor[HTML]{FF0000}{...}.

    \\textcolor is line-breakable and a no-op for images, so it is always safe
    to apply — that is why foreground color is the only wrapper applied
    unconditionally (background/highlight are gated on content safety).
    """
    latex = _convert_docx_to_latex(test_parameters.container, FIXTURE_PATH.read_bytes())
    assert "\\textcolor[HTML]{FF0000}{red foreground}" in _flatten_whitespace(latex), latex


def test_shd_emits_full_height_colorbox(test_parameters: TestParameters):
    """Background shading is rendered as a full-height \\colorbox (\\strut so the
    band is uniform with decorated highlights), not soul's \\hl. soul's \\hl band
    hugs the glyph height, which left a visible step where a plain highlight met
    an underlined one; boxing every highlight keeps the band level."""
    latex = _convert_docx_to_latex(test_parameters.container, FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(latex)
    assert "\\hl{" not in flat and "\\sethlcolor" not in flat, flat
    # A box per word, and a space in the colour of the background between them, where a line may end
    space = "{\\color[HTML]{00FF00}\\leaders\\vrule height\\ht\\strutbox depth\\dp\\strutbox\\hskip\\fontdimen2\\font plus\\fontdimen3\\font minus\\fontdimen4\\font}"
    words = "\\colorbox[HTML]{00FF00}{\\strut{}green}}" + space + "{\\setlength{\\fboxsep}{0pt}\\colorbox[HTML]{00FF00}{\\strut{}shading}}"
    assert words in flat, flat


def test_named_highlight_resolves_to_hex_and_boxes(test_parameters: TestParameters):
    """A Word "yellow" highlight becomes a \\colorbox with the matching hex
    color (resolved via the static name-to-hex table inside the filter)."""
    latex = _convert_docx_to_latex(test_parameters.container, FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(latex)
    assert "\\colorbox[HTML]{FFFF00}{\\strut{}yellow}}" in flat and "\\colorbox[HTML]{FFFF00}{\\strut{}highlight}}" in flat, flat


def test_soul_package_added_to_header_includes(test_parameters: TestParameters):
    """The filter must inject \\usepackage{soul} into the preamble so \\hl
    is defined when tectonic processes the document. Requires --standalone
    so pandoc emits a full document with header-includes applied."""
    preprocessed = docx_color_pre_process.preprocess(FIXTURE_PATH.read_bytes())
    stdout = _md_to_latex_standalone(test_parameters.container, preprocessed)
    assert "\\usepackage{soul}" in stdout, stdout
    # Both underline mechanisms must be pinned to the same fixed depth so an
    # underline under a larger run does not drop below its normal-size neighbour
    # (soul's \ul and ulem's \uline both scale their depth with the font by
    # default). Regression for the underline-step-on-font-size-change report.
    assert "\\setlength{\\ULdepth}{1.6pt}" in stdout, stdout
    assert "\\setul{1.6pt}{0.4pt}" in stdout, stdout


def test_superscript_subscript_routed_to_ulem_for_box_safety(test_parameters: TestParameters):
    """soul's \\ul/\\st abort with "Reconstruction failed" inside the boxes that
    \\textsuperscript/\\textsubscript build (i.e. underlined/struck text inside a
    <sup>/<sub>). The preamble must load ulem and redefine
    \\textsuperscript/\\textsubscript to swap soul's \\ul/\\st for ulem's box-safe
    \\uline/\\sout *locally* — globally \\ul/\\st stay soul so ordinary
    underlined/struck text is unchanged (no document-wide line-break/hyphenation
    regression). Requires --standalone so header-includes are emitted."""
    preprocessed = docx_color_pre_process.preprocess(FIXTURE_PATH.read_bytes())
    stdout = _md_to_latex_standalone(test_parameters.container, preprocessed)
    flat = _flatten_whitespace(stdout)
    assert "\\usepackage[normalem]{ulem}" in flat, flat
    # The \ul -> \uline / \st -> \sout swap is scoped inside the super/subscript
    # redefinitions, NOT applied globally.
    assert "\\renewcommand{\\textsuperscript}[1]{\\pdcOldSuperscript{\\let\\ul\\uline\\let\\st\\sout" in flat, flat
    assert "\\renewcommand{\\textsubscript}[1]{\\pdcOldSubscript{\\let\\ul\\uline\\let\\st\\sout" in flat, flat


def test_image_inside_styled_span_is_not_wrapped_in_hl(test_parameters: TestParameters):
    """When a styled span contains an Image, soul's \\hl would fail to
    typeset it (and \\colorbox would clip / pad it past the margin). The
    filter must skip BOTH background wrappers (soul \\hl and the \\colorbox
    fallback) in that case while still emitting the foreground color (which
    is harmless around images).
    """
    flat = _md_to_latex(test_parameters.container, '[![alt](img.png) and text]{custom-style="PandocColor__FG_FF0000__BG_00FF00"}\n')
    # Foreground color was applied.
    assert "\\textcolor[HTML]{FF0000}" in flat, flat
    # Background highlight was suppressed — no soul wrappers and, crucially, no
    # \colorbox fallback either (an Image is BOX_UNSAFE).
    assert "\\hl{" not in flat, flat
    assert "\\sethlcolor" not in flat, flat
    assert "\\colorbox" not in flat, flat


def test_font_size_emits_fontsize(test_parameters: TestParameters):
    """A PandocColor__SZ_<halfpoints> style becomes \\fontsize{pt}{...} (pt =
    half-points / 2, so 32 -> 16pt)."""
    flat = _md_to_latex(test_parameters.container, '[big]{custom-style="PandocColor__SZ_32"}\n')
    assert "\\fontsize{16bp}{19.2bp}\\selectfont" in flat, flat
    assert "big" in flat


def test_non_matching_custom_style_is_left_alone(test_parameters: TestParameters):
    """A Span whose custom-style does not start with "PandocColor" must be
    passed through unchanged — the filter only intervenes for its own
    style namespace.
    """
    flat = _md_to_latex(test_parameters.container, '[hello]{custom-style="OtherStyle"}\n')
    # No raw color command should appear since the style is foreign.
    assert "\\textcolor" not in flat
    assert "\\colorbox" not in flat
    assert "\\hl{" not in flat
    assert "hello" in flat


def test_a_style_name_which_is_not_a_colour_is_dropped(test_parameters: TestParameters):
    """A DOCX may carry a style of this namespace with TeX in its name, and the
    value reaches raw LaTeX which this filter emits after the strip filter has
    run. Only six hex digits are written out.
    """
    flat = _md_to_latex(test_parameters.container, '[hello]{custom-style="PandocColor__FG_AAAAAA}{\\input{/etc/hostname}"}\n')
    assert "\\input" not in flat
    assert "\\textcolor" not in flat
    assert "hello" in flat


def test_a_background_which_is_not_a_colour_is_dropped(test_parameters: TestParameters):
    flat = _md_to_latex(test_parameters.container, '[hello]{custom-style="PandocColor__BG_AAAAAA}{\\input{/etc/hostname}"}\n')
    assert "\\input" not in flat
    assert "\\colorbox" not in flat
    assert "hello" in flat


# ---------------------------------------------------------------------------
# Highlight on bold/italic/underlined text — all boxed, uniform band height
# ---------------------------------------------------------------------------
#
# Every highlight is a full-height \colorbox so the background band stays level
# regardless of the inline formatting it carries (bold, italic, underline). A
# \colorbox tolerates any content, so the formatting macros simply sit INSIDE
# the box — no hoisting, no soul \hl. \textcolor (and \fontsize) wrap outside.
# These build the exact AST shape pandoc produces from `-f docx+styles`
# (Span[custom-style][Emph/Strong/Underline[...]]) via a markdown bracketed
# span, so the test is pandoc-only (no tectonic needed).


def test_highlight_on_italic_is_boxed_with_color_outside(test_parameters: TestParameters):
    """The "injected humour" case: italic + foreground color + background.

    The highlight must NOT be dropped. \\emph sits INSIDE the \\colorbox,
    \\textcolor wraps OUTSIDE, and no soul \\hl is emitted."""
    flat = _md_to_latex(test_parameters.container, '[*injected humour*]{custom-style="PandocColor__FG_00B050__BG_FFFF00"}\n')
    assert "\\hl{" not in flat, f"soul \\hl used instead of \\colorbox: {flat}"
    assert "\\colorbox[HTML]{FFFF00}{\\strut{}\\emph{injected}}" in flat and "\\colorbox[HTML]{FFFF00}{\\strut{}\\emph{humour}}" in flat, flat
    # \textcolor is outermost, then the box.
    assert flat.index("\\textcolor") < flat.index("\\colorbox"), flat


def test_highlight_on_bold_is_boxed(test_parameters: TestParameters):
    """Bold + background: \\textbf sits inside the \\colorbox, highlight kept."""
    flat = _md_to_latex(test_parameters.container, '[**loud**]{custom-style="PandocColor__BG_FFFF00"}\n')
    assert "\\hl{" not in flat, flat
    assert "\\colorbox[HTML]{FFFF00}{\\strut{}\\textbf{loud}}" in flat, flat


def test_highlight_on_bold_italic_is_boxed(test_parameters: TestParameters):
    """Bold-italic nests two wrappers; both sit inside the box."""
    flat = _md_to_latex(test_parameters.container, '[***both***]{custom-style="PandocColor__BG_FFFF00"}\n')
    assert "\\hl{" not in flat, flat
    assert "\\colorbox[HTML]{FFFF00}" in flat, flat
    assert "\\textbf{" in flat and "\\emph{" in flat, flat
    assert flat.index("\\colorbox") < flat.index("\\textbf{"), flat


def test_partial_formatting_inside_highlight_is_boxed(test_parameters: TestParameters):
    """A highlight whose formatting only partially overlaps it (plain + italic)
    is still boxed as one unit — the background covers the whole run, with the
    foreground color outside."""
    flat = _md_to_latex(test_parameters.container, '[plain *and italic*]{custom-style="PandocColor__FG_00B050__BG_FFFF00"}\n')
    assert "\\hl{" not in flat and "\\sethlcolor" not in flat, flat
    assert "\\colorbox[HTML]{FFFF00}" in flat, f"background dropped instead of boxed: {flat}"
    assert "\\textcolor[HTML]{00B050}" in flat, flat
    assert "\\emph{" in flat and "italic" in flat, flat


def test_white_background_is_not_boxed(test_parameters: TestParameters):
    """A white (#FFFFFF) background is the page colour — invisible — and Polarion
    stamps it on nearly every run. Boxing it would wrap most of the document in
    non-breakable \\colorboxes and overflow the margin, so white must be treated
    as no highlight: no box, and the text stays line-breakable."""
    flat = _md_to_latex(test_parameters.container, '[If you are going to use a passage of text here]{custom-style="PandocColor__FG_FF0000__BG_FFFFFF"}\n')
    assert "\\colorbox" not in flat, f"white background was boxed: {flat}"
    assert "\\textcolor[HTML]{FF0000}" in flat, flat  # foreground colour still applied


def test_a_long_word_of_a_highlight_is_not_boxed(test_parameters: TestParameters):
    """A \\colorbox is unbreakable, and a highlight is a box per word: a word longer than a line drops the background
    (the foreground colour stays) rather than becoming an overfull box."""
    long_word = "x" * 70
    flat = _md_to_latex(test_parameters.container, f'[a {long_word}]{{custom-style="PandocColor__BG_FFFF00"}}\n')
    assert "\\colorbox" not in flat, f"a word longer than a line was boxed (overflow risk): {flat}"


def test_underline_inside_highlight_is_boxed(test_parameters: TestParameters):
    """A run that is BOTH highlighted and underlined boxes the whole run (with
    the underline inside), so the band stays continuous and level — no
    un-highlighted hole and no height step versus a neighbouring plain
    highlight (which is now also boxed)."""
    flat = _md_to_latex(test_parameters.container, '[[underlined]{.underline}]{custom-style="PandocColor__BG_CC99FF"}\n')
    assert "\\colorbox[HTML]{CC99FF}" in flat, f"background dropped on underlined run: {flat}"
    assert "\\strut" in flat, flat
    assert "\\ul{underlined}" in flat or "\\uline{underlined}" in flat, flat


def _font_style(family: str) -> str:
    return "PandocColor__FN_" + family.encode().hex().upper()


def test_a_run_font_becomes_fontspec_as_word_sets_it(test_parameters: TestParameters):
    latex = _md_to_latex(test_parameters.container, f'[in Segoe]{{custom-style="{_font_style("Segoe UI")}"}}')
    assert "{\\fontspec{Segoe UI}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}]\\pdcNoScriptFont{}in Segoe}" in latex, latex


def test_a_run_font_which_could_carry_latex_is_left_out(test_parameters: TestParameters):
    latex = _md_to_latex(test_parameters.container, f'[evil]{{custom-style="{_font_style("x}" + chr(92) + "input{y")}"}}')
    assert "fontspec" not in latex, latex
    odd = _md_to_latex(test_parameters.container, '[odd]{custom-style="PandocColor__FN_ABC"}')
    assert "fontspec" not in odd, odd


def test_a_run_size_is_in_postscript_points(test_parameters: TestParameters):
    latex = _md_to_latex(test_parameters.container, '[big]{custom-style="PandocColor__SZ_22"}')
    assert "\\fontsize{11bp}{13.2bp}\\selectfont" in latex, latex


def test_a_run_larger_than_the_body_text_makes_its_line_as_high_as_its_font(test_parameters: TestParameters):
    """Word sets a line as high as the ascent and as deep as the descent of the largest font in it."""
    container = test_parameters.container
    container.exec_run(["sh", "-c", 'mkdir -p /tmp/test && printf \'[big]{custom-style="PandocColor__SZ_32"} [small]{custom-style="PandocColor__SZ_20"}\\n\' > /tmp/test/run.md'])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex -M docx-body-size=12 --lua-filter={FILTER_PATH} /tmp/test/run.md"])
    assert exit_code == 0, output.decode()
    flat = _flatten_whitespace(output.decode())
    assert "\\fontsize{16bp}{19.2bp}\\selectfont \\pdcRunStrut{}big" in flat, flat
    assert flat.count("\\pdcRunStrut") == 1, "a smaller run leaves the line of the body"


def test_a_long_highlight_keeps_its_background_word_by_word(test_parameters: TestParameters):
    """A box per word breaks at the spaces: a run longer than a line keeps its background."""
    text = " ".join(["highlighted"] * 8)
    flat = _md_to_latex(test_parameters.container, f'[{text}]{{custom-style="PandocColor__BG_FFFF00"}}\n')
    assert flat.count("\\colorbox[HTML]{FFFF00}") == 8, flat


def _run_line_latex(container: Container, md: str) -> str:
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/line.md << 'EOF'\n{md}\nEOF"])
    metadata = "-M docx-body-size=12 -M docx-body-ascent=11.27 -M docx-body-descent=3.38 -M docx-body-line-multiple=1"
    fonts = f"-M docx-run-fonts={b'Segoe UI'.hex().upper()}:0.9280:1.1640:0.6499:0.4771:0.1382:1.0791:0.2510"
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex {metadata} {fonts} --lua-filter={FILTER_PATH} /tmp/test/line.md"])
    assert exit_code == 0, output.decode()
    return _flatten_whitespace(output.decode())


def test_a_paragraph_of_runs_in_another_font_gets_the_line_of_that_font(test_parameters: TestParameters):
    """Word sets a line as high as the font of its text: Segoe UI at 11 points, 12.8 points, not the body's 14.65."""
    style = f"PandocColor__SZ_22__{_font_style('Segoe UI').removeprefix('PandocColor__')}"
    latex = _run_line_latex(test_parameters.container, f'[one]{{custom-style="{style}"}} [two]{{custom-style="{style}"}}')
    assert latex.startswith("{\\pdcParLine{12.8040bp}{10.2080bp}{2.5960bp}"), latex
    assert latex.endswith("\\par}"), latex


def test_a_paragraph_with_text_in_the_body_font_keeps_its_line(test_parameters: TestParameters):
    style = f"PandocColor__SZ_22__{_font_style('Segoe UI').removeprefix('PandocColor__')}"
    assert "pdcParLine" not in _run_line_latex(test_parameters.container, f'plain [run]{{custom-style="{style}"}}')
    other = f"PandocColor__SZ_24__{_font_style('Segoe UI').removeprefix('PandocColor__')}"
    assert "pdcParLine" not in _run_line_latex(test_parameters.container, f'[one]{{custom-style="{style}"}} [two]{{custom-style="{other}"}}')


def test_a_paragraph_of_runs_inside_formatting_gets_the_line_of_their_font(test_parameters: TestParameters):
    style = f"PandocColor__SZ_22__{_font_style('Segoe UI').removeprefix('PandocColor__')}"
    latex = _run_line_latex(test_parameters.container, f'*[one]{{custom-style="{style}"}}* **[two]{{custom-style="{style}"}}**')
    assert latex.startswith("{\\pdcParLine{12.8040bp}"), latex


def test_a_run_in_another_font_and_size_gets_the_strut_of_that_font(test_parameters: TestParameters):
    """\\selectfont would set LaTeX's strut, 0.7 and 0.3 of the line: a highlight box of it reaches deeper than Word's."""
    style = f"PandocColor__SZ_22__{_font_style('Segoe UI').removeprefix('PandocColor__')}"
    latex = _run_line_latex(test_parameters.container, f'plain [run]{{custom-style="{style}"}}')
    assert "\\selectfont \\setbox\\strutbox\\hbox{\\vrule height 10.2080bp depth 2.5960bp width 0pt}" in latex, latex


def test_a_run_in_another_font_without_a_size_gets_the_strut_of_the_size_around_it(test_parameters: TestParameters):
    """A heading sets the size of a run without one of its own: the strut follows it, not the body size."""
    latex = _run_line_latex(test_parameters.container, f'# plain [run]{{custom-style="{_font_style("Segoe UI")}"}}')
    assert "\\setbox\\strutbox\\hbox{\\vrule height 0.9280\\dimexpr\\csname f@size\\endcsname pt\\relax depth 0.2360\\dimexpr\\csname f@size\\endcsname pt\\relax width 0pt}" in latex, latex


def test_a_paragraph_of_runs_in_another_font_without_a_size_gets_the_line_of_the_size_around_it(test_parameters: TestParameters):
    """A caption or a title sets the size of runs without one of their own: the line follows it, not the body size."""
    latex = _run_line_latex(test_parameters.container, f'[one]{{custom-style="{_font_style("Segoe UI")}"}} [two]{{custom-style="{_font_style("Segoe UI")}"}}')
    size = "\\dimexpr\\csname f@size\\endcsname pt\\relax"
    assert latex.startswith(f"{{\\pdcParLine{{1.1640{size}}}{{0.9280{size}}}{{0.2360{size}}}"), latex


def test_a_coloured_run_keeps_the_space_it_starts_with(test_parameters: TestParameters):
    """\\color ignores the spaces after it: an empty group keeps the space of the run."""
    latex = _md_to_latex(test_parameters.container, 'a[ red]{custom-style="PandocColor__FG_FF0000"}')
    assert "\\textcolor[HTML]{FF0000}{{} red}" in latex, latex


def test_ulem_draws_its_line_over_the_word(test_parameters: TestParameters):
    """Word draws a strike-through over a highlight: ulem drew the line first, and the box of the highlight covered it."""
    preprocessed = docx_color_pre_process.preprocess(FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(_md_to_latex_standalone(test_parameters.container, preprocessed))
    assert "{\\UL@skip\\wd\\UL@box\\box\\UL@box\\kern-\\UL@skip\\UL@leaders}" in flat, flat


def test_a_run_font_sets_its_scripts_as_word_does(test_parameters: TestParameters):
    """A superscript in a run takes the size and the raise of the font of the run, not of the body."""
    latex = _run_line_latex(test_parameters.container, f'plain [run^1^]{{custom-style="{_font_style("Segoe UI")}"}}')
    assert "Ligatures={TeXOff,NoCommon}]\\pdcScriptFont{0.6499}{0.4771}{0.1382}{1.0791}{0.2510}\\setbox" in latex, latex


def test_the_body_font_sets_the_scripts_of_the_document(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf 'x^1^\\n' > /tmp/test/script.md"])
    command = f"{PANDOC_PATH} -f markdown -t latex -s -M docx-body-script=0.7000:0.4800:0.1396:1.0000:0.2358 --lua-filter={FILTER_PATH} /tmp/test/script.md"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    assert "\n\\pdcScriptFont{0.7000}{0.4800}{0.1396}{1.0000}{0.2358}\n" in output.decode(), output.decode()


def test_a_body_script_field_which_could_carry_latex_is_left_out(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf 'x^1^\\n' > /tmp/test/script.md"])
    command = f"{PANDOC_PATH} -f markdown -t latex -s -M 'docx-body-script=1:2:3:4:5}}\\input{{y}}' --lua-filter={FILTER_PATH} /tmp/test/script.md"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    assert "\\pdcScriptFont{" not in output.decode().replace("\\newcommand{\\pdcScriptFont}", ""), output.decode()


def _scripts_docx(font: str, half_points: int) -> bytes:
    """A paragraph in the font and size given with a superscript and a subscript, as Word writes them."""
    document = Document()
    style = document.styles["Normal"]
    style.font.name = font
    style.font.size = Pt(half_points / 2)
    paragraph = document.add_paragraph("Base")
    paragraph.add_run("X").font.superscript = True
    paragraph.add_run(" low")
    paragraph.add_run("Z").font.subscript = True
    docx = io.BytesIO()
    document.save(docx)
    return docx.getvalue()


def _character(pdf: bytes, text: str) -> tuple[float, float]:
    """The size of the first character given and the height of its baseline, in points."""
    page = pdfium.PdfDocument(pdf)[0]
    textpage = page.get_textpage()
    found = textpage.search(text, match_case=True).get_next()
    assert found, text
    x, y = ctypes.c_double(), ctypes.c_double()
    pdfium_c.FPDFText_GetCharOrigin(textpage.raw, found[0], ctypes.byref(x), ctypes.byref(y))
    return pdfium_c.FPDFText_GetFontSize(textpage.raw, found[0]), y.value


def test_scripts_take_the_size_and_the_shift_word_gives_them(test_parameters: TestParameters):
    """DejaVu Sans at 11 points, as Word for Mac sets it: scripts of 7.5 points, 3.5 up and 0.83 down.

    Word sets a script at 0.7 of the size, the size of a script of the font, rounded to a half point. The
    superscript rises until its ascent meets the ascent of the text, an em, and the subscript sinks until the
    descents meet, 0.236 of the size.
    """
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", _scripts_docx("DejaVu Sans", 22), mime)})
    assert response.status_code == 200, response.text
    _, base = _character(response.content, "Base")
    size, superscript = _character(response.content, "X")
    _, subscript = _character(response.content, "Z")
    assert abs(size - 7.5) < 0.05, size
    assert abs(superscript - base - 3.5) < 0.05, superscript - base
    assert abs(base - subscript - 0.825) < 0.05, base - subscript


def test_a_script_takes_a_half_point_of_word(test_parameters: TestParameters):
    """Word rounds a script to a half PostScript point: DejaVu Sans at 12.5 points has scripts of 8.5, not 9 points of TeX."""
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", _scripts_docx("DejaVu Sans", 25), mime)})
    assert response.status_code == 200, response.text
    size, _ = _character(response.content, "X")
    assert abs(size - 8.5) < 0.05, size


def test_a_paragraph_may_start_with_a_script(test_parameters: TestParameters):
    """A superscript first in a paragraph starts the paragraph, as a letter does.

    A body font that is not installed gives the paragraph no strut, so the script of a run is first.
    """
    document = Document()
    document.styles["Normal"].font.name = "No Such Font"
    paragraph = document.add_paragraph()
    script = paragraph.add_run("1")
    script.font.superscript = True
    script.font.name = "DejaVu Sans"
    paragraph.add_run(" Source")
    docx = io.BytesIO()
    document.save(docx)
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), mime)})
    assert response.status_code == 200, response.text[-300:]
