"""End-to-end integration test for ``filters/docx_colors_to_latex.lua``.

Companion to ``test_inline_styles_filter_integration.py``: spawns the real
``pandoc`` binary with the lua filter loaded and asserts that the LaTeX
output contains the expected ``\\textcolor`` / ``\\colorbox`` raw commands
for runs carrying the synthetic ``custom-style`` attribute the
``DocxColorPreProcess`` preprocessor produces.

The test runs the full preprocessor + pandoc + filter pipeline on the
checked-in ``tests/data/colored.docx`` fixture, which contains:

* one paragraph of red foreground text (``<w:color w:val="FF0000"/>``)
* one paragraph of green shading (``<w:shd w:fill="00FF00"/>``)
* one paragraph with the Word "yellow" highlight (``<w:highlight w:val="yellow"/>``)

We use a LaTeX target rather than PDF so the test does not depend on
``tectonic`` being on PATH on dev machines.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

from app import DocxColorPreProcess

PANDOC = shutil.which("pandoc")
REPO_ROOT = Path(__file__).resolve().parents[1]
FILTER_PATH = REPO_ROOT / "filters" / "docx_colors_to_latex.lua"
FIXTURE_PATH = REPO_ROOT / "tests" / "data" / "colored.docx"

pytestmark = pytest.mark.skipif(
    PANDOC is None or not FILTER_PATH.exists() or not FIXTURE_PATH.exists(),
    reason="pandoc binary, docx_colors_to_latex.lua, or tests/data/colored.docx not available",
)


def _convert_docx_to_latex(docx_bytes: bytes) -> str:
    """Run the full preprocess + pandoc + filter pipeline, return LaTeX."""
    preprocessed = DocxColorPreProcess.preprocess(docx_bytes)
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.docx"
        src.write_bytes(preprocessed)
        result = subprocess.run(
            [
                PANDOC,
                "-f",
                "docx+styles",
                "-t",
                "latex",
                f"--lua-filter={FILTER_PATH}",
                str(src),
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            pytest.fail(f"pandoc failed (exit {result.returncode}):\n{result.stderr}")
        return result.stdout


def _flatten_whitespace(latex: str) -> str:
    """Pandoc wraps long output lines at column ~72, which can split our
    raw-LaTeX wrappers across a newline. Collapse whitespace so the assertions
    work regardless of where pandoc decides to wrap."""
    return " ".join(latex.split())


def test_fg_color_emits_textcolor():
    """A red foreground run becomes \\textcolor[HTML]{FF0000}{...}.

    \\textcolor is line-breakable and a no-op for images, so it is always safe
    to apply — that is why foreground color is the only wrapper applied
    unconditionally (background/highlight are gated on content safety).
    """
    latex = _convert_docx_to_latex(FIXTURE_PATH.read_bytes())
    assert "\\textcolor[HTML]{FF0000}{red foreground}" in _flatten_whitespace(latex), latex


def test_shd_emits_full_height_colorbox():
    """Background shading is rendered as a full-height \\colorbox (\\strut so the
    band is uniform with decorated highlights), not soul's \\hl. soul's \\hl band
    hugs the glyph height, which left a visible step where a plain highlight met
    an underlined one; boxing every highlight keeps the band level."""
    latex = _convert_docx_to_latex(FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(latex)
    assert "\\hl{" not in flat and "\\sethlcolor" not in flat, flat
    assert "\\colorbox[HTML]{00FF00}{\\strut{}green shading}" in flat, flat


def test_named_highlight_resolves_to_hex_and_boxes():
    """A Word "yellow" highlight becomes a \\colorbox with the matching hex
    color (resolved via the static name-to-hex table inside the filter)."""
    latex = _convert_docx_to_latex(FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(latex)
    assert "\\colorbox[HTML]{FFFF00}{\\strut{}yellow highlight}" in flat, flat


def test_soul_package_added_to_header_includes():
    """The filter must inject \\usepackage{soul} into the preamble so \\hl
    is defined when tectonic processes the document. Requires --standalone
    so pandoc emits a full document with header-includes applied."""
    with tempfile.TemporaryDirectory() as tmpdir:
        preprocessed = DocxColorPreProcess.preprocess(FIXTURE_PATH.read_bytes())
        src = Path(tmpdir) / "in.docx"
        src.write_bytes(preprocessed)
        result = subprocess.run(
            [PANDOC, "-f", "docx+styles", "-t", "latex", "--standalone", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    assert "\\usepackage{soul}" in result.stdout, result.stdout
    # Both underline mechanisms must be pinned to the same fixed depth so an
    # underline under a larger run does not drop below its normal-size neighbour
    # (soul's \ul and ulem's \uline both scale their depth with the font by
    # default). Regression for the underline-step-on-font-size-change report.
    assert "\\setlength{\\ULdepth}{1.6pt}" in result.stdout, result.stdout
    assert "\\setul{1.6pt}{0.4pt}" in result.stdout, result.stdout


def test_superscript_subscript_routed_to_ulem_for_box_safety():
    """soul's \\ul/\\st abort with "Reconstruction failed" inside the boxes that
    \\textsuperscript/\\textsubscript build (i.e. underlined/struck text inside a
    <sup>/<sub>). The preamble must load ulem and redefine
    \\textsuperscript/\\textsubscript to swap soul's \\ul/\\st for ulem's box-safe
    \\uline/\\sout *locally* — globally \\ul/\\st stay soul so ordinary
    underlined/struck text is unchanged (no document-wide line-break/hyphenation
    regression). Requires --standalone so header-includes are emitted."""
    with tempfile.TemporaryDirectory() as tmpdir:
        preprocessed = DocxColorPreProcess.preprocess(FIXTURE_PATH.read_bytes())
        src = Path(tmpdir) / "in.docx"
        src.write_bytes(preprocessed)
        result = subprocess.run(
            [PANDOC, "-f", "docx+styles", "-t", "latex", "--standalone", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    flat = _flatten_whitespace(result.stdout)
    assert "\\usepackage[normalem]{ulem}" in flat, flat
    # The \ul -> \uline / \st -> \sout swap is scoped inside the super/subscript
    # redefinitions, NOT applied globally.
    assert "\\renewcommand{\\textsuperscript}[1]{\\pdcOldSuperscript{\\let\\ul\\uline\\let\\st\\sout" in flat, flat
    assert "\\renewcommand{\\textsubscript}[1]{\\pdcOldSubscript{\\let\\ul\\uline\\let\\st\\sout" in flat, flat


def test_image_inside_styled_span_is_not_wrapped_in_hl():
    """When a styled span contains an Image, soul's \\hl would fail to
    typeset it (and \\colorbox would clip / pad it past the margin). The
    filter must skip BOTH background wrappers (soul \\hl and the \\colorbox
    fallback) in that case while still emitting the foreground color (which
    is harmless around images).
    """
    # Markdown bracketed span carrying both color and an image. Pandoc
    # turns this into a Span node with our custom-style attribute and
    # an Image child — exactly the AST shape we want to exercise.
    md = '[![alt](img.png) and text]{custom-style="PandocColor__FG_FF0000__BG_00FF00"}\n'
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.md"
        src.write_text(md, encoding="utf-8")
        result = subprocess.run(
            [PANDOC, "-f", "markdown", "-t", "latex", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    flat = _flatten_whitespace(result.stdout)
    # Foreground color was applied.
    assert "\\textcolor[HTML]{FF0000}" in flat, flat
    # Background highlight was suppressed — no soul wrappers and, crucially, no
    # \colorbox fallback either (an Image is BOX_UNSAFE).
    assert "\\hl{" not in flat, flat
    assert "\\sethlcolor" not in flat, flat
    assert "\\colorbox" not in flat, flat


def test_font_size_emits_fontsize():
    """A PandocColor__SZ_<halfpoints> style becomes \\fontsize{pt}{...} (pt =
    half-points / 2, so 32 -> 16pt)."""
    md = '[big]{custom-style="PandocColor__SZ_32"}\n'
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.md"
        src.write_text(md, encoding="utf-8")
        result = subprocess.run(
            [PANDOC, "-f", "markdown", "-t", "latex", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    flat = _flatten_whitespace(result.stdout)
    assert "\\fontsize{16}{19.2}\\selectfont" in flat, flat
    assert "big" in flat


def test_non_matching_custom_style_is_left_alone():
    """A Span whose custom-style does not start with "PandocColor" must be
    passed through unchanged — the filter only intervenes for its own
    style namespace.
    """
    md = '[hello]{custom-style="OtherStyle"}\n'
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.md"
        src.write_text(md, encoding="utf-8")
        result = subprocess.run(
            [PANDOC, "-f", "markdown", "-t", "latex", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    # No raw color command should appear since the style is foreign.
    assert "\\textcolor" not in result.stdout
    assert "\\colorbox" not in result.stdout
    assert "\\hl{" not in result.stdout
    assert "hello" in result.stdout


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


def _md_to_latex(md: str) -> str:
    """Convert a markdown snippet to LaTeX through the filter, return stdout."""
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.md"
        src.write_text(md, encoding="utf-8")
        result = subprocess.run(
            [PANDOC, "-f", "markdown", "-t", "latex", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    return _flatten_whitespace(result.stdout)


def test_highlight_on_italic_is_boxed_with_color_outside():
    """The "injected humour" case: italic + foreground color + background.

    The highlight must NOT be dropped. \\emph sits INSIDE the \\colorbox,
    \\textcolor wraps OUTSIDE, and no soul \\hl is emitted."""
    flat = _md_to_latex('[*injected humour*]{custom-style="PandocColor__FG_00B050__BG_FFFF00"}\n')
    assert "\\hl{" not in flat, f"soul \\hl used instead of \\colorbox: {flat}"
    assert "\\colorbox[HTML]{FFFF00}{\\strut{}\\emph{injected humour}}" in flat, flat
    # \textcolor is outermost, then the box.
    assert flat.index("\\textcolor") < flat.index("\\colorbox"), flat


def test_highlight_on_bold_is_boxed():
    """Bold + background: \\textbf sits inside the \\colorbox, highlight kept."""
    flat = _md_to_latex('[**loud**]{custom-style="PandocColor__BG_FFFF00"}\n')
    assert "\\hl{" not in flat, flat
    assert "\\colorbox[HTML]{FFFF00}{\\strut{}\\textbf{loud}}" in flat, flat


def test_highlight_on_bold_italic_is_boxed():
    """Bold-italic nests two wrappers; both sit inside the box."""
    flat = _md_to_latex('[***both***]{custom-style="PandocColor__BG_FFFF00"}\n')
    assert "\\hl{" not in flat, flat
    assert "\\colorbox[HTML]{FFFF00}" in flat, flat
    assert "\\textbf{" in flat and "\\emph{" in flat, flat
    assert flat.index("\\colorbox") < flat.index("\\textbf{"), flat


def test_partial_formatting_inside_highlight_is_boxed():
    """A highlight whose formatting only partially overlaps it (plain + italic)
    is still boxed as one unit — the background covers the whole run, with the
    foreground color outside."""
    flat = _md_to_latex('[plain *and italic*]{custom-style="PandocColor__FG_00B050__BG_FFFF00"}\n')
    assert "\\hl{" not in flat and "\\sethlcolor" not in flat, flat
    assert "\\colorbox[HTML]{FFFF00}" in flat, f"background dropped instead of boxed: {flat}"
    assert "\\textcolor[HTML]{00B050}" in flat, flat
    assert "\\emph{" in flat and "italic" in flat, flat


def test_white_background_is_not_boxed():
    """A white (#FFFFFF) background is the page colour — invisible — and Polarion
    stamps it on nearly every run. Boxing it would wrap most of the document in
    non-breakable \\colorboxes and overflow the margin, so white must be treated
    as no highlight: no box, and the text stays line-breakable."""
    flat = _md_to_latex('[If you are going to use a passage of text here]{custom-style="PandocColor__FG_FF0000__BG_FFFFFF"}\n')
    assert "\\colorbox" not in flat, f"white background was boxed: {flat}"
    assert "\\textcolor[HTML]{FF0000}" in flat, flat  # foreground colour still applied


def test_long_highlight_is_not_boxed():
    """A \\colorbox is unbreakable, so boxing a long run overflows the page by
    metres. Only SHORT runs are boxed; a long highlighted run drops the
    background (foreground color stays) rather than becoming an overfull box."""
    long_text = "lorem ipsum dolor sit amet " * 6  # well over the ~60-char cap
    flat = _md_to_latex(f'[{long_text}*it*]{{custom-style="PandocColor__BG_FFFF00"}}\n')
    assert "\\colorbox" not in flat, f"long run was boxed (overflow risk): {flat}"
    assert "\\hl{" not in flat, flat


def test_underline_inside_highlight_is_boxed():
    """A run that is BOTH highlighted and underlined boxes the whole run (with
    the underline inside), so the band stays continuous and level — no
    un-highlighted hole and no height step versus a neighbouring plain
    highlight (which is now also boxed)."""
    flat = _md_to_latex('[[underlined]{.underline}]{custom-style="PandocColor__BG_CC99FF"}\n')
    assert "\\colorbox[HTML]{CC99FF}" in flat, f"background dropped on underlined run: {flat}"
    assert "\\strut" in flat, flat
    assert "\\ul{underlined}" in flat or "\\uline{underlined}" in flat, flat
