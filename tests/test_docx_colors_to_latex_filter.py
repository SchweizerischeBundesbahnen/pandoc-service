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


def test_shd_emits_soul_hl_not_colorbox():
    """Background shading must use soul's line-breakable \\hl, NOT
    \\colorbox. Regression for the bug where long shaded runs overflowed
    the right margin because \\colorbox produces a non-breakable hbox.
    """
    latex = _convert_docx_to_latex(FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(latex)
    assert "\\colorbox" not in flat, flat
    assert "{\\definecolor{pdc_hl}{HTML}{00FF00}\\sethlcolor{pdc_hl}\\hl{green shading}}" in flat, flat


def test_named_highlight_resolves_to_hex_and_uses_hl():
    """A Word "yellow" highlight becomes \\hl{...} with the matching hex
    color (resolved via the static name-to-hex table inside the filter)."""
    latex = _convert_docx_to_latex(FIXTURE_PATH.read_bytes())
    flat = _flatten_whitespace(latex)
    assert "{\\definecolor{pdc_hl}{HTML}{FFFF00}\\sethlcolor{pdc_hl}\\hl{yellow highlight}}" in flat, flat


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
    filter must skip the background wrapper in that case while still
    emitting the foreground color (which is harmless around images).
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
    # Background highlight was suppressed — no soul wrappers.
    assert "\\hl{" not in flat, flat
    assert "\\sethlcolor" not in flat, flat


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
# Highlight on bold/italic text — formatting macros hoisted OUTSIDE \hl
# ---------------------------------------------------------------------------
#
# Regression for the bug where a highlighted (background/shaded) run that is
# also bold/italic lost its highlight entirely in DOCX->PDF: soul's \hl can't
# wrap a span containing \emph/\textbf, so the old filter dropped the highlight.
# The fix peels uniform formatting wrappers and emits them as macros AROUND
# \hl, which only ever sees plain text. These build the exact AST shape pandoc
# produces from `-f docx+styles` (Span[custom-style][Emph/Strong[...]]) via a
# markdown bracketed span, so the test is pandoc-only (no tectonic needed).


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


def test_highlight_survives_italic_with_emph_hoisted_outside_hl():
    """The "injected humour" case: italic + foreground color + background.

    The highlight must NOT be dropped. \\emph is hoisted OUTSIDE \\hl (soul
    can't wrap \\emph), \\textcolor stays outermost, and \\hl wraps plain text.
    """
    flat = _md_to_latex('[*injected humour*]{custom-style="PandocColor__FG_00B050__BG_FFFF00"}\n')
    assert "\\hl{injected humour}" in flat, f"highlight dropped on italic text: {flat}"
    assert "\\emph{" in flat, flat
    assert "\\textcolor[HTML]{00B050}" in flat, flat
    # Ordering: \textcolor outermost, then \emph, then \hl innermost.
    assert flat.index("\\textcolor") < flat.index("\\emph{") < flat.index("\\hl{"), flat


def test_highlight_survives_bold():
    """Bold + background: \\textbf hoisted outside \\hl, highlight preserved."""
    flat = _md_to_latex('[**loud**]{custom-style="PandocColor__BG_FFFF00"}\n')
    assert "\\hl{loud}" in flat, f"highlight dropped on bold text: {flat}"
    assert "\\textbf{" in flat and flat.index("\\textbf{") < flat.index("\\hl{"), flat


def test_highlight_survives_bold_italic_nested():
    """Bold-italic nests two wrappers; both must hoist out, \\hl innermost."""
    flat = _md_to_latex('[***both***]{custom-style="PandocColor__BG_FFFF00"}\n')
    assert "\\hl{both}" in flat, f"highlight dropped on bold-italic text: {flat}"
    assert flat.index("\\textbf{") < flat.index("\\hl{"), flat
    assert flat.index("\\emph{") < flat.index("\\hl{"), flat


def test_partial_formatting_inside_highlight_drops_hl_but_keeps_color():
    """When formatting only partially overlaps the highlighted span there is no
    single macro to hoist, so soul can't wrap it. The highlight is gracefully
    dropped (current limitation) while the foreground color and the text — bold
    part included — are preserved."""
    flat = _md_to_latex('[plain *and italic*]{custom-style="PandocColor__FG_00B050__BG_FFFF00"}\n')
    assert "\\hl{" not in flat and "\\sethlcolor" not in flat, f"soul wrapped mixed content: {flat}"
    assert "\\textcolor[HTML]{00B050}" in flat, flat
    assert "\\emph{" in flat and "italic" in flat, flat
