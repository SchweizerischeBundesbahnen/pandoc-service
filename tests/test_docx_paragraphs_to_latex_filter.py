"""Integration tests for ``filters/docx_paragraphs_to_latex.lua``.

Spawns the real ``pandoc`` binary with the filter loaded and asserts the LaTeX
output for a ``Div`` carrying the synthetic ``custom-style="PandocPara__..."``
attribute the ``DocxParagraphPreProcess`` preprocessor produces. The AST shape
is built via a markdown fenced div (``::: {custom-style="..."}``), which pandoc
turns into exactly that ``Div`` node — so the test is pandoc-only and needs no
DOCX fixture or tectonic.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

PANDOC = shutil.which("pandoc")
FILTER_PATH = Path(__file__).resolve().parents[1] / "filters" / "docx_paragraphs_to_latex.lua"

pytestmark = pytest.mark.skipif(
    PANDOC is None or not FILTER_PATH.exists(),
    reason="pandoc binary or filters/docx_paragraphs_to_latex.lua not available",
)


def _md_to_latex(md: str) -> str:
    """Convert a markdown snippet to LaTeX through the filter; collapse
    whitespace so assertions don't depend on pandoc's line wrapping."""
    with tempfile.TemporaryDirectory() as tmpdir:
        src = Path(tmpdir) / "in.md"
        src.write_text(md, encoding="utf-8")
        result = subprocess.run(
            [PANDOC, "-f", "markdown", "-t", "latex", f"--lua-filter={FILTER_PATH}", str(src)],
            capture_output=True,
            text=True,
            check=True,
        )
    return " ".join(result.stdout.split())


def _div(style: str, text: str = "content") -> str:
    return f'::: {{custom-style="{style}"}}\n{text}\n:::\n'


def test_center_alignment_emits_centering():
    flat = _md_to_latex(_div("PandocPara__ALIGN_center"))
    assert "\\centering" in flat, flat
    assert "\\par}" in flat, flat


def test_right_alignment_emits_raggedleft():
    assert "\\raggedleft" in _md_to_latex(_div("PandocPara__ALIGN_right"))


def test_left_alignment_emits_raggedright():
    assert "\\raggedright" in _md_to_latex(_div("PandocPara__ALIGN_left"))


def test_indent_emits_leftskip_in_points():
    """600 twips = 30pt, 1200 twips = 60pt — distinct, not collapsed."""
    assert "\\leftskip=30.00pt" in _md_to_latex(_div("PandocPara__IND_600"))
    assert "\\leftskip=60.00pt" in _md_to_latex(_div("PandocPara__IND_1200"))


def test_alignment_and_indent_combined():
    flat = _md_to_latex(_div("PandocPara__ALIGN_right__IND_600"))
    assert "\\leftskip=30.00pt" in flat and "\\raggedleft" in flat, flat
    # Balanced group: one opening brace before content, \par} after.
    assert flat.count("\\par}") == 1, flat


def test_invalid_indent_value_is_dropped():
    """A non-integer / out-of-range twips value must not reach \\leftskip."""
    flat = _md_to_latex(_div("PandocPara__IND_99999999"))  # exceeds the 31680 cap
    assert "\\leftskip" not in flat, flat


def test_non_pandocpara_div_is_left_alone():
    flat = _md_to_latex(_div("SomeOtherStyle"))
    assert "\\centering" not in flat
    assert "\\raggedleft" not in flat
    assert "\\leftskip" not in flat
    assert "content" in flat
