"""End-to-end test for DOCX table width/alignment surviving into LaTeX.

Builds a real DOCX with a table carrying ``<w:tblW>`` / ``<w:jc>`` (what the
HTML->DOCX post-processing writes), runs it through the docx->latex
preprocessing + ``filters/docx_tables_to_latex.lua``, and checks that:

* pandoc's DOCX reader would normalise the column widths to sum to 1.0, but the
  filter scales them back to the table's real page fraction (``\\real{...}``);
* the table's alignment is re-applied as longtable ``\\LTleft``/``\\LTright`` glue.
"""

from __future__ import annotations

import io
import re
import shutil
import subprocess

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from app import DocxLatexPreProcess

_PANDOC = shutil.which("pandoc")
pytestmark = pytest.mark.skipif(_PANDOC is None, reason="pandoc binary not available")

_FILTER = "filters/docx_tables_to_latex.lua"


def _docx_with_table(width_pct: int, jc: str, cols: int = 3) -> bytes:
    """A one-row DOCX table with the given tblW (pct) and jc alignment."""
    doc = Document()
    table = doc.add_table(rows=1, cols=cols)
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = f"c{i}"
    tblpr = table._tbl.tblPr
    # python-docx already emits a <w:tblW>; replace it (and any jc) so the table
    # has exactly one of each — mirroring app/DocxPostProcess.py's output.
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for tag in ("w:tblW", "w:jc"):
        for existing in tblpr.findall(tag, ns):
            tblpr.remove(existing)
    tblpr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{width_pct}" w:type="pct"/>'))
    tblpr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc}"/>'))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _to_latex(docx_bytes: bytes) -> str:
    pre = DocxLatexPreProcess.preprocess(docx_bytes)
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "docx+styles", "-t", "latex", "--lua-filter", _FILTER, "-o", "-"],
        input=pre,
        capture_output=True,
        check=True,
    )
    return completed.stdout.decode()


def _real_widths(latex: str) -> list[float]:
    return [float(x) for x in re.findall(r"\\real\{([0-9.]+)\}", latex)]


def test_40pct_left_table_scaled_and_left_aligned():
    latex = _to_latex(_docx_with_table(2000, "left", cols=3))
    widths = _real_widths(latex)
    assert widths, "expected \\real{} column widths"
    assert abs(sum(widths) - 0.40) < 0.02  # 3 cols summing to ~40%
    assert "\\setlength{\\LTleft}{0pt}" in latex
    assert "\\setlength{\\LTright}{\\fill}" in latex


def test_25pct_right_table_right_aligned():
    latex = _to_latex(_docx_with_table(1250, "right", cols=1))
    widths = _real_widths(latex)
    assert abs(sum(widths) - 0.25) < 0.02
    assert "\\setlength{\\LTright}{0pt}" in latex


def test_25pct_center_table_centered():
    latex = _to_latex(_docx_with_table(1250, "center", cols=1))
    assert abs(sum(_real_widths(latex)) - 0.25) < 0.02
    # Centered uses \fill on both sides (also the reset value) and is never
    # pinned to an edge.
    assert "\\setlength{\\LTleft}{0pt}" not in latex
    assert "\\setlength{\\LTright}{0pt}" not in latex


def test_full_width_table_fills_line_and_is_left_aligned():
    """A 100% table keeps full width (columns sum to ~1.0) and is pinned
    flush-left, so it fills the text column from the left edge rather than
    floating content-width in the centre."""
    latex = _to_latex(_docx_with_table(5000, "left", cols=3))
    widths = _real_widths(latex)
    assert widths
    assert abs(sum(widths) - 1.0) < 0.05
    assert "\\setlength{\\LTleft}{0pt}" in latex
