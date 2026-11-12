"""End-to-end test for DOCX table width/alignment surviving into LaTeX.

Builds a real DOCX with a table carrying ``<w:tblW>`` / ``<w:jc>`` (what the
HTML->DOCX post-processing writes), runs it through the docx->latex
preprocessing + ``filters/docx/docx_tables_to_latex.lua``, and checks that:

* pandoc's DOCX reader would normalise the column widths to sum to 1.0, but the
  filter scales them back to the table's real page fraction (``\\dimexpr<share>\\linewidth...``);
* the table's alignment is re-applied as longtable ``\\LTleft``/``\\LTright`` glue.
"""

from __future__ import annotations

import io
import re
from typing import TYPE_CHECKING

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from app import docx_latex_pre_process
from tests.pandoc_in_container import filter_path, run_pandoc

if TYPE_CHECKING:
    from docker.models.containers import Container

    from tests.test_container import TestParameters

_FILTER = filter_path("docx/docx_tables_to_latex.lua")


def _docx_with_table(width_pct: int, jc: str, cols: int = 3) -> bytes:
    """A one-row DOCX table with the given tblW (pct) and jc alignment."""
    doc = Document()
    table = doc.add_table(rows=1, cols=cols)
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = f"c{i}"
    tblpr = table._tbl.tblPr
    # python-docx already emits a <w:tblW>; replace it (and any jc) so the table
    # has exactly one of each — mirroring app/docx_post_process.py's output.
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for tag in ("w:tblW", "w:jc"):
        for existing in tblpr.findall(tag, ns):
            tblpr.remove(existing)
    tblpr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{width_pct}" w:type="pct"/>'))
    tblpr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc}"/>'))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _to_latex(container: Container, docx_bytes: bytes, standalone: bool = False) -> str:
    """Run the docx->latex pipeline. `standalone` also returns the preamble."""
    pre = docx_latex_pre_process.preprocess(docx_bytes)
    arguments = ["-f", "docx+styles", "-t", "latex", f"--lua-filter={_FILTER}"]
    if standalone:
        arguments.append("-s")
    return run_pandoc(container, arguments, pre, "in.docx").decode()


def _line_shares(latex: str) -> list[float]:
    """The share of the line each column takes: the filter writes p{\\dimexpr<share>\\linewidth...}, margins included as Word counts them."""
    return [float(x) for x in re.findall(r"p\{\\dimexpr([0-9.]+)\\linewidth", latex)]


def test_40pct_left_table_scaled_and_left_aligned(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_table(2000, "left", cols=3))
    widths = _line_shares(latex)
    assert widths, "expected column widths as shares of the line"
    assert abs(sum(widths) - 0.40) < 0.02  # 3 cols summing to ~40%
    assert "\\setlength{\\LTleft}{0pt}" in latex
    assert "\\setlength{\\LTright}{\\fill}" in latex


def test_25pct_right_table_right_aligned(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_table(1250, "right", cols=1))
    widths = _line_shares(latex)
    assert abs(sum(widths) - 0.25) < 0.02
    assert "\\setlength{\\LTright}{0pt}" in latex


def test_25pct_center_table_centered(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_table(1250, "center", cols=1))
    assert abs(sum(_line_shares(latex)) - 0.25) < 0.02
    # Centered uses \fill on both sides (also the reset value) and is never
    # pinned to an edge.
    assert "\\setlength{\\LTleft}{0pt}" not in latex
    assert "\\setlength{\\LTright}{0pt}" not in latex


def test_full_width_table_fills_line_and_is_left_aligned(test_parameters: TestParameters):
    """A 100% table keeps full width (columns sum to ~1.0) and is pinned
    flush-left, so it fills the text column from the left edge rather than
    floating content-width in the centre."""
    latex = _to_latex(test_parameters.container, _docx_with_table(5000, "left", cols=3))
    widths = _line_shares(latex)
    assert widths
    assert abs(sum(widths) - 1.0) < 0.05
    assert "\\setlength{\\LTleft}{0pt}" in latex


# ---- Indented cells (Polarion's cell padding) ---------------------------


def _docx_with_shaded_indented_cells(indent_twips: int) -> bytes:
    """A shaded two-row table whose cell paragraphs carry a left indent.

    Polarion pads its table cells with a small <w:ind w:left="142"/>, and
    pandoc's DOCX reader reads any indented paragraph as a BlockQuote. Two
    rows, because pandoc reads the first as the header and wraps header cells
    in a \\minipage of its own accord — the body row is what shows whether the
    indent added one.
    """
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cell.text = f"cell{r}{c}"
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F2F2F2"/>'))
            p_pr = cell.paragraphs[0]._p.get_or_add_pPr()
            p_pr.append(parse_xml(f'<w:ind {nsdecls("w")} w:left="{indent_twips}"/>'))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _table_body(latex: str) -> str:
    """The rows after the longtable header/footer declarations."""
    start = latex.index("\\endlastfoot")
    return latex[start : latex.index("\\end{longtable}")]


def test_indented_cells_still_get_their_background(test_parameters: TestParameters):
    """The sentinel sits inside the BlockQuote pandoc makes from the indent.

    Read from the top-level blocks only, it was never found, so every indented
    cell silently lost its shading while an empty (unindented) one kept it.
    """
    latex = _to_latex(test_parameters.container, _docx_with_shaded_indented_cells(142))

    assert latex.count("\\cellcolor[HTML]{F2F2F2}") == 4, f"expected all four cells shaded, got:\n{latex}"


def test_indented_cells_are_not_wrapped_in_a_quote(test_parameters: TestParameters):
    """LaTeX's quote environment adds ~2.5em of margin on each side.

    Inside an already narrow column that pushes the text into wrapping and
    hyphenating where there was room for it, and the extra block makes the
    cell multi-block, which the writer renders through a \\minipage.
    """
    latex = _to_latex(test_parameters.container, _docx_with_shaded_indented_cells(142))

    assert "\\begin{quote}" not in latex, f"cell content still wrapped in a quote:\n{latex}"
    body = _table_body(latex)
    assert "\\begin{minipage}" not in body, f"body cell still wrapped in a minipage:\n{body}"


def test_unindented_cells_are_unaffected(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_shaded_indented_cells(0))

    assert latex.count("\\cellcolor[HTML]{F2F2F2}") == 4
    assert "\\begin{quote}" not in latex


# ---- Per-cell horizontal alignment --------------------------------------
#
# Word keeps alignment per paragraph, so a Polarion table routinely centres one
# cell in an otherwise left-aligned column. Pandoc's LaTeX writer reads only the
# COLUMN's alignment for an ordinary cell; up to 3.6 its DOCX reader hid that by
# folding the cells' alignment into the colspec, and 3.11 no longer does.


def _docx_with_cell_alignments(alignments: list[str | None], merge_first_column: bool = False) -> bytes:
    """A two-row table; row 1 gets the given per-cell alignments, row 2 none."""
    doc = Document()
    table = doc.add_table(rows=2, cols=len(alignments))
    for c, align in enumerate(alignments):
        cell = table.rows[0].cells[c]
        cell.text = f"top{c}"
        if align is not None:
            cell.paragraphs[0]._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{align}"/>'))
        table.rows[1].cells[c].text = f"bottom{c}"
    if merge_first_column:
        table.cell(0, 0).merge(table.cell(1, 0))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def test_centered_cell_emits_centering_in_the_cell(test_parameters: TestParameters):
    """The column stays left; only the one centred cell carries the switch."""
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments(["center", None]))

    assert _switch_on(latex, "top0") == "\\pdcCellCentering", f"centred cell did not get its alignment:\n{latex}"
    assert _switch_on(latex, "bottom0") != "\\pdcCellCentering", "alignment leaked onto an unaligned cell"


def test_alignment_switch_carries_arraybackslash(test_parameters: TestParameters):
    r"""Regression guard: \centering redefines \\.

    Without \arraybackslash the row-terminating \\ stops ending the row and the
    whole document fails to compile with "Extra alignment tab has been changed
    to \cr". The bare "\centering{}" form is what this filter emitted before
    the pairing was added; pandoc's own \centering inside a header minipage is
    fine and not what is checked here.
    """
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments(["center", "right", "left"]))

    for bare in ("\\centering{}", "\\raggedleft{}", "\\raggedright{}"):
        assert bare not in latex, f"{bare} emitted without \\arraybackslash in:\n{latex}"
    assert "\\pdcCellCentering{}" in latex


@pytest.mark.parametrize(
    ("jc", "expected"),
    [("center", "\\pdcCellCentering"), ("right", "\\pdcCellRaggedleft"), ("left", "\\pdcCellRaggedright")],
)
def test_each_alignment_maps_to_its_latex_switch(test_parameters: TestParameters, jc: str, expected: str):
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments([jc]))

    assert _switch_on(latex, "top0") == expected


def test_unaligned_cells_get_the_left_switch(test_parameters: TestParameters):
    """A cell with no <w:jc> is left-aligned in Word, and says so explicitly.

    It cannot stay silent: pandoc synthesises the column alignment from the
    other cells, and silence would mean inheriting it. See
    test_unaligned_cell_does_not_inherit_a_synthesised_column_alignment.
    """
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments([None, None]))

    assert _switch_on(latex, "top0") == "\\pdcCellRaggedright"
    assert "\\pdcCellCentering" not in latex
    assert "\\pdcCellRaggedleft" not in latex


def _docx_with_merged_body_cell() -> bytes:
    """Three rows, so the merge lands in the body rather than the header.

    Pandoc reads the first row as the header and wraps header cells in a
    minipage carrying its own \\centering, which would mask what this checks.
    """
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cell.text = f"r{r}c{c}"
    merged = table.cell(1, 0).merge(table.cell(2, 0))
    merged.paragraphs[0]._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def test_merged_cell_keeps_its_alignment(test_parameters: TestParameters):
    r"""A rowspan becomes \multirow, which takes no alignment argument.

    The column's alignment cannot reach it at all, so the cell's own switch is
    the only thing that can centre it.
    """
    latex = _to_latex(test_parameters.container, _docx_with_merged_body_cell())

    start = latex.index("\\multirow")
    assert "\\pdcCellCentering{}" in latex[start : start + 200], f"merged cell lost its alignment:\n{latex}"


def _docx_with_multi_paragraph_aligned_cell() -> bytes:
    r"""A centred cell holding two paragraphs.

    Pandoc renders a multi-block cell through a \minipage, where \\ is an
    ordinary line break rather than the row terminator.
    """
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cell.text = f"r{r}c{c} first"
    target = table.rows[1].cells[0]
    target.add_paragraph("second line")
    for para in target.paragraphs:
        para._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def test_alignment_switch_never_touches_the_row_terminator(test_parameters: TestParameters):
    r"""Regression: the switch must not redefine \\.

    \centering and friends also do \let\\\@centercr, and what \\ must mean
    depends on where the content lands - row terminator in a bare p{} cell,
    line break inside the minipage pandoc wraps a multi-line cell in. Pairing
    the switch with \arraybackslash fixes the first and breaks the second
    ("Extra }, or forgotten \endgroup"), so neither LaTeX built-in is usable
    and the filter emits its own glue-only macros instead.
    """
    latex = _to_latex(test_parameters.container, _docx_with_multi_paragraph_aligned_cell())
    # Strip pandoc's own colspec preamble; only what the filter emits matters,
    # and whether pandoc reaches for a minipage here is version-dependent.
    body = latex[latex.index("\\endlastfoot") :]

    assert "\\arraybackslash" not in body, f"the filter emitted \\arraybackslash, which breaks \\\\ inside a minipage:\n{body}"
    for builtin in ("\\centering{}", "\\raggedleft{}", "\\raggedright{}"):
        assert builtin not in body, f"{builtin} redefines \\\\ and must not be emitted:\n{body}"
    assert "\\pdcCellCentering{}" in body


def test_alignment_macros_are_defined_in_the_preamble(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments(["center"]), standalone=True)

    for macro in ("\\pdcCellCentering", "\\pdcCellRaggedleft", "\\pdcCellRaggedright"):
        assert f"\\providecommand{{{macro}}}" in latex, f"{macro} used but never defined"


def test_colortbl_is_not_loaded_for_a_table_with_no_cell_background(test_parameters: TestParameters):
    """colortbl redefines internal table macros, so it stays out when unused."""
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments(["center"]), standalone=True)

    assert "\\usepackage{colortbl}" not in latex, "colortbl loaded for a table that has no \\cellcolor"


def test_unaligned_cell_does_not_inherit_a_synthesised_column_alignment(test_parameters: TestParameters):
    """The heart of it: Word has no column alignment.

    A DOCX carries <w:jc> per paragraph only, and a paragraph without one is
    left-aligned. Pandoc's DOCX reader synthesises the COLUMN alignment from
    the cells it saw, so one centred cell can leave a whole column centred -
    and an unaligned cell in that column came out centred where Word shows it
    flush left.
    """
    latex = _to_latex(test_parameters.container, _docx_with_cell_alignments(["center", None]))

    assert _switch_on(latex, "top0") == "\\pdcCellCentering", "the centred cell lost its alignment"
    assert _switch_on(latex, "bottom0") == "\\pdcCellRaggedright", "the unaligned cell inherited the column's alignment"
    assert _switch_on(latex, "top1") == "\\pdcCellRaggedright"


def test_empty_cells_get_no_alignment_switch(test_parameters: TestParameters):
    """Nothing to align, and injecting would give the cell a paragraph."""
    doc = Document()
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "filled"
    table.rows[0].cells[0].paragraphs[0]._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
    buffer = io.BytesIO()
    doc.save(buffer)

    latex = _to_latex(test_parameters.container, buffer.getvalue())

    assert latex.count("\\pdcCell") - latex.count("\\pdcCellPos") == 1, f"expected one switch, for the one non-empty cell:\n{latex}"


def test_blank_shaded_cell_gets_no_alignment_switch(test_parameters: TestParameters):
    """A cell that renders nothing is left alone, background and all.

    consume_sentinel strips the sentinel text but keeps the paragraph that
    held it, so a shaded spacer cell arrives as a single empty Plain. Counting
    blocks would see content there and inject a switch into an empty cell.
    """
    doc = Document()
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0]._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F2F2F2"/>'))
    table.rows[0].cells[1].text = "real"
    buffer = io.BytesIO()
    doc.save(buffer)

    latex = _to_latex(test_parameters.container, buffer.getvalue())

    assert latex.count("\\pdcCell") - latex.count("\\pdcCellPos") == 1, f"the blank shaded cell was given a switch:\n{latex}"
    assert "\\cellcolor[HTML]{F2F2F2}" in latex, "the blank cell lost its background"


# ---- Justified cells ----------------------------------------------------


def _docx_with_justified_cell(jc: str | None, shaded: bool = False) -> bytes:
    """A two-row table whose body cell carries the given <w:jc>."""
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cell.text = "Justified body text long enough to wrap" if (r, c) == (1, 0) else f"r{r}c{c}"
    cell = table.rows[1].cells[0]
    if jc is not None:
        cell.paragraphs[0]._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc}"/>'))
    if shaded:
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F2F2F2"/>'))
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _switch_on(latex: str, needle: str) -> str | None:
    """The alignment switch of the cell which starts with needle. A cell of the header row has the strut of a row after it."""
    match = re.search(r"(\\pdcCell\w+)\{\}(?:\\pdcHeadStrut\{\})?" + re.escape(needle), latex)
    return match.group(1) if match else None


@pytest.mark.parametrize("jc", ["both", "distribute"])
def test_justified_cell_keeps_its_justification(test_parameters: TestParameters, jc):
    r"""Pandoc's Alignment has no justified value.

    <w:jc w:val="both"> reaches the reader as AlignLeft and "distribute" as
    AlignDefault, so the Cell cannot carry it and every cell would be flushed
    left. It rides in on the sentinel instead.
    """
    latex = _to_latex(test_parameters.container, _docx_with_justified_cell(jc))

    assert _switch_on(latex, "Justified") == "\\pdcCellJustify"


@pytest.mark.parametrize(
    ("jc", "expected"),
    [("left", "\\pdcCellRaggedright"), ("center", "\\pdcCellCentering"), ("right", "\\pdcCellRaggedleft"), (None, "\\pdcCellRaggedright")],
)
def test_non_justified_cells_are_unaffected(test_parameters: TestParameters, jc, expected):
    latex = _to_latex(test_parameters.container, _docx_with_justified_cell(jc))

    assert _switch_on(latex, "Justified") == expected


def test_justified_cell_keeps_its_background_too(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_justified_cell("both", shaded=True))

    assert _switch_on(latex, "Justified") == "\\pdcCellJustify"
    assert "\\cellcolor[HTML]{F2F2F2}" in latex


def test_justify_macro_is_defined_in_the_preamble(test_parameters: TestParameters):
    latex = _to_latex(test_parameters.container, _docx_with_justified_cell("both"), standalone=True)

    assert "\\providecommand{\\pdcCellJustify}" in latex
