"""Integration tests for the lines of a table in ``filters/docx/docx_tables_to_latex.lua``.

app/docx_table_pre_process.py tags the first cell of a styled table with
``rl=none`` or ``rl=head``; the filter sets the table without the booktabs
rules Word does not draw. The tests run pandoc inside the container on a
markdown table whose first cell carries the sentinel.
"""

from __future__ import annotations

import base64
import ctypes
import io
import itertools
import re
import zipfile
from xml.etree import ElementTree as ET

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c
import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor, Twips
from PIL import Image
from pypdf import PdfReader

from app import docx_table_pre_process
from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_tables_to_latex.lua"
OPEN, CLOSE = "", ""


def _latex(container, payload: str) -> str:
    markdown = f"| {OPEN}{payload}{CLOSE}Head | Other |\n|---|---|\n| a | b |\n"
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/rules.md << 'EOF'\n{markdown}\nEOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/rules.md"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode()


def test_a_table_without_lines_drops_every_rule(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "rl=none")
    assert latex.startswith("\\begingroup\\def\\toprule{}\\def\\midrule{}\\def\\bottomrule{}")
    assert latex.rstrip().endswith("\\endgroup")
    assert "rl=none" not in latex


def test_a_table_with_a_header_line_keeps_the_rule_under_the_header(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "rl=head")
    assert latex.startswith("\\begingroup\\def\\toprule{}\\def\\bottomrule{}")
    assert "\\def\\midrule{}" not in latex


def test_an_unknown_kind_of_lines_is_left_out(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "rl=dashed")
    assert "\\begingroup" not in latex
    assert "\\toprule" in latex


def test_a_table_of_compact_paragraphs_takes_the_rows_of_compact(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "cc=1")
    assert latex.startswith("\\begingroup\\ifdefined\\pdcCompactCellStretch\\renewcommand{\\arraystretch}{\\pdcCompactCellStretch}\\setlength{\\extrarowheight}{\\pdcCompactCellExtra}\\fi")
    assert latex.count("\\begingroup") == latex.count("\\endgroup")


GRID = "rl=grid;lc=CCCCCC;lw=0.75"


def _native_latex(container, native: str) -> str:
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/grid.native << 'EOF'\n{native}\nEOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/grid.native"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode()


def _cell(content: str, rows: int = 1) -> str:
    return f'Cell ("",[],[]) AlignDefault (RowSpan {rows}) (ColSpan 1) [Plain [{content}]]'


def _grid_table(*rows: list[str]) -> str:
    body = ", ".join(f'Row ("",[],[]) [{", ".join(cells)}]' for cells in rows)
    columns = ", ".join(["(AlignDefault, ColWidth 0.5)"] * 2)
    return f'[Table ("",[],[]) (Caption Nothing []) [{columns}] (TableHead ("",[],[]) []) [TableBody ("",[],[]) (RowHeadColumns 0) [] [{body}]] (TableFoot ("",[],[]) [])]'


def test_a_grid_draws_a_line_around_every_cell(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, GRID)
    assert "\\begin{longtable}[]{!{\\pdcGridV{}}" in latex
    assert latex.count("!{\\pdcGridV{}}") >= 3
    assert "\\setlength{\\arrayrulewidth}{0.75pt}\\arrayrulecolor[HTML]{CCCCCC}" in latex
    assert "\\\\\\hline" in latex
    assert "pdcGridRowEnd" not in latex
    assert latex.count("\\begingroup") == latex.count("\\endgroup")


def test_a_merged_cell_keeps_no_line_inside_it(test_parameters: TestParameters):
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"', rows=2), _cell('Str "b"')], [_cell('Str "c"')])
    latex = _native_latex(test_parameters.container, native)
    assert "\\multirow[t]{2}" in latex, "Word sets the text of a merged cell at its top"
    assert "\\cline{2-2}" in latex


def test_a_grid_with_a_picture_draws_its_lines_and_keeps_the_picture(test_parameters: TestParameters):
    """pandoc writes out the file of a picture, a GIF as PNG, only for an Image it sees: the picture stays one."""
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell('Image ("",[],[("width","0.2in")]) [] ("x.gif","")')])
    latex = _native_latex(test_parameters.container, native)
    assert "\\begin{longtable}[]{!{\\pdcGridV{}}" in latex
    assert "\\includegraphics[width=0.2in" in latex
    # The document after the filter, as the writer of LaTeX gets it
    dump = "function Pandoc(doc) io.stderr:write(pandoc.write(doc, 'native')) end"
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"cat > /tmp/test/dump.lua << 'EOF'\n{dump}\nEOF"])
    command = f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} --lua-filter=/tmp/test/dump.lua /tmp/test/grid.native 2>&1 >/dev/null"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    after = " ".join(output.decode().split())
    assert '( "x.gif" , "" )' in after
    assert "Table (" not in after


def test_the_text_of_a_cell_without_a_style_stands_at_its_top(test_parameters: TestParameters):
    """Word sets the space of Normal below the text of the cell: it goes into the depth of the cell."""
    latex = _latex(test_parameters.container, "tw=1.0000")
    assert latex.count("\\rule[-\\dimexpr\\dp\\strutbox+\\pdcPlainCellSpace\\relax]{0pt}{0pt}") == 2, "the header row"
    assert latex.count("\\pdcWrapEnd{\\dimexpr\\dp\\strutbox+\\pdcPlainCellSpace\\relax}") == 2, "the body row, which goes out a line at a time"


def test_a_table_of_compact_paragraphs_gets_no_space_below_its_text(test_parameters: TestParameters):
    assert "pdcPlainCellSpace" not in _latex(test_parameters.container, "cc=1")


def test_a_caption_under_a_table_takes_the_space_before_of_its_style(test_parameters: TestParameters):
    native = _grid_table([_cell('Str "a"'), _cell('Str "b"')])[:-1] + ', Div ("",[],[("custom-style","PandocCaption")]) [Para [Str "Table 1"]], Div ("",[],[("custom-style","PandocCaption")]) [Para [Str "Other"]]]'
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("\\vspace{\\dimexpr\\pdcCaptionBefore-\\parskip\\relax}") == 1, "only the caption right under the table"


def test_a_caption_above_a_table_takes_the_space_after_of_its_style(test_parameters: TestParameters):
    table = _grid_table([_cell('Str "a"'), _cell('Str "b"')])
    native = '[Div ("",[],[("custom-style","PandocCaption")]) [Para [Str "Table 1"]], Div ("",[],[("custom-style","PandocCaption")]) [Para [Str "Other"]], ' + table[1:]
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("\\vspace{\\dimexpr\\pdcCaptionAfter-\\parskip\\relax}") == 1, "only the caption right above the table"
    assert latex.index("pdcCaptionAfter") > latex.index("Other")


def test_a_cell_of_compact_paragraphs_ends_as_deep_as_its_row(test_parameters: TestParameters):
    """The last line of a cell of several lines stands as far above the bottom of the row as a single line."""
    latex = _latex(test_parameters.container, "cc=1")
    assert latex.count("\\ifdefined\\pdcCompactCellDepth\\pdcCompactCellDepth\\fi") == 4
    assert "pdcCompactCellBottom" not in latex, "the row holds the space below its last line"


def test_a_grid_keeps_the_curly_quotes_of_its_text(test_parameters: TestParameters):
    """The fonts of the PDF set no TeX ligatures: TeX quotes of the smart extension would print as backticks."""
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}“a”"'), _cell('Str "b"')])
    latex = _native_latex(test_parameters.container, native)
    assert "“a”" in latex
    assert "``" not in latex


def test_a_column_takes_the_width_of_word_with_the_cell_margins_in_it(test_parameters: TestParameters):
    """Word sets a table as wide as the line out to the margins of its cells, and a column holds them on both sides."""
    native = _grid_table([_cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}a"'), _cell('Str "b"')])
    latex = _native_latex(test_parameters.container, native)
    assert "p{\\dimexpr0.5000\\linewidth+0.5000\\tabcolsep+0.5000\\tabcolsep-2\\tabcolsep\\relax}" in latex
    assert "\\real{" not in latex


def test_a_column_of_struck_text_takes_the_width_of_word(test_parameters: TestParameters):
    """A deleted linked work item is struck through: its cell is no reason to leave the widths to pandoc."""
    native = _grid_table([_cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}a"'), _cell('Strikeout [Str "b"]')])
    latex = _native_latex(test_parameters.container, native)
    assert "p{\\dimexpr0.5000\\linewidth+0.5000\\tabcolsep+0.5000\\tabcolsep-2\\tabcolsep\\relax}" in latex
    assert "\\real{" not in latex


def test_a_table_in_a_cell_stays_a_table_for_the_filter_of_nested_tables(test_parameters: TestParameters):
    inner = _grid_table([_cell(f'Str "{OPEN}tw=0.5000{CLOSE}x"'), _cell('Str "y"')])[1:-1]
    outer = _grid_table([_cell(f'Str "{OPEN}tw=1.0000{CLOSE}a"'), _cell('Str "b"')])
    native = outer.replace('[Plain [Str "b"]]', f"[{inner}]")
    latex = _native_latex(test_parameters.container, native)
    assert "\\real{" in latex, "pandoc writes the outer table, with the inner one in it"


def test_a_grid_in_a_cell_is_set_as_a_tabular_in_the_grid_of_the_outer_table(test_parameters: TestParameters):
    """Word draws the lines of both tables. A longtable cannot stand in the minipage of a cell: the inner one is a tabular."""
    inner = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}x"'), _cell('Str "y"')])[1:-1]
    outer = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell('Str "b"')])
    native = outer.replace('[Plain [Str "b"]]', f"[{inner}]")
    latex = _native_latex(test_parameters.container, native)
    assert "\\real{" not in latex, "both tables are grids, written as LaTeX"
    outer_start = latex.index("\\begin{longtable}")
    inner_start = latex.index("\\begin{longtable}", outer_start + 1)
    assert latex.rfind("\\def\\longtable{", 0, inner_start) > outer_start, "the inner longtable is a tabular"
    # Word sets a table in a cell from the start of the line of the cell, as wide as that line
    assert latex.rfind("\\setlength{\\LTleft}{0pt}", 0, inner_start) > outer_start
    assert "\\linewidth-2\\tabcolsep\\relax}" in latex[inner_start:]


def test_a_row_of_line_breaks_becomes_a_row_per_line_which_a_page_may_break(test_parameters: TestParameters):
    """Word breaks a row between the lines of a cell; longtable only between rows."""
    lines = 'Plain [Str "one", LineBreak, Str "two", LineBreak, Str "three"]'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}Label"'), f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [{lines}]'])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\pdcSplitRow" not in latex
    assert latex.index("one") < latex.index("\\\\\\pdcSplitFirstMiddle") < latex.index("two") < latex.index("\\\\\\pdcSplitLast")
    assert latex.index("three") < latex.index("\\\\\\pdcSplitEnd \\hline")
    assert latex.count("\\hline") == 2, "a line above the table and under its last row of lines only"
    assert "\\raisebox{0pt}[\\height][0pt]{\\begin{minipage}[t]{\\linewidth}" in latex, "the label runs on beside the lines"


def test_a_row_of_two_cells_of_line_breaks_stays_one_row(test_parameters: TestParameters):
    lines = 'Plain [Str "one", LineBreak, Str "two"]'
    cell = f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [{lines}]'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a", LineBreak, Str "b"'), cell])
    assert "pdcSplit" not in _native_latex(test_parameters.container, native)


def test_a_picture_alone_on_its_line_hangs_from_the_top_of_the_cell(test_parameters: TestParameters):
    """Word sets it from the top of its line down: TeX lines cells up at the bottom of such a picture."""
    picture = 'Image ("",[],[("width","1in"),("height","1in")]) [] ("x.png","")'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell(f'Str "Label:", LineBreak, {picture}')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "Label: \\\\ \\pdcCellPicture{\\includegraphics" in latex or "\\pdcCellPicture{\\includegraphics" in latex
    assert latex.count("pdcCellPicture{") == 1, "the label stays as it is"


def test_a_picture_of_a_cell_keeps_the_space_of_the_row_below_it(test_parameters: TestParameters):
    """Word sets it without the descent of the font, and below it the space the cell has below a line."""
    picture = 'Image ("",[],[("width","1in"),("height","1in")]) [] ("x.png","")'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell(picture)])
    _native_latex(test_parameters.container, native)
    exit_code, output = test_parameters.container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/grid.native"])
    assert exit_code == 0, output.decode()
    rule = "\\ifdim\\dp\\@arstrutbox>\\dp\\strutbox\\rule[\\dimexpr\\ht\\strutbox-\\ht\\pdcCellPictureBox-\\dp\\@arstrutbox+\\dp\\strutbox\\relax]{0pt}{0pt}\\fi"
    assert rule in output.decode()


def test_a_coloured_cell_merged_down_colours_the_rows_it_spans(test_parameters: TestParameters):
    """colortbl colours a \\multirow in its first row only, and pandoc writes nothing in the rows below."""
    color = 'RawInline (Format "latex") "\\\\cellcolor[HTML]{F4CCCC}"'
    merged = f'Cell ("",[],[]) AlignDefault (RowSpan 2) (ColSpan 1) [Plain [{color}, Str "Category"]]'
    native = _grid_table([merged, _cell(f'Str "{OPEN}{GRID};tw=1.0000{CLOSE}a"')], [_cell('Str "b"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\cellcolor[HTML]{F4CCCC}\\multirow[t]{2}{=}{\\pdcCellRaggedright{}Category" in latex, latex
    assert latex.count("\\cellcolor[HTML]{F4CCCC}") == 2, "the row below gets a cell of the colour"
    assert "\\cline{2-2}" in latex, "no line inside the merged cell"


def test_a_merged_cell_at_the_edge_of_a_table_keeps_its_margin_inside_its_lines(test_parameters: TestParameters):
    """Without a margin at the edge, the colour of the cell would reach past the line of the table."""
    cell = f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 2) [Plain [Str "{OPEN}{GRID};tw=1.0000{CLOSE}Title"]]'
    native = _grid_table([cell], [_cell('Str "a"'), _cell('Str "b"')])
    latex = _native_latex(test_parameters.container, native)
    assert "\\multicolumn{2}{!{\\pdcGridV{}}>" in latex, latex
    assert "-2\\tabcolsep\\relax}!{\\pdcGridV{}}}" in latex, latex


def _bordered_pdf(test_parameters: TestParameters, eighths: int) -> bytes:
    """A table of two shaded cells with black borders of the width given in eighths of a point."""
    document = Document()
    table = document.add_table(rows=1, cols=2)
    for cell, text in zip(table.rows[0].cells, ("Left", "Right"), strict=True):
        cell.text = text
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="FF0000"/>'))
    sides = "".join(f'<w:{side} w:val="single" w:sz="{eighths}" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    properties = table._tbl.tblPr
    properties.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    properties.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>'))
    docx = io.BytesIO()
    document.save(docx)
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), mime)})
    assert response.status_code == 200, response.text
    return response.content


def _text_left(pdf: bytes, text: str) -> float:
    page = pdfium.PdfDocument(pdf)[0]
    found = page.get_textpage().search(text, match_case=True).get_next()
    assert found, text
    return page.get_textpage().get_charbox(found[0])[0]


def test_the_vertical_lines_of_a_grid_take_no_room_from_the_text(test_parameters: TestParameters):
    """Word draws a vertical border over the edge between two cells: the text stands where it would without it."""
    thin, thick = _bordered_pdf(test_parameters, 4), _bordered_pdf(test_parameters, 48)
    for text in ("Left", "Right"):
        assert abs(_text_left(thick, text) - _text_left(thin, text)) < 0.1, text


def test_the_colour_of_a_cell_leaves_the_vertical_line_whole(test_parameters: TestParameters):
    """colortbl paints a cell over the half of the line before it which lies in the cell."""
    pdf = _bordered_pdf(test_parameters, 48)
    page = pdfium.PdfDocument(pdf)[0]
    textpage = page.get_textpage()
    found = textpage.search("Right", match_case=True).get_next()
    left, bottom, _, top = textpage.get_charbox(found[0])
    scale = 4
    image = page.render(scale=scale).to_pil().convert("RGB")
    row = (page.get_height() - (bottom + top) / 2) * scale
    # The widest black run on the row between the two words: the line between the cells, 6 points wide
    run, widest = 0, 0
    for x in range(round(_text_left(pdf, "Left") * scale), round(left * scale)):
        red, green, blue = image.getpixel((x, round(row)))
        run = run + 1 if red < 80 and green < 80 and blue < 80 else 0
        widest = max(widest, run)
    assert widest >= 5.5 * scale, widest / scale


def test_the_line_under_a_header_row_is_the_border_of_word_over_the_rows(test_parameters: TestParameters):
    """Word draws a border of the default width, 2 eighths of a point, and the rows keep their pitch."""
    latex = _latex(test_parameters.container, "rl=head")
    rule = "\\def\\midrule{\\noalign{\\kern-0.125pt}\\multispan{\\csname LT@cols\\endcsname}\\kern-\\tabcolsep\\leaders\\hrule height 0.25pt\\hfill\\kern-\\tabcolsep\\cr\\noalign{\\kern-0.125pt}}"
    assert rule in latex, "as wide as the table, not the line"


def test_a_coloured_cell_merged_down_with_line_breaks_stays_one_cell(test_parameters: TestParameters):
    """Its \\multirow opens in its first row and closes after its text: a row per line would part the braces."""
    color = 'RawInline (Format "latex") "\\\\cellcolor[HTML]{F4CCCC}"'
    merged = f'Cell ("",[],[]) AlignDefault (RowSpan 2) (ColSpan 1) [Plain [{color}, Str "one", LineBreak, Str "two"]]'
    native = _grid_table([merged, _cell(f'Str "{OPEN}{GRID};tw=1.0000{CLOSE}a"')], [_cell('Str "b"')])
    latex = _native_latex(test_parameters.container, native)
    assert "pdcSplit" not in latex
    assert "\\multirow[t]{2}{=}{one\\\\" in " ".join(latex.split()).replace("\\pdcCellRaggedright{}", "")


def test_a_cell_of_the_header_row_starts_its_first_line_at_the_height_of_a_row(test_parameters: TestParameters):
    """pandoc sets a header cell on its last line: two lines of it lost the space Word puts above the first one."""
    first = _cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}Account", LineBreak, Str "type"')
    head = f'TableHead ("",[],[]) [Row ("",[],[]) [{first}, {_cell('Str "Rights"')}]]'
    native = _grid_table([_cell('Str "a"'), _cell('Str "b"')]).replace('(TableHead ("",[],[]) [])', f"({head})")
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("\\pdcCellRaggedright{}\\pdcHeadStrut{}") == 2, latex
    exit_code, output = test_parameters.container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/grid.native"])
    assert exit_code == 0, output.decode()
    assert "\\providecommand{\\pdcHeadStrut}{\\rule{0pt}{\\ht\\@arstrutbox}}" in output.decode()


def test_a_header_cell_which_starts_with_a_list_gets_no_strut_line(test_parameters: TestParameters):
    """Before a list the strut would be an empty line of its own, which Word does not draw."""
    first = _cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}Account"')
    listed = 'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [BulletList [[Plain [Str "a"]]]]'
    head = f'TableHead ("",[],[]) [Row ("",[],[]) [{first}, {listed}]]'
    native = _grid_table([_cell('Str "a"'), _cell('Str "b"')]).replace('(TableHead ("",[],[]) [])', f"({head})")
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("\\pdcHeadStrut{}") == 1, latex


def test_the_last_line_of_a_cell_is_as_deep_as_the_font_of_its_text(test_parameters: TestParameters):
    """The run closes its group before the end of the cell: the space below it measures the strut of the run."""
    strut = "\\\\setbox\\\\strutbox\\\\hbox{\\\\vrule height 13.1483bp depth 3.9445bp width 0pt}"
    run = f'RawInline (Format "latex") "{{\\\\fontsize{{14bp}}{{16.8bp}}\\\\selectfont {strut}", Str "Large", RawInline (Format "latex") "}}"'
    native = _grid_table([_cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}Body"'), _cell(run)])
    latex = _native_latex(test_parameters.container, native)
    expected = "Large}\\setbox\\strutbox\\hbox{\\vrule height 13.1483bp depth 3.9445bp width 0pt}\\ifdefined\\pdcPlainCellSpace"
    assert expected in latex, latex
    assert latex.count("depth 3.9445bp") == 2, "a cell of the body text keeps the strut of the body"


def test_a_cell_which_ends_after_its_run_ends_as_deep_as_the_body(test_parameters: TestParameters):
    """A run of 14 points, then text of the body: the last line is of the body, not of the run."""
    strut = "\\\\setbox\\\\strutbox\\\\hbox{\\\\vrule height 13.1483bp depth 3.9445bp width 0pt}"
    run = f'RawInline (Format "latex") "{{\\\\fontsize{{14bp}}{{16.8bp}}\\\\selectfont {strut}", Str "Large", RawInline (Format "latex") "}}"'
    native = _grid_table([_cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}Body"'), _cell(f'{run}, Space, Str "small"')])
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("depth 3.9445bp") == 1, latex


def test_a_cell_which_ends_in_a_run_inside_text_ends_as_deep_as_the_run(test_parameters: TestParameters):
    strut = "\\\\setbox\\\\strutbox\\\\hbox{\\\\vrule height 13.1483bp depth 3.9445bp width 0pt}"
    run = f'RawInline (Format "latex") "{{\\\\fontsize{{14bp}}{{16.8bp}}\\\\selectfont {strut}", Str "Large", RawInline (Format "latex") "}}"'
    native = _grid_table([_cell(f'Str "{OPEN}tw=1.0000;rl=none{CLOSE}Body"'), _cell(f'Str "small", Space, {run}')])
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("depth 3.9445bp") == 2, latex


def test_a_row_which_goes_out_a_line_at_a_time_starts_at_the_line_of_its_text(test_parameters: TestParameters):
    """Its first line is as high as the font of its cells, not of the body, as in any other row."""
    strut = "\\\\setbox\\\\strutbox\\\\hbox{\\\\vrule height 10.3308bp depth 3.0993bp width 0pt}"

    def run(text: str) -> str:
        return f'RawInline (Format "latex") "{{\\\\fontsize{{11bp}}{{13.2bp}}\\\\selectfont {strut}", Str "{text}", RawInline (Format "latex") "}}"'

    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}", {run("Label")}'), _cell(run("Text"))])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\pdcWrapBegin" in latex, latex
    assert "\\pdcRowLine{10.3308bp}" in latex, latex


def test_a_row_of_paragraphs_goes_out_a_line_at_a_time_which_a_page_may_break(test_parameters: TestParameters):
    """Word breaks a row between any two lines of its cells; longtable only between rows."""
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}Label"'), _cell('Str "a", Space, Str "long", Space, Str "text"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\pdcWrapBegin\\pdcWrapCell{1}{\\pdcCellRaggedright{}Label" in latex, latex
    assert "\\pdcWrapCellLast{2}{\\pdcCellRaggedright{}a long text" in latex, latex
    assert "\\\\\\pdcWrapRows{\\pdcWrapPiece{1}&\\pdcWrapPieceLast{2}}\\hline" in latex, "the line under the row after its last line"
    assert "\\pdcWrapRowEnd" not in latex


def test_a_coloured_cell_keeps_its_colour_in_every_line_of_its_row(test_parameters: TestParameters):
    native = _grid_table([_cell(f'Str "{OPEN}{GRID};bg=F4CCCC{CLOSE}Label"'), _cell('Str "text"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\pdcWrapBegin\\cellcolor[HTML]{F4CCCC}\\pdcWrapCell{1}{" in latex, latex
    assert "\\pdcWrapRows{\\cellcolor[HTML]{F4CCCC}\\pdcWrapPiece{1}&" in latex, latex


def test_a_row_with_a_picture_stays_one_row(test_parameters: TestParameters):
    picture = 'Image ("",[],[("width","1in"),("height","1in")]) [] ("x.png","")'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell(picture)])
    assert "pdcWrap" not in _native_latex(test_parameters.container, native)


def test_a_row_with_a_picture_in_its_text_goes_out_a_line_at_a_time(test_parameters: TestParameters):
    """An icon in a line of text keeps to its line: Word breaks the row between its lines all the same."""
    picture = 'Image ("",[],[("width","0.16667in"),("height","0.16667in")]) [] ("x.png","")'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell(f'Str "has", Space, {picture}, Space, Str "parent"')])
    assert "pdcWrap" in _native_latex(test_parameters.container, native)


def test_a_row_with_a_picture_alone_on_a_line_stays_one_row(test_parameters: TestParameters):
    picture = 'Image ("",[],[("width","1in"),("height","1in")]) [] ("x.png","")'
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell(f'Str "Label:", LineBreak, Span ("",[],[]) [{picture}]')])
    assert "pdcWrap" not in _native_latex(test_parameters.container, native)


def test_a_row_longer_than_the_rest_of_the_page_breaks_between_its_lines(test_parameters: TestParameters):
    document = Document()
    document.add_paragraph("Before the table.")
    # pandoc reads the first row as the header, which repeats on each page and does not break
    table = document.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = "Name"
    table.rows[0].cells[1].text = "Text"
    table.rows[1].cells[0].text = "Label"
    table.rows[1].cells[1].text = " ".join(f"w{number}" for number in range(700))
    properties = table._tbl.tblPr
    for width in properties.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblW"):
        properties.remove(width)
    properties.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>'))
    docx = io.BytesIO()
    document.save(docx)
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), mime)})
    assert response.status_code == 200, response.text
    pages = [page.extract_text() for page in PdfReader(io.BytesIO(response.content)).pages]
    assert "Before the table." in pages[0]
    assert "w0 " in pages[0] or "w0\n" in pages[0], "the row starts on the page of the text before it"
    assert "w699" in pages[-1]
    assert len(pages) > 1
    words = [word for word in " ".join(pages).split() if word.startswith("w") and word[1:].isdigit()]
    assert words == [f"w{number}" for number in range(700)], "every line once, in its order"


def test_a_row_of_cells_with_line_breaks_goes_out_a_line_at_a_time(test_parameters: TestParameters):
    """pandoc would set such a cell in a minipage, before which a row of its lines could not start: the box of the cell breaks its lines."""
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a", LineBreak, Str "b"'), _cell('Str "c", LineBreak, Str "d"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "minipage" not in latex, latex
    assert "\\pdcWrapBegin\\pdcWrapCell{1}{\\pdcCellRaggedright{}a\\strut\\newline{}b" in latex, latex


def test_a_row_under_a_cell_merged_down_stays_one_row(test_parameters: TestParameters):
    """Its first cell is not the first column: a row of its lines could not start there."""
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}merged"', rows=2), _cell('Str "a"')], [_cell('Str "b"')])
    latex = _native_latex(test_parameters.container, native)
    assert latex.count("\\pdcWrapBegin") == 0, latex


def _bordered(sides: str) -> str:
    return f"bd={sides}"


def test_the_cells_of_a_grid_draw_lines_of_their_own(test_parameters: TestParameters):
    """Word draws each side of a cell in its own style, width and color; a line between rows is as high as its widest piece."""
    first = f'Str "{OPEN}{GRID};{_bordered("single.8.FF0000/single.8.000000/double.12.0000FF/single.8.000000")}{CLOSE}a"'
    second = f'Str "{OPEN}{_bordered("single.8.FF0000/single.8.000000/dashed.8.00FF00/single.8.000000")}{CLOSE}b"'
    third = f'Str "{OPEN}{_bordered("single.4.000000/single.8.000000/single.4.000000/single.8.000000")}{CLOSE}c"'
    fourth = f'Str "{OPEN}{_bordered("single.4.000000/single.8.000000/single.4.000000/single.8.000000")}{CLOSE}d"'
    latex = " ".join(_native_latex(test_parameters.container, _grid_table([_cell(first), _cell(second)], [_cell(third), _cell(fourth)])).split())
    assert "\\begin{longtable}[]{!{\\pdcGridV{0}}" in latex, latex
    assert "!{\\pdcGridV{2}}}" in latex, "the separators are numbered"
    assert "\\pdcGVset{0}{1.0000}{000000}" in latex, "each row sets its vertical lines"
    assert "{\\color[HTML]{FF0000}\\leaders\\hrule height1.0000pt" in latex, "the line over the table in the color of its cells"
    assert "\\vrule height4.5000pt width0pt" in latex, "a double line of 1.5 points takes 4.5 points of height"
    assert "{\\color[HTML]{0000FF}\\hrule height1.5000pt" in latex
    assert "\\xleaders" in latex and "{\\color[HTML]{00FF00}\\hrule height1.0000pt width4.0000pt}" in latex, "a dashed line of 1 point: dashes of 4"
    table = latex[latex.index("\\begin{longtable}") :]
    assert "\\toprule" not in table and "\\hline" not in table, "the lines of the cells take the place of the rules"


def test_a_merged_cell_of_a_grid_numbers_the_separators_of_its_columns(test_parameters: TestParameters):
    sides = _bordered("single.8.000000/single.8.000000/single.8.000000/single.8.FF0000")
    title = f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 2) [Plain [Str "{OPEN}{GRID};{sides}{CLOSE}Title"]]'
    native = _grid_table([title], [_cell(f'Str "{OPEN}{sides}{CLOSE}a"'), _cell(f'Str "{OPEN}{sides}{CLOSE}b"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\multicolumn{2}{!{\\pdcGridV{0}}" in latex, latex
    assert "!{\\pdcGridV{2}}}{" in latex, latex
    assert "pdcFirstColumn" not in latex


def _pdf_of(test_parameters: TestParameters, document: Document) -> bytes:
    docx = io.BytesIO()
    document.save(docx)
    return _pdf_of_docx(test_parameters, docx.getvalue())


def _pdf_of_docx(test_parameters: TestParameters, docx: bytes) -> bytes:
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx, mime)})
    assert response.status_code == 200, response.text
    return response.content


def _two_rows(eighths: int) -> Document:
    """A grid of two rows whose first row ends in a line of the width given in eighths of a point."""
    document = Document()
    table = document.add_table(rows=2, cols=1)
    table.rows[0].cells[0].text = "Upper"
    table.rows[1].cells[0].text = "Lower"
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    properties = table.rows[0].cells[0]._tc.get_or_add_tcPr()
    properties.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="{eighths}" w:space="0" w:color="FF0000"/></w:tcBorders>'))
    return document


def _top(pdf: bytes, text: str) -> float:
    page = pdfium.PdfDocument(pdf)[0]
    found = page.get_textpage().search(text, match_case=True).get_next()
    assert found, text
    return page.get_height() - page.get_textpage().get_charbox(found[0])[3]


def test_a_wide_line_between_rows_adds_its_width_to_the_table(test_parameters: TestParameters):
    """Word sets the row under a line of 6 points lower by the width of the line."""
    thin, thick = _pdf_of(test_parameters, _two_rows(8)), _pdf_of(test_parameters, _two_rows(48))
    assert abs(_top(thick, "Upper") - _top(thin, "Upper")) < 0.1
    assert abs((_top(thick, "Lower") - _top(thin, "Lower")) - 5) < 0.3


def _head_of_two_lines(alignment: str | None) -> Document:
    """A table whose header row has a cell of two lines beside cells of one line, aligned as given; #372."""
    document = Document()
    document.add_paragraph("A table whose header row has a cell of two lines.")
    table = document.add_table(rows=3, cols=3)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    widths = (1800, 3600, 3626)
    for column, width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths, strict=True):
        column.set(qn("w:w"), str(width))
    rows = (("Requirement of the system", "Status", "Owner"), ("REQ-1 Stores every change", "Approved", "Admin"), ("REQ-2", "Draft", "Editor"))
    for row, texts in zip(table.rows, rows, strict=True):
        for cell, text, width in zip(row.cells, texts, widths, strict=True):
            cell.text = text
            cell.width = Twips(width)
    if alignment:
        for cell in table.rows[0].cells:
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{alignment}"/>'))
    return document


def test_a_table_in_a_header_cell_leaves_the_alignment_of_the_cell(test_parameters: TestParameters):
    """Word sets "Second header cell" on the last line of the header cell beside it, both at the bottom; #372.

    That header cell holds a table whose header stands at the top. The position of the inner cell took the place
    of the outer one, and the second cell stood 75pt high, at the top of the row.
    """
    document = Document()
    table = document.add_table(rows=2, cols=2)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    head = table.rows[0].cells
    for cell in head:
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="bottom"/>'))
    head[0].text = "Outer header above a table:"
    inner = head[0].add_table(rows=2, cols=2)
    inner._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    for row, texts in zip(inner.rows, (("Inner A", "Inner B"), ("a", "b")), strict=True):
        for cell, text in zip(row.cells, texts, strict=True):
            cell.text = text
    for cell in inner.rows[0].cells:
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="top"/>'))
    head[0].add_paragraph("Last line of the outer header")
    head[1].text = "Second header cell"
    table.rows[1].cells[0].text = "Body"

    pdf = _pdf_of(test_parameters, document)
    assert abs(_baseline(pdf, "Second header") - _baseline(pdf, "Last line")) < 1


def test_every_repeated_header_row_takes_the_first_row_alignment_of_the_style(test_parameters: TestParameters):
    """Word sets "Short" of both rows which repeat on the last line of the cell beside it; #372.

    The style aligns its first row at the bottom. Only the first row took it: "Short1" stood at the top.
    """
    document = Document()
    style = document.styles.add_style("Header at the bottom", WD_STYLE_TYPE.TABLE)
    style.element.append(parse_xml(f'<w:tblStylePr {nsdecls("w")} w:type="firstRow"><w:tcPr><w:vAlign w:val="bottom"/></w:tcPr></w:tblStylePr>'))
    table = document.add_table(rows=3, cols=2)
    table.style = style
    properties = table._tbl.tblPr
    for found in properties.findall(qn("w:tblLook")):
        properties.remove(found)
    properties.append(parse_xml(f'<w:tblLook {nsdecls("w")} w:firstRow="1"/>'))
    for index, row in enumerate(table.rows):
        row.cells[0].text = f"Lead{index}"
        row.cells[0].add_paragraph(f"End{index}")
        row.cells[1].text = f"Short{index}"
        if index < 2:
            row._tr.get_or_add_trPr().append(parse_xml(f"<w:tblHeader {nsdecls('w')}/>"))

    pdf = _pdf_of(test_parameters, document)
    for index in range(2):
        assert abs(_baseline(pdf, f"Short{index}") - _baseline(pdf, f"End{index}")) < 1, index
    assert abs(_baseline(pdf, "Short2") - _baseline(pdf, "Lead2")) < 1, "a row of the body stands at the top"


@pytest.mark.parametrize(
    ("look", "repeated", "repeat"),
    [
        ('w:firstRow="1"', 0, "on"),
        ('w:firstRow="0"', 2, "on"),
        ('w:firstRow="1"', 2, "on"),
        (None, 0, "on"),
        ('w:firstRow="0"', 0, "on"),
        (None, 1, "on"),
        (None, 2, "on"),
        ('w:firstRow="true"', 0, "on"),
        ('w:firstRow="on"', 0, "on"),
        ('w:val="0020"', 0, "on"),
        ('w:firstRow="0" w:val="0020"', 0, "on"),
        (None, 1, "0"),
        (None, 1, "false"),
    ],
    ids=["first-row", "repeated-rows", "both", "neither", "first-row-off", "one-repeated-row", "two-repeated-rows", "true", "on", "val", "first-row-off-over-val", "repeat-off", "repeat-false"],
)
def test_the_header_rows_of_the_preprocessor_are_those_of_pandoc(test_parameters: TestParameters, look: str | None, repeated: int, repeat: str):
    """The preprocessor carries an alignment for each row pandoc makes the header of, and no other; #372.

    pandoc's reader makes a header of the rows which repeat on each page, and of the first row of a table whose
    look turns its formatting on.
    """
    document = Document()
    table = document.add_table(rows=4, cols=2)
    properties = table._tbl.tblPr
    for found in properties.findall(qn("w:tblLook")):
        properties.remove(found)
    if look is not None:
        properties.append(parse_xml(f"<w:tblLook {nsdecls('w')} {look}/>"))
    for index in range(repeated):
        table.rows[index]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")} w:val="{repeat}"/>'))
    for index, row in enumerate(table.rows):
        for cell in row.cells:
            cell.text = f"row {index}"
    docx = io.BytesIO()
    document.save(docx)

    with zipfile.ZipFile(docx) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    tbl = root.find(f".//{w}tbl")
    tagged = docx_table_pre_process._head_cell_alignments(tbl, {})
    ours = sum(1 for row in tbl.findall(f"{w}tr") if any(tc in tagged for tc in row.findall(f"{w}tc")))

    encoded = base64.b64encode(docx.getvalue()).decode()
    exit_code, output = test_parameters.container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && echo {encoded} | base64 -d > /tmp/test/head.docx && {PANDOC_PATH} -f docx -t native /tmp/test/head.docx"])
    assert exit_code == 0, output.decode()
    head = output.decode().split("TableHead", 1)[1].split("TableBody", 1)[0]
    assert ours == len(re.findall(r"\bRow\b", head))


@pytest.mark.parametrize(("alignment", "word"), [(None, 0.0), ("center", 7.2), ("bottom", 14.6)], ids=["none", "center", "bottom"])
def test_a_cell_of_the_header_takes_its_vertical_alignment(test_parameters: TestParameters, alignment: str | None, word: float):
    """Word sets "Status" as far below "Requirement" as the alignment of its cell says: on its line without one; #372.

    pandoc sets every cell of the header at the bottom of its row: "Status" stood on the second line.
    """
    pdf = _pdf_of(test_parameters, _head_of_two_lines(alignment))
    below = _baseline(pdf, "Status") - _baseline(pdf, "Requirement")
    assert abs(below - word) < 1, (below, word)
    assert _baseline(pdf, "system") - _baseline(pdf, "Requirement") > 10, "the first cell takes two lines"


def _sized(text: str, payload: str = "") -> str:
    start = "{\\\\fontsize{8bp}{9.6bp}\\\\selectfont \\\\setbox\\\\strutbox\\\\hbox{\\\\vrule height 7.5133bp depth 2.2541bp width 0pt}"
    marker = f'Str "{OPEN}{payload}{CLOSE}", ' if payload else ""
    return f'{marker}RawInline (Format "latex") "{start}", Str "{text}", RawInline (Format "latex") "}}"'


def test_a_row_whose_cells_start_in_a_size_of_their_own_takes_the_line_of_that_size(test_parameters: TestParameters):
    """Word sets the first line of a cell as high as its text; LaTeX as high as the body text, unless told."""
    # Two paragraphs: the row stays one, it does not go out a line at a time
    two = f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [{_sized("b")}], Plain [{_sized("b2")}]]'
    native = _grid_table([_cell(_sized("a", GRID)), two], [_cell('Str "c"'), _cell(_sized("d"))])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\pdcRowLine{7.5133bp}" in latex, latex
    assert "\\pdcRowLine{}" in latex, "a row with a cell of the body text keeps the line of the table"
    assert "\\pdcRowLineEnd" in latex


def test_a_merged_cell_without_a_position_starts_at_its_top(test_parameters: TestParameters):
    """pandoc writes \\multirow without a position, which centres the text: Word sets it at the top of the cell."""
    native = _grid_table([_cell('Str "a"', rows=2), _cell('Str "b"')], [_cell('Str "c"')])
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/merged.native << 'EOF'\n{native}\nEOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/merged.native"])
    assert exit_code == 0, output.decode()
    assert "\\renewcommand{\\multirow}{\\@ifnextchar[\\pdc@multirow{\\pdc@multirow[t]}}" in output.decode()


def test_a_grid_on_the_right_reaches_past_the_margin_by_its_cell_margin(test_parameters: TestParameters):
    """Word sets a table on the right out past the right margin, as one on the left past the left one."""
    native = _grid_table([_cell(f'Str "{OPEN}{GRID};tw=0.4000;ta=right{CLOSE}a"'), _cell('Str "b"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "\\setlength{\\LTleft}{\\fill}\\setlength{\\LTright}{-\\tabcolsep}" in latex, latex


def test_an_equation_of_its_own_paragraph_in_a_cell_is_a_centered_line(test_parameters: TestParameters):
    """Word sets an equation of its own paragraph in a cell as in the body: centered, on a line of the font of the equations."""
    last = f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [Str "{OPEN}{GRID}{CLOSE}Formula:"], Plain [Math DisplayMath "x"]]'
    first = 'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [Math DisplayMath "y"], Plain [Str "after"]]'
    latex = " ".join(_native_latex(test_parameters.container, _grid_table([last, first])).split())
    # the cell ends the line of its last equation, with the space of the cell below it and no depth of a line of text
    assert "\\pdcDisplayLineEnd{\\(\\displaystyle x\\)}\\ifdefined\\pdcPlainCellSpace\\vadjust{\\kern\\pdcPlainCellSpace}\\fi &" in latex, latex
    assert "\\pdcCellFirstLine\\pdcCellDisplayLine{\\(\\displaystyle y\\)}" in latex, "an equation which starts its cell hangs from the top of the row"
    assert "\\pdcCellFirstLine\\pdcDisplayLineEnd" not in latex, "an equation below a paragraph starts a line of its own"


def _formula_table(display: bool) -> Document:
    document = Document()
    table = document.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "Step"
    cell = table.rows[0].cells[1]
    cell.paragraphs[0].add_run("Expected formula:")
    paragraph = cell.add_paragraph()
    if display:
        math = '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>'
        paragraph._p.append(parse_xml(math))
    else:
        paragraph.add_run("x")
    document.add_paragraph("After")
    return document


def test_a_cell_which_ends_with_an_equation_ends_its_row_there(test_parameters: TestParameters):
    """\\centering made the end of the row a break of the line: the row got an empty line more."""
    text, display = _pdf_of(test_parameters, _formula_table(False)), _pdf_of(test_parameters, _formula_table(True))
    assert abs(_top(display, "After") - _top(text, "After")) < 4, (_top(display, "After"), _top(text, "After"))


FRACTION = '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:f><m:num><m:r><m:t>5</m:t></m:r></m:num><m:den><m:r><m:t>7</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>'


def _fraction_rows() -> Document:
    """Rows with lines between them: two of text, one whose second cell is a fraction of its own paragraph, one of text."""
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(10)
    # The styles of pandoc's reference document, which give a cell without a style the space of Normal
    document.styles.add_style("Compact", WD_STYLE_TYPE.PARAGRAPH)
    table = document.add_table(rows=4, cols=2)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    for row, text in enumerate(("Upper", "Middle", "Label", "Lower")):
        table.rows[row].cells[0].text = text
    table.rows[2].cells[1].paragraphs[0]._p.append(parse_xml(FRACTION))
    return document


def test_an_equation_which_starts_a_cell_hangs_from_the_top_of_its_row(test_parameters: TestParameters):
    """Word sets the numerator of a fraction which starts a cell about where the first line of the next cell is.

    The cell had begun its first line, and the equation stood a line below it: its numerator below the baseline of that line.
    """
    pdf = _pdf_of(test_parameters, _fraction_rows())
    label, numerator, denominator = _baseline(pdf, "Label"), _baseline(pdf, "5"), _baseline(pdf, "7")
    assert numerator < label < denominator, (numerator, label, denominator)
    assert label - numerator < 8, "the equation is no higher than the top of the row"


def test_the_space_of_a_cell_goes_below_an_equation_which_ends_it(test_parameters: TestParameters):
    """Word sets the space of the paragraph below the fraction, as below a line of text: the fraction stood on the line of the cell."""
    pdf = _pdf_of(test_parameters, _fraction_rows())
    rows_of_text = _baseline(pdf, "Middle") - _baseline(pdf, "Upper")
    below_the_fraction = _baseline(pdf, "Lower") - _baseline(pdf, "7")
    assert abs(below_the_fraction - rows_of_text) < 5, (below_the_fraction, rows_of_text)


def _fraction_in_the_header() -> Document:
    """As _fraction_rows, with a header row whose second cell is a fraction of its own paragraph as well."""
    document = _fraction_rows()
    table = document.tables[0]
    table.rows[0].cells[0].paragraphs[0].runs[0].text = "Head"
    table.rows[1].cells[0].paragraphs[0].runs[0].text = "Upper"
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f"<w:tblHeader {nsdecls('w')}/>"))
    table.rows[0].cells[1].paragraphs[0]._p.append(parse_xml(FRACTION.replace(">5<", ">1<").replace(">7<", ">2<")))
    return document


def test_the_space_of_a_header_cell_goes_below_an_equation_which_ends_it(test_parameters: TestParameters):
    """A cell of the header stands on its last line, and gets its space as the depth of that line.

    An equation which ended one got none: the fraction stood on the line below the header.
    """
    pdf = _pdf_of(test_parameters, _fraction_in_the_header())
    below_the_header = _baseline(pdf, "Upper") - _baseline(pdf, "2")
    below_a_row = _baseline(pdf, "Lower") - _baseline(pdf, "7")
    assert abs(below_the_header - below_a_row) < 1, (below_the_header, below_a_row)


SUM = '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>+</m:t></m:r><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>'


def _compact_rows(equation: str) -> Document:
    """Rows of Compact, 1.8pt before and after: two of text, one whose second cell ends with the equation, one of text."""
    document = _fraction_rows()
    compact = document.styles["Compact"]
    compact.paragraph_format.space_before = Pt(1.8)
    compact.paragraph_format.space_after = Pt(1.8)
    table = document.tables[0]
    cell = table.rows[2].cells[1]
    cell.paragraphs[0]._p.remove(cell.paragraphs[0]._p.find(qn("m:oMathPara")))
    cell.paragraphs[0].add_run("Formula:")
    cell.add_paragraph()._p.append(parse_xml(equation))
    for row in table.rows:
        for each in row.cells:
            for paragraph in each.paragraphs:
                paragraph.style = compact
    return document


def test_an_equation_as_deep_as_a_line_ends_a_row_of_compact_as_a_line(test_parameters: TestParameters):
    """Word sets the row below such an equation as far below it as the rows of text: 0.02pt off.

    The last line of a cell has the depth of the strut of its row already, which holds the space of Compact below
    the line. The space went below that once more, and the rows below the equation stood 1.8pt lower.
    """
    pdf = _pdf_of(test_parameters, _compact_rows(SUM))
    rows_of_text = _baseline(pdf, "Middle") - _baseline(pdf, "Upper")
    below_the_equation = _baseline(pdf, "Lower") - _baseline(pdf, "+")
    assert abs(below_the_equation - rows_of_text) < 0.5, (below_the_equation, rows_of_text)


def test_an_equation_deeper_than_a_line_ends_a_row_of_compact_with_the_space_below_it(test_parameters: TestParameters):
    """Word sets the space of Compact below the denominator of a fraction: 2.1pt less than a row of text below its baseline."""
    pdf = _pdf_of(test_parameters, _compact_rows(FRACTION))
    rows_of_text = _baseline(pdf, "Middle") - _baseline(pdf, "Upper")
    below_the_fraction = _baseline(pdf, "Lower") - _baseline(pdf, "7")
    assert abs(below_the_fraction - (rows_of_text - 2.1)) < 1, (below_the_fraction, rows_of_text)


def _nested_table_html(size: str) -> str:
    """A table with a table nested in a cell below "Attributes:", the text of every cell in the size given; #379."""
    line, text = "border: 1px solid #000;", f"font-size: {size};"

    def cell(content: str) -> str:
        return f'<td style="{line}"><span style="{text}">{content}</span></td>'

    nested = f'<table style="border-collapse: collapse;"><tr>{cell("Status")}{cell("Approved")}</tr><tr>{cell("Owner")}{cell("Admin")}</tr></table>'
    attributes = f'<td style="{line}"><span style="{text}">Attributes:</span>{nested}</td>'
    rows = f"<tr>{cell('Work item')}{cell('Details')}</tr><tr>{cell('EL-20')}{attributes}</tr><tr>{cell('Next')}{cell('Row')}</tr>"
    return f'<html><body><p>A table with a nested table in a cell.</p><table style="border-collapse: collapse;">{rows}</table><p>After the table.</p></body></html>'


def _html_to_pdf(test_parameters: TestParameters, html: str) -> bytes:
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx?preserve_table_styles=true", data=html.encode())
    assert docx.status_code == 200, docx.text
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    pdf = session.post(f"{base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.content, mime)})
    assert pdf.status_code == 200, pdf.text
    return pdf.content


def _styled_cell_pdf(test_parameters: TestParameters, content: str) -> bytes:
    """A table with the styles of its cells, as docx-exporter sends it: the paragraphs of its cells are in Normal."""
    line = "border: 1px solid #CCCCCC; padding: 5px;"
    rows = f'<tr><td style="{line}">Step</td><td style="{line}">{content}</td></tr><tr><td style="{line}">Next</td><td style="{line}">Row</td></tr>'
    html = f'<html><body><p>Before the table.</p><table style="width: 100%; border-collapse: collapse;">{rows}</table></body></html>'
    return _html_to_pdf(test_parameters, html)


def _html_to_docx_with_normal_before(test_parameters: TestParameters, html: str, before_twips: int) -> bytes:
    """The DOCX of an HTML with preserved table styles, its paragraphs of Normal given a space before."""
    docx = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/html/to/docx?preserve_table_styles=true", data=html.encode())
    assert docx.status_code == 200, docx.text
    source, output = zipfile.ZipFile(io.BytesIO(docx.content)), io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "word/styles.xml":
                data = data.replace(b'<w:spacing w:after="200"/>', f'<w:spacing w:before="{before_twips}" w:after="200"/>'.encode(), 1)
            target.writestr(item, data)
    return output.getvalue()


@pytest.mark.parametrize(("size", "word"), [("9pt", 21.6), ("12pt", 25.4)])
def test_a_nested_table_stands_below_a_paragraph_in_a_size_of_its_own_as_in_word(test_parameters: TestParameters, size: str, word: float):
    """Word sets the first row of the nested table 21.6pt below "Attributes:" in 9pt, 25.4pt below it in the size of the body.

    A paragraph in a size of its own ends with the line of that size. The space of the cell went after that end,
    where it started an empty line: the nested table stood 16.6pt lower than in Word.
    """
    pdf = _html_to_pdf(test_parameters, _nested_table_html(size))
    below = _baseline(pdf, "Status") - _baseline(pdf, "Attributes")
    assert abs(below - word) < 1, (below, word)


def test_a_row_of_a_cell_of_two_paragraphs_keeps_the_space_between_them_out_of_the_split_of_a_row(test_parameters: TestParameters):
    """The space between paragraphs goes only into a cell of two paragraphs or more; #380.

    Such a row never goes out a line at a time (\\pdcWrapCell, split by \\vsplit): that takes a cell of one
    paragraph in each column. A row of one paragraph per cell gets no space between paragraphs.
    """
    two = f'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [Str "{OPEN}{GRID}{CLOSE}first"], Plain [Str "second"]]'
    latex = " ".join(_native_latex(test_parameters.container, _grid_table([two, _cell('Str "b"')], [_cell('Str "c"'), _cell('Str "d"')])).split())
    assert "first\\ifdefined\\pdcPlainParagraphSpace\\vadjust{\\kern\\pdcPlainParagraphSpace}\\fi second" in latex, latex
    rows = latex.split("\\hline")
    assert "\\pdcWrapCell" not in next(row for row in rows if "first" in row), "the row of two paragraphs goes out whole"
    assert "\\pdcWrapCell" in next(row for row in rows if "{c" in row or "c\\ifdefined" in row), "the row of one paragraph per cell goes out a line at a time"
    assert latex.count("\\vadjust{\\kern\\pdcPlainParagraphSpace}") == 1


def test_two_paragraphs_of_a_cell_in_normal_stand_apart_by_its_space_after(test_parameters: TestParameters):
    """Word sets the second paragraph 24.7pt below the first, and the next row 25.4pt below the second; #380.

    The space after of Normal went only after the last paragraph of a cell: the second stood 10pt higher.
    """
    pdf = _styled_cell_pdf(test_parameters, "<p>First paragraph:</p><p>Second paragraph</p>")
    between, below = _baseline(pdf, "Second") - _baseline(pdf, "First"), _baseline(pdf, "Next") - _baseline(pdf, "Second")
    assert abs(between - 24.7) < 1, between
    assert abs(below - 25.4) < 1, below


def test_an_equation_after_a_paragraph_of_a_cell_in_normal_stands_below_its_space_after(test_parameters: TestParameters):
    """Word sets the equation 25.0pt below "Expected formula:", and the next row 24.5pt below the equation; #380."""
    pdf = _styled_cell_pdf(test_parameters, 'Expected formula:<script type="math/tex; mode=display">x + x = y</script>')
    between, below = _baseline(pdf, "+") - _baseline(pdf, "Expected"), _baseline(pdf, "Next") - _baseline(pdf, "+")
    assert abs(between - 25.0) < 1, between
    assert abs(below - 24.5) < 1, below


def test_a_paragraph_after_an_equation_which_starts_its_cell_stands_below_its_space_after(test_parameters: TestParameters):
    """Word sets "After the formula" 23.8pt below an equation which starts its cell; #380.

    The equation stands behind the commands which start the cell: the space after it took the form for a line
    of text, which LaTeX refused after the line of the equation, and the conversion failed.
    """
    pdf = _styled_cell_pdf(test_parameters, '<script type="math/tex; mode=display">x + x = y</script>After the formula')
    below = _baseline(pdf, "After the formula") - _baseline(pdf, "+")
    assert abs(below - 23.8) < 1, below


def test_two_paragraphs_of_a_cell_take_the_larger_space_of_normal_between_them(test_parameters: TestParameters):
    """Normal 6pt before and 10pt after: Word sets the second paragraph 24.7pt below the first, the next row 31.4pt below; #380.

    The sum of both spaces went between them, as at the end of a cell: 30.6pt.
    """
    line = "border: 1px solid #CCCCCC; padding: 5px;"
    rows = f'<tr><td style="{line}">Step</td><td style="{line}"><p>First paragraph:</p><p>Second paragraph</p></td></tr><tr><td style="{line}">Next</td><td style="{line}">Row</td></tr>'
    html = f'<html><body><p>Before the table.</p><table style="width: 100%; border-collapse: collapse;">{rows}</table></body></html>'
    pdf = _pdf_of_docx(test_parameters, _html_to_docx_with_normal_before(test_parameters, html, 120))
    between, below = _baseline(pdf, "Second") - _baseline(pdf, "First"), _baseline(pdf, "Next") - _baseline(pdf, "Second")
    assert abs(between - 24.7) < 1, between
    assert abs(below - 31.4) < 1, below


def test_a_paragraph_after_an_equation_of_a_cell_in_normal_stands_below_its_space_after(test_parameters: TestParameters):
    """Word sets "After the formula" 23.8pt below the equation, which ends its own paragraph; #380."""
    pdf = _styled_cell_pdf(test_parameters, 'Expected formula:<script type="math/tex; mode=display">x + x = y</script>After the formula')
    below = _baseline(pdf, "After the formula") - _baseline(pdf, "+")
    assert abs(below - 23.8) < 1, below


def test_a_table_of_an_absolute_width_takes_the_widths_and_margins_of_word(test_parameters: TestParameters):
    """A table of twips sets each column as wide as its twips, with the cell margins of Word."""
    native = _grid_table([_cell(f'Str "{OPEN}tw=1.0000;aw=1{CLOSE}a"'), _cell('Str "b"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "p{\\dimexpr0.5000\\linewidth-2\\tabcolsep\\relax}" in latex, latex
    assert "\\setlength{\\tabcolsep}{3pt}" not in latex


def _baseline(pdf: bytes, text: str) -> float:
    """The height of the baseline of the first character of the text given, in points from the top of the page."""
    page = pdfium.PdfDocument(pdf)[0]
    textpage = page.get_textpage()
    found = textpage.search(text, match_case=True).get_next()
    assert found, text
    x, y = ctypes.c_double(), ctypes.c_double()
    pdfium_c.FPDFText_GetCharOrigin(textpage.raw, found[0], ctypes.byref(x), ctypes.byref(y))
    return page.get_height() - y.value


def test_the_space_of_a_cell_goes_below_a_picture_lowered_below_the_descent(test_parameters: TestParameters):
    """A picture of 12 points lowered by 3 makes the line deeper than the descent of Liberation Sans, 2.54 points."""
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(10)
    # The styles of pandoc's reference document, which give a cell without a style the space of Normal
    document.styles.add_style("Compact", WD_STYLE_TYPE.PARAGRAPH)
    table = document.add_table(rows=3, cols=1)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    table.rows[0].cells[0].text = "Upper"
    paragraph = table.rows[1].cells[0].paragraphs[0]
    paragraph.add_run("Icon ")
    png = io.BytesIO()
    Image.new("RGB", (16, 16), (40, 90, 170)).save(png, format="PNG")
    picture = paragraph.add_run()
    picture.add_picture(io.BytesIO(png.getvalue()), width=Pt(12), height=Pt(12))
    picture._r.get_or_add_rPr().append(parse_xml(f'<w:position {nsdecls("w")} w:val="-6"/>'))
    table.rows[2].cells[0].text = "Lower"
    pdf = _pdf_of(test_parameters, document)
    plain = _baseline(pdf, "Icon") - _baseline(pdf, "Upper")
    lowered = _baseline(pdf, "Lower") - _baseline(pdf, "Icon")
    assert abs(lowered - plain - (3 - 0.2119 * 12)) < 0.1, (plain, lowered)


def test_a_header_row_which_fits_at_the_foot_of_a_page_stays_there(test_parameters: TestParameters):
    """Word sets the header row where it fits, at the foot of the page, and repeats it above the row on the next page.

    longtable moved the whole table to the next page unless its first row fitted as well.
    """
    document = Document()
    for number in range(23):
        document.add_paragraph(f"Filler {number}")
    # A last line which leaves room for the header row, not for the row below it
    document.add_paragraph().add_run("Last").font.size = Pt(15.75)
    table = document.add_table(rows=2, cols=2)
    for index, text in enumerate(("Head A", "Head B")):
        table.rows[0].cells[index].text = text
    for index, text in enumerate(("Body A", "Body B")):
        table.rows[1].cells[index].text = text
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f"<w:tblHeader {nsdecls('w')}/>"))
    document.add_paragraph("After")
    pdf = pdfium.PdfDocument(_pdf_of(test_parameters, document))
    pages = [pdf[index].get_textpage().get_text_range().split() for index in range(len(pdf))]
    assert pages[0][-4:] == ["Head", "A", "Head", "B"], pages
    assert pages[1][:8] == ["Head", "A", "Head", "B", "Body", "A", "Body", "B"], pages


def test_a_column_narrower_than_its_margins_keeps_the_width_of_word(test_parameters: TestParameters):
    """Word sets a table of 50px, three columns of 12.5pt, as wide as that: the text of each starts its margin in.

    The margins of the columns made them wider, 18.1pt with the tight margins of pandoc.
    """
    document = Document()
    table = document.add_table(rows=1, cols=3)
    grid = table._tbl.tblGrid
    for column in grid.findall(qn("w:gridCol")):
        column.set(qn("w:w"), "250")
    width = table._tbl.tblPr.find(qn("w:tblW"))
    width.set(qn("w:w"), "750")
    width.set(qn("w:type"), "dxa")
    table._tbl.tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    for index, cell in enumerate(table.rows[0].cells):
        cell.text = str(index + 1)
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="250" w:type="dxa"/>'))
    textpage = pdfium.PdfDocument(_pdf_of(test_parameters, document))[0].get_textpage()
    text = textpage.get_text_range()
    origins = []
    for digit in "123":
        x, y = ctypes.c_double(), ctypes.c_double()
        pdfium_c.FPDFText_GetCharOrigin(textpage.raw, text.index(digit), ctypes.byref(x), ctypes.byref(y))
        origins.append(x.value)
    assert [round(right - left, 1) for left, right in itertools.pairwise(origins)] == [12.5, 12.5]


def test_the_hyphen_of_a_shared_glyph_in_a_table_copies_out_as_a_hyphen(test_parameters: TestParameters):
    """Noto Sans CJK maps - and U+2011 to one glyph: a table written as LaTeX gives its cells their own text as well."""
    document = Document()
    document.styles["Normal"].font.name = "Noto Sans CJK SC"
    table = document.add_table(rows=1, cols=1)
    table.rows[0].cells[0].text = "e-mail"
    width = table._tbl.tblPr.find(qn("w:tblW"))
    width.set(qn("w:w"), "2000")
    width.set(qn("w:type"), "dxa")
    assert "e-mail" in pdfium.PdfDocument(_pdf_of(test_parameters, document))[0].get_textpage().get_text_range()


def test_an_indented_paragraph_of_a_cell_starts_its_indent_in(test_parameters: TestParameters):
    """Word sets the text of a cell paragraph with a left indent of 142 twips 7.1pt further right.

    pandoc read the paragraph as a BlockQuote, which the filter unwraps, and the indent was lost.
    """
    document = Document()
    table = document.add_table(rows=2, cols=1)
    table.rows[0].cells[0].text = "Plain"
    table.rows[1].cells[0].text = "Indented"
    table.rows[1].cells[0].paragraphs[0]._p.get_or_add_pPr().append(parse_xml(f'<w:ind {nsdecls("w")} w:left="142"/>'))
    textpage = pdfium.PdfDocument(_pdf_of(test_parameters, document))[0].get_textpage()
    text = textpage.get_text_range()
    origins = []
    for word in ("Plain", "Indented"):
        x, y = ctypes.c_double(), ctypes.c_double()
        pdfium_c.FPDFText_GetCharOrigin(textpage.raw, text.index(word), ctypes.byref(x), ctypes.byref(y))
        origins.append(x.value)
    assert round(origins[1] - origins[0], 1) == 7.1


def test_a_centered_paragraph_of_a_cell_centers_between_its_indent_and_the_right_margin(test_parameters: TestParameters):
    """Word centers a cell paragraph with a left indent of 1440 twips 36pt right of one without (35.7pt measured).

    The centered cell dropped the indent and stayed in the middle of the cell.
    """
    document = Document()
    table = document.add_table(rows=2, cols=1)
    for row, (text, indent) in zip(table.rows, (("Plain", None), ("Indented", 1440)), strict=True):
        paragraph = row.cells[0].paragraphs[0]
        paragraph.add_run(text)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if indent:
            paragraph._p.get_or_add_pPr().append(parse_xml(f'<w:ind {nsdecls("w")} w:left="{indent}"/>'))
    textpage = pdfium.PdfDocument(_pdf_of(test_parameters, document))[0].get_textpage()
    text = textpage.get_text_range()
    middles = []
    for word in ("Plain", "Indented"):
        start = text.index(word)
        left = textpage.get_charbox(start)[0]
        right = textpage.get_charbox(start + len(word) - 1)[2]
        middles.append((left + right) / 2)
    assert abs(middles[1] - middles[0] - 36) < 0.5, middles


@pytest.mark.parametrize(("alignment", "expected"), [("AlignDefault", "\\pdcCellRaggedright{}\\advance\\leftskip7.1bp\\relax Indented"), ("AlignCenter", "\\pdcCellCentering{}\\advance\\leftskip7.1bp\\relax Indented")])
def test_the_indent_of_a_cell_adds_to_its_left_skip(test_parameters: TestParameters, alignment: str, expected: str):
    """The alignment switch sets \\leftskip, so the indent comes after it."""
    cell = f'Cell ("",[],[]) {alignment} (RowSpan 1) (ColSpan 1) [Plain [Str "{OPEN}li=7.1{CLOSE}Indented"]]'
    latex = _native_latex(test_parameters.container, _grid_table([cell, _cell('Str "Other"')]))
    assert expected in latex, latex


def test_a_cell_of_small_text_hangs_from_the_top_of_a_row_with_large_text(test_parameters: TestParameters):
    """Word sets the first line of each cell as high as its own text, from the top of the row.

    The row took the line of its largest text, and the small text of the other cells sat on its baseline.
    """
    document = Document()
    # A font the image knows the lines of: Word's line of 11pt is 10.32pt high, of 14pt 13.13pt
    document.styles["Normal"].font.name = "Liberation Sans"
    table = document.add_table(rows=2, cols=2)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    for row, texts in zip(table.rows, ((("Above", 11), ("Other", 11)), (("Small", 11), ("Large", 14))), strict=True):
        for cell, (text, size) in zip(row.cells, texts, strict=True):
            run = cell.paragraphs[0].add_run(text)
            # A run of a size and a color of its own, as Polarion writes it
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(0x67, 0x4E, 0xA7)
    pdf = _pdf_of(test_parameters, document)
    textpage = pdfium.PdfDocument(pdf)[0].get_textpage()
    text = textpage.get_text_range()
    baselines = []
    for word in ("Small", "Large"):
        x, y = ctypes.c_double(), ctypes.c_double()
        pdfium_c.FPDFText_GetCharOrigin(textpage.raw, text.index(word), ctypes.byref(x), ctypes.byref(y))
        baselines.append(y.value)
    assert abs(baselines[0] - baselines[1] - 2.81) < 0.1, baselines


def test_a_line_between_rows_under_a_merged_cell_adds_its_width_as_well(test_parameters: TestParameters):
    """Word sets the row below a line its width lower, also where a cell merged down crosses the line.

    The line under part of a row is a \\cline, which takes no room of its own: the row below it sat a line higher, 3.55pt
    here. Word for Mac sets the two rows 0.16pt apart in pitch, this the rows of a merged cell and of a plain row 0.56pt.
    """
    document = Document()
    # The cells of Polarion's tables keep a space below their paragraph, which makes them deeper than a row
    document.styles["Normal"].paragraph_format.space_after = Pt(10)
    table = document.add_table(rows=4, cols=2)
    sides = "".join(f'<w:{side} w:val="single" w:sz="24" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    # Below a first row, which may be a header of its own
    for (row, column), text in {(0, 0): "Head", (0, 1): "Top", (1, 0): "Merged", (1, 1): "First", (2, 1): "Second", (3, 0): "Low", (3, 1): "Third"}.items():
        table.rows[row].cells[column].text = text
    table.rows[1].cells[0]._tc.get_or_add_tcPr().append(parse_xml(f'<w:vMerge {nsdecls("w")} w:val="restart"/>'))
    table.rows[2].cells[0]._tc.get_or_add_tcPr().append(parse_xml(f"<w:vMerge {nsdecls('w')}/>"))
    pdf = _pdf_of(test_parameters, document)
    first, second, third = (_top(pdf, text) for text in ("First", "Second", "Third"))
    assert abs((second - first) - (third - second)) < 1, (first, second, third)


def test_the_lines_of_cells_which_take_their_room_get_no_more(test_parameters: TestParameters):
    """A table whose cells draw lines of their own sets each line with its width already: no cell reaches deeper."""
    document = Document()
    table = document.add_table(rows=3, cols=2)
    for (row, column), text in {(0, 0): "Head", (0, 1): "Top", (1, 0): "Merged", (1, 1): "First", (2, 1): "Second"}.items():
        table.rows[row].cells[column].text = text
    # Lines of two colors: one grid for the table could not draw them
    for row_index, row in enumerate(table.rows):
        for cell in row.cells:
            color = "FF0000" if row_index % 2 else "0000FF"
            borders = "".join(f'<w:{side} w:val="single" w:sz="12" w:space="0" w:color="{color}"/>' for side in ("top", "left", "bottom", "right"))
            cell._tc.get_or_add_tcPr().append(parse_xml(f"<w:tcBorders {nsdecls('w')}>{borders}</w:tcBorders>"))
    table.rows[1].cells[0]._tc.get_or_add_tcPr().append(parse_xml(f'<w:vMerge {nsdecls("w")} w:val="restart"/>'))
    table.rows[2].cells[0]._tc.get_or_add_tcPr().append(parse_xml(f"<w:vMerge {nsdecls('w')}/>"))
    docx = io.BytesIO()
    document.save(docx)
    latex = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/latex", files={"source": ("in.docx", docx.getvalue())}).text
    assert "\\pdcGridV" in latex, "a table which draws the lines of its cells"
    assert "\\vadjust{\\kern\\arrayrulewidth}" not in latex


def test_a_space_after_a_line_break_in_a_cell_of_a_grid_stays(test_parameters: TestParameters):
    """A grid is written as LaTeX before the styles see its cells: the space after a line break stays there as well."""
    document = Document()
    table = document.add_table(rows=1, cols=2)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    paragraph = table.rows[0].cells[0].paragraphs[0]
    paragraph.add_run("Title").add_break()
    # As Polarion writes the attribute line of a work item: zero width spaces, a space, then the attributes
    for text in ("\u200b\u200b", " ", "[Draft]"):
        paragraph.add_run(text)
    table.rows[0].cells[1].text = "Other"
    docx = io.BytesIO()
    document.save(docx)
    latex = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/latex", files={"source": ("in.docx", docx.getvalue())}).text
    assert "\\pdcGridV" in latex, "a table written as LaTeX"
    assert "\\\\ \\hbox{}\\hspace{0pt}\\hspace{0pt} {[}Draft" in " ".join(latex.split()), latex


def test_a_cell_paragraph_of_a_document_without_compact_keeps_the_space_of_normal(test_parameters: TestParameters):
    """Word sets a cell paragraph in Normal, with its space below: 10pt from the defaults of python-docx.

    Word for Mac sets the rows of this table 25.5pt apart. Without the Compact style of pandoc the space of Normal
    went missing in the cells, and the rows stood 16pt apart.
    """
    document = Document()
    table = document.add_table(rows=3, cols=2)
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table._tbl.tblPr.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    for row, text in enumerate(("Top", "First", "Second")):
        table.rows[row].cells[0].text = text
        table.rows[row].cells[1].text = "x"
    pdf = _pdf_of(test_parameters, document)
    first, second = _top(pdf, "First"), _top(pdf, "Second")
    assert 25 < second - first < 26.2, (first, second)


def _rows_of_normal(after: int) -> Document:
    """A table of three rows, in a document whose styles give no size, and the space below Normal given in twips."""
    document = Document()
    defaults = document.styles.element.find(qn("w:docDefaults"))
    for size in defaults.iter(qn("w:sz")):
        size.getparent().remove(size)
    for spacing in defaults.iter(qn("w:spacing")):
        spacing.set(qn("w:after"), str(after))
    table = document.add_table(rows=3, cols=1)
    for row, text in enumerate(("Top", "First", "Second")):
        table.rows[row].cells[0].text = text
    return document


def test_a_cell_keeps_the_space_of_normal_where_the_styles_give_no_size(test_parameters: TestParameters):
    """The space of Normal below a cell paragraph needs no size of the body: only the rows of Compact do."""
    spaced = _pdf_of(test_parameters, _rows_of_normal(200))
    tight = _pdf_of(test_parameters, _rows_of_normal(0))
    pitch = (_top(spaced, "Second") - _top(spaced, "First")) - (_top(tight, "Second") - _top(tight, "First"))
    assert abs(pitch - 10) < 0.5, pitch


URL = "https://example.com/polarion/project/elibrary/wiki/Specification/ProductSpecification"


def _narrow_first_column() -> Document:
    """A table as wide as the line, its first column narrower than the word in it: #371."""
    document = Document()
    table = document.add_table(rows=2, cols=2)
    properties = table._tbl.tblPr
    properties.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>'))
    properties.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>'))
    for column, width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), (916, 8110), strict=True):
        column.set(qn("w:w"), str(width))
    for row, texts in zip(table.rows, (("Requirements", "Link"), ("Spec", URL)), strict=True):
        for cell, text, width in zip(row.cells, texts, (916, 8110), strict=True):
            cell.text = text
            cell.width = Twips(width)
    return document


def _boxes(pdf: bytes, text: str) -> list[tuple[float, float, float, float]]:
    """The boxes of the characters of TEXT, which may run over lines: left, bottom, right, top."""
    page = pdfium.PdfDocument(pdf)[0]
    textpage = page.get_textpage()
    chars = [(textpage.get_text_range(index, 1), textpage.get_charbox(index)) for index in range(textpage.count_chars())]
    printed = [(char, box) for char, box in chars if not char.isspace()]
    joined = "".join(char for char, _ in printed)
    start = joined.find(text)
    assert start >= 0, joined
    return [box for _, box in printed[start : start + len(text)]]


def test_a_word_wider_than_its_cell_breaks_at_the_end_of_the_line(test_parameters: TestParameters):
    """Word keeps the widths of the grid and breaks such a word at a character, as far out as the line goes."""
    pdf = _pdf_of(test_parameters, _narrow_first_column())

    word, link = _boxes(pdf, "Requirements"), _boxes(pdf, "Link")
    assert max(box[2] for box in word) < link[0][0], "the word stays in its column"
    first = [box for box in word if box[3] > word[0][1]]
    assert len(first) < len(word), "the word goes on over the next line"
    following = word[len(first)]
    # The margins of the cells between the text of the two columns: 5.4pt each
    assert first[-1][2] + following[2] - following[0] > link[0][0] - 10.8, "the first line holds as many characters as fit"
    page_width = pdfium.PdfDocument(pdf)[0].get_width()
    url = _boxes(pdf, URL)
    assert max(box[2] for box in url) < page_width - 71, "the URL stays inside the right margin"
    assert url[-1][3] < url[0][1], "the URL goes on over a second line"


def test_a_word_which_fits_its_cell_keeps_breaking_at_spaces_only(test_parameters: TestParameters):
    document = _narrow_first_column()
    document.tables[0].rows[1].cells[1].text = "Specification of the product"
    pdf = _pdf_of(test_parameters, document)

    words = _boxes(pdf, "Specificationoftheproduct")
    assert max(box[3] for box in words) - min(box[1] for box in words) < 15, "the text fits one line and stays on it"


def test_a_short_word_of_wide_letters_wider_than_its_cell_breaks_as_well(test_parameters: TestParameters):
    """Six letters W are wider than a column of 916 twips: the count of the letters says nothing of their width."""
    document = _narrow_first_column()
    document.tables[0].rows[0].cells[0].text = "WWWWWW"
    pdf = _pdf_of(test_parameters, document)

    word, link = _boxes(pdf, "WWWWWW"), _boxes(pdf, "Link")
    assert max(box[2] for box in word) < link[0][0], "the word stays in its column"


def test_a_word_in_an_indented_cell_breaks_within_the_indent(test_parameters: TestParameters):
    """The indent of the paragraph takes its part of the line of the cell: each line of the word ends before the next column."""
    document = _narrow_first_column()
    document.tables[0].rows[0].cells[0].paragraphs[0].paragraph_format.left_indent = Pt(12)
    pdf = _pdf_of(test_parameters, document)

    word, link = _boxes(pdf, "Requirements"), _boxes(pdf, "Link")
    # The margins of the cells between the text of the two columns: 5.4pt each
    assert max(box[2] for box in word) < link[0][0] - 10.8, "each line ends at the line of the cell"
    drops = [above[1] - below[1] for above, below in itertools.pairwise(word) if above[1] - below[1] > 8]
    assert drops and all(drop < 20 for drop in drops), "the word goes on line by line, each right below the one before"


def _span(pdf: bytes, text: str) -> float:
    boxes = _boxes(pdf, text)
    return boxes[-1][2] - boxes[0][0]


def test_a_word_which_fits_its_cell_keeps_its_kerning(test_parameters: TestParameters):
    """A narrow column makes the filter measure the words of the table: one which fits is set as it was, kerned."""
    narrow = _narrow_first_column()
    narrow.tables[0].rows[1].cells[1].text = "WAVEAVATAR"
    even = _narrow_first_column()
    for column in even.tables[0]._tbl.tblGrid.findall(qn("w:gridCol")):
        column.set(qn("w:w"), "4513")
    for row in even.tables[0].rows:
        for cell in row.cells:
            cell.width = Twips(4513)
    even.tables[0].rows[1].cells[1].text = "WAVEAVATAR"

    measured, plain = _pdf_of(test_parameters, narrow), _pdf_of(test_parameters, even)
    assert abs(_span(measured, "WAVEAVATAR") - _span(plain, "WAVEAVATAR")) < 0.05


def test_a_hyphen_of_a_long_word_stays_a_place_to_break_the_line(test_parameters: TestParameters):
    """Word breaks a line after the hyphen of "drop-down": the hyphen stays out of the box which measures the word."""
    word = "short-averyveryverylongpiece-end"
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell(f'Str "{word}"')])
    latex = " ".join(_native_latex(test_parameters.container, native).split())
    assert "short-\\pdcWord{{a}{v}{e}{r}{y}" in latex, latex
    assert "{g}{p}{i}{e}{c}{e}}-end" in latex, latex
