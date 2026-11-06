"""Integration tests for the lines of a table in ``filters/docx/docx_tables_to_latex.lua``.

app/docx_table_pre_process.py tags the first cell of a styled table with
``rl=none`` or ``rl=head``; the filter sets the table without the booktabs
rules Word does not draw. The tests run pandoc inside the container on a
markdown table whose first cell carries the sentinel.
"""

from __future__ import annotations

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


def test_a_narrow_absolute_table_restores_the_padding_of_the_document_after_it(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "tw=0.3000;aw=1")
    assert "\\begingroup\\setlength{\\tabcolsep}{3pt}" in latex
    assert "\\setlength{\\tabcolsep}{6pt}" not in latex
    assert latex.count("\\begingroup") == latex.count("\\endgroup")


def test_a_table_of_compact_paragraphs_takes_the_rows_of_compact(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "cc=1")
    assert latex.startswith("\\begingroup\\ifdefined\\pdcCompactCellStretch\\renewcommand{\\arraystretch}{\\pdcCompactCellStretch}\\fi")
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
    assert "\\begin{longtable}[]{|" in latex
    assert latex.count("|") >= 3
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
    assert "\\begin{longtable}[]{|" in latex
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
    assert latex.count("\\rule[-\\dimexpr\\dp\\strutbox+\\pdcPlainCellSpace\\relax]{0pt}{0pt}") == 4


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
    assert "\\ifdefined\\pdcCompactCellBottom\\vspace{\\pdcCompactCellBottom}\\fi" in latex


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


def test_a_table_in_a_cell_stays_a_table_for_the_filter_of_nested_tables(test_parameters: TestParameters):
    inner = _grid_table([_cell(f'Str "{OPEN}tw=0.5000{CLOSE}x"'), _cell('Str "y"')])[1:-1]
    outer = _grid_table([_cell(f'Str "{OPEN}tw=1.0000{CLOSE}a"'), _cell('Str "b"')])
    native = outer.replace('[Plain [Str "b"]]', f"[{inner}]")
    latex = _native_latex(test_parameters.container, native)
    assert "\\real{" in latex, "pandoc writes the outer table, with the inner one in it"
