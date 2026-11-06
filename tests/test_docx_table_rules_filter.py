"""Integration tests for the lines of a table in ``filters/docx_tables_to_latex.lua``.

app/docx_table_pre_process.py tags the first cell of a styled table with
``rl=none`` or ``rl=head``; the filter sets the table without the booktabs
rules Word does not draw. The tests run pandoc inside the container on a
markdown table whose first cell carries the sentinel.
"""

from __future__ import annotations

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_tables_to_latex.lua"
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


def test_a_grid_with_a_picture_keeps_the_table_of_pandoc(test_parameters: TestParameters):
    native = _grid_table([_cell(f'Str "{OPEN}{GRID}{CLOSE}a"'), _cell('Image ("",[],[]) [] ("x.png","")')])
    latex = _native_latex(test_parameters.container, native)
    assert "{|" not in latex
    assert "\\toprule" in latex


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
