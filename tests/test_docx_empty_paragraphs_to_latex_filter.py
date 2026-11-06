"""Integration tests for ``filters/docx_empty_paragraphs_to_latex.lua``.

Word prints an empty paragraph as a blank line. The filter turns each empty
paragraph of the document body into a ``\\strut`` paragraph, so the PDF keeps
the line. The unit cases feed pandoc a native AST inside the container; the
end-to-end cases send a DOCX through the service.
"""

from __future__ import annotations

from io import BytesIO

from docx import Document

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_empty_paragraphs_to_latex.lua"


def _native_to_latex(container, native: str) -> str:
    """Convert a native AST to LaTeX through the filter, whitespace collapsed."""
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/in.native"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return " ".join(output.decode("utf-8").split())


def test_empty_paragraph_becomes_strut(test_parameters: TestParameters):
    flat = _native_to_latex(test_parameters.container, '[ Para [ Str "before" ], Para [], Para [ Str "after" ] ]')
    assert flat == "before \\strut after", flat


def test_whitespace_and_empty_anchor_count_as_empty(test_parameters: TestParameters):
    native = '[ Para [ Space ], Para [ Span ( "_Toc1" , [ "anchor" ] , [] ) [] ] ]'
    assert _native_to_latex(test_parameters.container, native).count("\\strut") == 2


def test_each_manual_line_break_adds_a_line(test_parameters: TestParameters):
    """Word prints a paragraph with two manual breaks as three lines."""
    flat = _native_to_latex(test_parameters.container, "[ Para [ LineBreak, Space, LineBreak ] ]")
    assert flat.count("\\strut") == 3, flat


def test_bookmark_stays_on_the_line_it_marks(test_parameters: TestParameters):
    """A bookmark after a manual break marks the second line, not the first."""
    native = '[ Para [ LineBreak, Span ( "_Toc2" , [ "anchor" ] , [] ) [] ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.index("\\\\") < flat.index("\\label{_Toc2}"), flat


def test_bookmark_of_an_empty_paragraph_is_kept(test_parameters: TestParameters):
    """A link to the bookmark still has its target."""
    flat = _native_to_latex(test_parameters.container, '[ Para [ Span ( "_Toc1" , [ "anchor" ] , [] ) [] ] ]')
    assert "\\label{_Toc1}" in flat, flat
    assert "\\strut" in flat, flat


def test_empty_paragraph_inside_div_becomes_strut(test_parameters: TestParameters):
    native = '[ Div ( "" , [] , [ ( "custom-style" , "Body Text" ) ] ) [ Para [] ] ]'
    assert "\\strut" in _native_to_latex(test_parameters.container, native)


def _marked(*inlines: str) -> str:
    return f'[ Div ( "" , [] , [ ( "custom-style" , "PandocEmptyLines" ) ] ) [ Para [ {", ".join(inlines)} ] ] ]'


def test_a_marked_paragraph_gets_its_lines(test_parameters: TestParameters):
    """The preprocessor marks a paragraph of one line break with two lines, as Word prints it."""
    flat = _native_to_latex(test_parameters.container, _marked('Span ( "_Toc3" , [ "anchor" ] , [] ) []', 'Str "2"'))
    assert flat.count("\\strut") == 2, flat
    assert "\\label{_Toc3}" in flat, flat
    assert "2" not in flat.replace("\\label{_Toc3}", ""), flat


SHOW_STYLE = 'function Div(div) if div.attributes["pdc-style"] then return {pandoc.RawBlock("latex", "STYLE=" .. div.attributes["pdc-style"]), table.unpack(div.content)} end end'


def test_a_marked_paragraph_keeps_its_style(test_parameters: TestParameters):
    """docx_document_styles_to_latex puts the space of that style around the lines."""
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/show.lua << 'HEREDOC_EOF'\n{SHOW_STYLE}\nHEREDOC_EOF"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{_marked('Str "2"', 'Space', 'Str "BodyText"')}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} --lua-filter=/tmp/test/show.lua /tmp/test/in.native"])
    assert exit_code == 0, output.decode()
    assert "STYLE=BodyText" in output.decode()
    assert output.decode().count("\\strut") == 2


def test_a_marked_paragraph_without_a_number_is_left_as_it_is(test_parameters: TestParameters):
    assert "\\strut" not in _native_to_latex(test_parameters.container, _marked('Str "x"'))


def test_a_kept_line_break_marker_becomes_an_empty_line(test_parameters: TestParameters):
    flat = _native_to_latex(test_parameters.container, '[ Para [ Str "PDCKEEPBREAK" , LineBreak , Str "text" ] ]')
    assert flat == "\\strut\\\\ text", flat


def test_service_keeps_a_line_break_before_the_text(test_parameters: TestParameters):
    document = Document()
    paragraph = document.add_paragraph()
    paragraph.add_run().add_break()
    paragraph.add_run("text")
    buffer = BytesIO()
    document.save(buffer)
    assert "\\strut\\\\ text" in _convert_docx_to_latex(test_parameters, buffer.getvalue())


def test_paragraph_with_text_is_untouched(test_parameters: TestParameters):
    assert "\\strut" not in _native_to_latex(test_parameters.container, '[ Para [ Str "text" ] ]')


def test_empty_paragraph_in_table_cell_is_untouched(test_parameters: TestParameters):
    native = (
        '[ Table ( "" , [] , [] ) ( Caption Nothing [] ) [ ( AlignDefault , ColWidthDefault ) ] '
        '( TableHead ( "" , [] , [] ) [] ) '
        '[ TableBody ( "" , [] , [] ) ( RowHeadColumns 0 ) [] '
        '[ Row ( "" , [] , [] ) [ Cell ( "" , [] , [] ) AlignDefault ( RowSpan 1 ) ( ColSpan 1 ) [ Para [] ] ] ] ] '
        '( TableFoot ( "" , [] , [] ) [] ) ]'
    )
    assert "\\strut" not in _native_to_latex(test_parameters.container, native)


def _docx(paragraphs: list[str]) -> bytes:
    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _convert_docx_to_latex(test_parameters: TestParameters, docx: bytes) -> str:
    files = {"source": ("in.docx", docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/latex", files=files)
    assert response.status_code == 200, response.text
    return " ".join(response.text.split())


def test_service_keeps_empty_paragraphs_of_a_docx(test_parameters: TestParameters):
    """Two empty paragraphs between the texts become two empty lines."""
    latex = _convert_docx_to_latex(test_parameters, _docx(["before", "", "", "after"]))
    assert latex.count("\\strut") == 2, latex


def test_service_output_differs_with_and_without_a_blank_line(test_parameters: TestParameters):
    """The gap is visible: the documents with and without it do not convert alike."""
    with_gap = _convert_docx_to_latex(test_parameters, _docx(["before", "", "after"]))
    without_gap = _convert_docx_to_latex(test_parameters, _docx(["before", "after"]))
    assert with_gap != without_gap
