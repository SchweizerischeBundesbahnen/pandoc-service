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


def test_bookmark_of_an_empty_paragraph_is_kept(test_parameters: TestParameters):
    """A link to the bookmark still has its target."""
    flat = _native_to_latex(test_parameters.container, '[ Para [ Span ( "_Toc1" , [ "anchor" ] , [] ) [] ] ]')
    assert "\\label{_Toc1}" in flat, flat
    assert "\\strut" in flat, flat


def test_empty_paragraph_inside_div_becomes_strut(test_parameters: TestParameters):
    native = '[ Div ( "" , [] , [ ( "custom-style" , "Body Text" ) ] ) [ Para [] ] ]'
    assert "\\strut" in _native_to_latex(test_parameters.container, native)


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
