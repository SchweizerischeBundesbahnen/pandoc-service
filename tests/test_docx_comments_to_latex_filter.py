"""Integration tests for ``filters/docx/docx_comments_to_latex.lua``.

Word prints the comments of a document in balloons right of the text, on a
smaller page. The filter replaces the comment spans of the reader with
anchors and draws the balloons in the PDF. The unit cases feed pandoc a native
AST inside the container; the end-to-end cases send a DOCX through the service.
"""

from __future__ import annotations

import io

from docx import Document
from pypdf import PdfReader

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_comments_to_latex.lua"
MARKUP = '-M docx-comment-mag=747 -M docx-comment-area="547 134 259.4 792" -M docx-comment-balloon="572 232.8 546"'


def _comment(number: int, text: str) -> str:
    start = f'Span ( "" , [ "comment-start" ] , [ ( "comment-id" , "{number}" ) , ( "author" , "A" ) ] ) [ Str "{text}" ]'
    end = f'Span ( "" , [ "comment-end" ] , [ ( "comment-id" , "{number}" ) ] ) []'
    return f"{start} , {end}"


def _native_to_latex(container, native: str, options: str = MARKUP) -> str:
    """Convert a native AST to standalone LaTeX through the filter, whitespace collapsed."""
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f native -t latex -s {options} --lua-filter={FILTER_PATH} /tmp/test/in.native"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return " ".join(output.decode("utf-8").split())


def _body(latex: str) -> str:
    return latex.split("\\begin{document}", 1)[1]


def test_a_comment_becomes_an_anchor_and_leaves_the_text(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, f'[ Para [ Str "text" , {_comment(0, "remark")} ] ]')
    assert "\\pdcCommentStart{1}" in _body(latex)
    assert "\\pdcCommentEnd" in _body(latex)
    assert "remark" not in _body(latex)
    assert "\\@namedef{pdc@text@1}{remark}" in latex


def test_the_page_is_shrunk_by_the_magnification_given(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, f'[ Para [ Str "text" , {_comment(0, "remark")} ] ]')
    assert "\\mag=747" in latex
    assert "\\paperwidth*1000/747" in latex


def test_comments_are_numbered_in_the_order_of_the_document(test_parameters: TestParameters):
    native = f'[ Header 1 ( "h" , [] , [] ) [ Str "Title" , {_comment(5, "first")} ] , Para [ Str "text" , {_comment(2, "second")} ] ]'
    latex = _native_to_latex(test_parameters.container, native)
    assert "\\@namedef{pdc@text@1}{first}" in latex
    assert "\\@namedef{pdc@text@2}{second}" in latex
    assert "\\def\\pdc@count{2}" in latex


def test_a_paragraph_of_comments_only_keeps_its_line(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, f"[ Para [ {_comment(0, 'remark')} ] ]")
    assert "\\pdcCommentEnd}{}\\strut" in _body(latex)


def test_a_paragraph_with_text_gets_no_strut(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, f'[ Para [ Str "text" , {_comment(0, "remark")} ] ]')
    assert "\\strut" not in _body(latex)


def test_without_the_markup_the_document_is_left_as_it_is(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, f'[ Para [ Str "text" , {_comment(0, "remark")} ] ]', options="")
    assert "remark" in _body(latex)
    assert "\\pdcCommentStart" not in latex
    assert "\\mag" not in latex


def _commented_docx() -> bytes:
    document = Document()
    paragraph = document.add_paragraph("The commented text.")
    document.add_comment(paragraph.runs, text="A remark", author="Author")
    document.add_paragraph("The text after it.")
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _convert_docx_to_pdf(test_parameters: TestParameters, docx: bytes) -> PdfReader:
    files = {"source": ("in.docx", docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files=files)
    assert response.status_code == 200, response.text
    return PdfReader(io.BytesIO(response.content))


def test_service_prints_a_comment_in_a_balloon(test_parameters: TestParameters):
    reader = _convert_docx_to_pdf(test_parameters, _commented_docx())
    page = reader.pages[0]
    # The page keeps the size of the DOCX page: the content is shrunk, not the paper
    assert (round(float(page.mediabox.width)), round(float(page.mediabox.height))) == (612, 792)
    text = page.extract_text()
    assert "Commented [1]: A remark" in text
    assert "The commented text." in text
    assert "text.Aremark" not in text.replace(" ", "")


def test_service_prints_the_comments_of_a_heading_a_table_and_a_list(test_parameters: TestParameters):
    document = Document()
    heading = document.add_heading("A heading", level=1)
    document.add_comment(heading.runs, text="On the heading", author="Author")
    cell = document.add_table(rows=1, cols=2).cell(0, 1).paragraphs[0]
    cell.add_run("cell")
    document.add_comment(cell.runs, text="On the cell", author="Author")
    item = document.add_paragraph("An item", style="List Bullet")
    document.add_comment(item.runs, text="On the item", author="Author")
    buffer = io.BytesIO()
    document.save(buffer)

    text = _convert_docx_to_pdf(test_parameters, buffer.getvalue()).pages[0].extract_text()
    for number, remark in enumerate(("On the heading", "On the cell", "On the item"), start=1):
        assert f"Commented [{number}]: {remark}" in text


def test_balloons_which_reach_past_the_page_move_up_as_word_moves_them(test_parameters: TestParameters):
    """Each balloon goes below the one before it: those near the bottom of a page move up to stay on it."""
    import pypdfium2

    document = Document()
    # The last line of the first page, with more balloons than the rest of it holds
    for number in range(23):
        document.add_paragraph(f"Line {number} of the text.")
    last = document.add_paragraph("The last commented line.")
    for number in range(8):
        document.add_comment(last.runs, text=f"Remark {number}: " + "a long remark which takes several lines " * 4, author="Author")
    buffer = io.BytesIO()
    document.save(buffer)
    files = {"source": ("in.docx", buffer.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files=files)
    assert response.status_code == 200, response.text
    pdf = pypdfium2.PdfDocument(response.content)
    for index in range(len(pdf)):
        text = pdf[index].get_textpage()
        bottoms = [text.get_charbox(char)[1] for char in range(text.count_chars())]
        assert min(bottoms, default=0) >= 0, f"text below the paper on page {index + 1}"
    assert "Remark 7" in "".join(pdf[index].get_textpage().get_text_range() for index in range(len(pdf)))
