"""Visual tests for formulas in tables: in the DOCX of an HTML, and in the PDF of a DOCX; #374.

Each test converts its document through the service and compares the pages of the PDF with the
references (see tests/visual.py).
"""

from __future__ import annotations

import io

from docx import Document

from tests.test_container import TestParameters
from tests.test_docx_table_rules_filter import FRACTION, SUM, _compact_rows, _fraction_in_the_header
from tests.visual import assert_pages_match

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _docx_to_pdf(test_parameters: TestParameters, docx: bytes) -> bytes:
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx, DOCX_MIME)})
    assert response.status_code == 200, response.text
    return response.content


def _saved(document: Document) -> bytes:
    docx = io.BytesIO()
    document.save(docx)
    return docx.getvalue()


def test_formulas_in_a_table_with_the_styles_of_its_cells(test_parameters: TestParameters):
    """With preserve_table_styles, a formula in a cell came out as its TeX source, and a column of formulas got no room.

    The DOCX holds Word equations now: inline beside text, alone in a cell, and display in a paragraph of its own.
    The column of the quadratic formula is as wide as its numerator, which no line breaks.
    """
    cell = 'style="border: 1px solid #999999; background-color: #EEF4FA;"'
    quadratic = '<script type="math/tex; mode=display">x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}</script>'
    html = f"""<html><body>
        <p>A table with the styles of its cells, as docx-exporter sends it.</p>
        <table style="width: 100%;">
            <tr><th {cell}>Step</th><th {cell}>Formula</th><th {cell}>N</th></tr>
            <tr><td {cell}>An inline formula beside text</td><td {cell}>Area <script type="math/tex">a^2+b^2</script> of both</td><td {cell}>1</td></tr>
            <tr><td {cell}>An inline formula alone in its cell</td><td {cell}><strong><script type="math/tex">e^{{i\\pi}}+1=0</script></strong></td><td {cell}>2</td></tr>
            <tr><td {cell}>A display formula after text</td><td {cell}>Roots: {quadratic}</td><td {cell}>3</td></tr>
        </table>
        <p>A paragraph after the table.</p>
    </body></html>"""
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx", data=html.encode(), params={"preserve_table_styles": "true"})
    assert docx.status_code == 200, docx.text

    assert_pages_match("formulas_in_a_styled_table", _docx_to_pdf(test_parameters, docx.content))


def test_fractions_which_start_and_end_cells(test_parameters: TestParameters):
    """Word hangs a fraction which starts a cell from the top of its row, and sets the space of the cell below one which ends it.

    The fraction stood a line lower than the text of the row beside it, and on the line below its row.
    """
    assert_pages_match("fractions_which_start_and_end_cells", _docx_to_pdf(test_parameters, _saved(_fraction_in_the_header())))


def test_equations_which_end_rows_of_compact(test_parameters: TestParameters):
    """Word sets the row below an equation as deep as a line as below a line of text, and the space of Compact below a fraction.

    The space of Compact went below the first one twice: every row below it stood 1.8pt lower.
    """
    document = _compact_rows(SUM)
    second = _compact_rows(FRACTION).tables[0]._tbl
    document.add_paragraph("A table whose equation is a fraction.")
    document.element.body.insert(len(document.element.body) - 1, second)
    document.add_paragraph("A paragraph after the tables.")
    assert_pages_match("equations_which_end_rows_of_compact", _docx_to_pdf(test_parameters, _saved(document)))


def test_paragraphs_of_cells_in_normal(test_parameters: TestParameters):
    """Word puts the space after of Normal between two paragraphs of a cell, before and after an equation as well; #380.

    The paragraphs of a cell stood a line apart, the space only after the last one of the cell.
    """
    cell = 'style="border: 1px solid #999999;"'
    html = f"""<html><body>
        <p>A table with the styles of its cells: the paragraphs of its cells are in Normal.</p>
        <table style="width: 100%; border-collapse: collapse;">
            <tr><td {cell}>Two paragraphs</td><td {cell}><p>First paragraph:</p><p>Second paragraph</p></td></tr>
            <tr><td {cell}>A paragraph and an equation</td><td {cell}>Expected formula:<script type="math/tex; mode=display">x + x = y</script></td></tr>
            <tr><td {cell}>An equation and a paragraph</td><td {cell}>Expected formula:<script type="math/tex; mode=display">x + x = y</script>After the formula</td></tr>
            <tr><td {cell}>An equation which starts its cell</td><td {cell}><script type="math/tex; mode=display">x + x = y</script>After the formula</td></tr>
            <tr><td {cell}>One paragraph</td><td {cell}>Row</td></tr>
        </table>
        <p>A paragraph after the table.</p>
    </body></html>"""
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx", data=html.encode(), params={"preserve_table_styles": "true"})
    assert docx.status_code == 200, docx.text

    assert_pages_match("paragraphs_of_cells_in_normal", _docx_to_pdf(test_parameters, docx.content))
