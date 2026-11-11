"""Visual tests for the tables of a DOCX converted to PDF.

Each test builds a DOCX, converts it to PDF through the service, and compares
the pages with the references (see tests/visual.py).
"""

from __future__ import annotations

import io

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Twips

from tests.test_container import TestParameters
from tests.visual import assert_pages_match

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
URL = "https://example.com/polarion/project/elibrary/wiki/Specification/ProductSpecification"


def _docx_to_pdf(test_parameters: TestParameters, document) -> bytes:
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    return response.content


def _lined_table(document, widths: tuple[int, ...], rows: tuple[tuple[str, ...], ...]):
    """A table as wide as the line with the lines of a grid, its columns WIDTHS twips wide."""
    table = document.add_table(rows=len(rows), cols=len(widths))
    properties = table._tbl.tblPr
    # Without the look of the first row, every row is a row of the body
    for look in properties.findall(qn("w:tblLook")):
        properties.remove(look)
    properties.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>'))
    sides = "".join(f'<w:{side} w:val="single" w:sz="4" w:color="000000"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    properties.append(parse_xml(f"<w:tblBorders {nsdecls('w')}>{sides}</w:tblBorders>"))
    properties.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>'))
    for column, width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths, strict=True):
        column.set(qn("w:w"), str(width))
    for row, texts in zip(table.rows, rows, strict=True):
        for cell, text, width in zip(row.cells, texts, widths, strict=True):
            cell.text = text
            cell.width = Twips(width)
    return table


def test_a_word_wider_than_its_cell_breaks_at_a_character(test_parameters: TestParameters):
    """Word keeps the widths of the grid and breaks a word wider than its cell at a character; #371.

    The table of the integration tests of docx-exporter: the first column is narrower than "Resource",
    the second than the URL. Word sets them as "Resourc" and "e", and ".../ProductSpe" and "cification".
    The text of the last row fits and breaks at spaces only.
    """
    document = Document()
    document.add_paragraph("A table as wide as the line. The first column is narrower than the word in it.")
    _lined_table(document, (916, 8110), (("Resource", "Link"), ("Spec", URL), ("Note", "Specification of the product")))
    document.add_paragraph("A paragraph after the table.")

    assert_pages_match("table_word_wider_than_its_cell", _docx_to_pdf(test_parameters, document))
