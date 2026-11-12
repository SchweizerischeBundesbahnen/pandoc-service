"""Visual tests for the header row of a table of a DOCX converted to PDF.

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


def _docx_to_pdf(test_parameters: TestParameters, document) -> bytes:
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    return response.content


def _table(document, alignment: str | None) -> None:
    """A table whose header row has a cell of two lines beside cells of one line, its header cells aligned as given."""
    document.add_paragraph(f"The cells of the header aligned {alignment or 'nowhere'}:")
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


def test_the_cells_of_a_header_row_take_their_vertical_alignment(test_parameters: TestParameters):
    """Word sets the cells of a header row as their alignment says, at the top without one; #372.

    The first cell takes two lines, the others one. pandoc set every cell of the header at the bottom of its row.
    """
    document = Document()
    for alignment in (None, "center", "bottom"):
        _table(document, alignment)
    document.add_paragraph("A paragraph after the tables.")

    assert_pages_match("table_head_vertical_alignment", _docx_to_pdf(test_parameters, document))
