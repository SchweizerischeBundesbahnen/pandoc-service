"""Visual tests for a table nested in a table cell of a DOCX converted to PDF.

Each test converts HTML to DOCX through the service, the DOCX to PDF, and
compares the pages with the references (see tests/visual.py).
"""

from __future__ import annotations

from tests.test_container import TestParameters
from tests.visual import assert_pages_match

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
LINE = "border: 1px solid #000;"


def test_a_table_with_the_lines_of_its_style_in_a_cell(test_parameters: TestParameters):
    """Both tables keep the lines of every cell, as Word draws them.

    The inner table starts at the line of its cell and is as wide as it. The paragraph before it keeps the space
    below it. The empty paragraph which ends the cell after the inner table takes no line.
    """
    html = f"""<html><body>
        <p>A requirement table with a nested table in a cell.</p>
        <table style="border-collapse: collapse;">
        <tr><th style="{LINE} background-color: #dddddd;">Requirement</th><th style="{LINE} background-color: #dddddd;">Details</th></tr>
        <tr><td style="{LINE}">REQ-1 The system stores every change of a work item.</td>
        <td style="{LINE}">Review state:
        <table style="border-collapse: collapse;">
        <tr><td style="{LINE} background-color: #eeeeee;">Status</td><td style="{LINE}">Approved</td></tr>
        <tr><td style="{LINE} background-color: #eeeeee;">Owner</td><td style="{LINE}">Admin</td></tr>
        </table></td></tr>
        <tr><td style="{LINE}">REQ-2 A user restores an earlier revision.</td><td style="{LINE}">No details</td></tr>
        </table>
        <p>A paragraph after the table.</p>
    </body></html>"""
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx?preserve_table_styles=true", data=html.encode())
    assert docx.status_code == 200, docx.text
    pdf = session.post(f"{base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.content, DOCX_MIME)})
    assert pdf.status_code == 200, pdf.text

    assert_pages_match("nested_table_with_the_lines_of_its_style", pdf.content)
