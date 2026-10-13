"""With preserved table styles, a table nested in a cell stays a table; #280.

filters/inline_styles.lua rebuilds a styled table as raw OOXML, and wrote any table in one of its cells
as plain text.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET

from docx import Document

from tests.test_container import TestParameters

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _convert(test_parameters: TestParameters, html: str) -> bytes:
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/html/to/docx?preserve_table_styles=true", data=html.encode())
    assert response.status_code == 200, response.text
    return response.content


def _body(docx: bytes) -> ET.Element:
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        return ET.fromstring(archive.read("word/document.xml")).find(f"{W}body")


def _text(element: ET.Element) -> str:
    return "".join(t.text or "" for t in element.iter(f"{W}t"))


def test_a_table_nested_in_a_styled_table_stays_a_table(test_parameters: TestParameters):
    html = '<table><tr><td style="background-color: #eeeeee">Outer</td><td>Details:<table><tr><td>Status</td><td>Approved</td></tr><tr><td>Owner</td><td>Admin</td></tr></table></td></tr></table>'

    docx = _convert(test_parameters, html)

    outer = _body(docx).find(f"{W}tbl")
    assert outer is not None
    cell = outer.findall(f"{W}tr")[0].findall(f"{W}tc")[1]
    nested = cell.find(f"{W}tbl")
    assert nested is not None, "the nested table is a table in the cell"
    assert [[_text(tc) for tc in tr.findall(f"{W}tc")] for tr in nested.findall(f"{W}tr")] == [["Status", "Approved"], ["Owner", "Admin"]]
    assert list(cell)[-1].tag == f"{W}p", "a cell ends with a paragraph"
    Document(io.BytesIO(docx))


def test_two_tables_nested_in_one_cell_stay_two_tables(test_parameters: TestParameters):
    html = '<table><tr><td style="background-color: #eeeeee">Outer</td><td><table><tr><td>First</td></tr></table><table><tr><td>Second</td></tr></table></td></tr></table>'

    docx = _convert(test_parameters, html)

    cell = _body(docx).find(f"{W}tbl").findall(f"{W}tr")[0].findall(f"{W}tc")[1]
    children = [child.tag for child in cell if child.tag != f"{W}tcPr"]
    assert children == [f"{W}tbl", f"{W}p", f"{W}tbl", f"{W}p"], "a paragraph between the tables keeps them two"
    assert [_text(table) for table in cell.findall(f"{W}tbl")] == ["First", "Second"]
