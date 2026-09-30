"""Visual tests for images in the table cells of a produced DOCX.

Each test converts HTML to DOCX through the service, the DOCX to PDF, and
compares the pages with the references (see tests/visual.py).
"""

from __future__ import annotations

import base64
import io

from PIL import Image

from tests.test_container import TestParameters
from tests.visual import assert_pages_match

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _png_data_uri(width: int, height: int, color: tuple[int, int, int]) -> str:
    """A PNG of one color with a black frame, so its edges show on the page."""
    image = Image.new("RGB", (width, height), (0, 0, 0))
    image.paste(color, (4, 4, width - 4, height - 4))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode()


def _html_to_pdf(test_parameters: TestParameters, html: str) -> bytes:
    """Convert HTML to DOCX, then that DOCX to PDF, both through the service."""
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx", data=html.encode())
    assert docx.status_code == 200, docx.text
    pdf = session.post(f"{base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.content, DOCX_MIME)})
    assert pdf.status_code == 200, pdf.text
    return pdf.content


def test_image_in_a_narrow_column_stays_in_it(test_parameters: TestParameters):
    """An 800 px image in a column of 30 %: it fits the column, not half of the page; #246."""
    html = (
        '<html><body><table style="width: 100%; border: 1px solid black; border-collapse: collapse;">'
        '<colgroup><col style="width: 30%"/><col style="width: 70%"/></colgroup>'
        f'<tr><td style="border: 1px solid black;"><img src="{_png_data_uri(800, 300, (40, 90, 170))}"/></td>'
        '<td style="border: 1px solid black;">The image on the left stays inside its column.</td></tr>'
        "</table></body></html>"
    )
    assert_pages_match("image_in_a_narrow_column", _html_to_pdf(test_parameters, html))
