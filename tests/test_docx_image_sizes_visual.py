"""Visual tests for the size of the images in a produced DOCX.

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


def test_image_taller_than_the_page_fits_the_page(test_parameters: TestParameters):
    """A 300 x 3000 px image fills the page from the top to the bottom margin and keeps its shape; #245.

    Without the cap it runs far past the bottom of the page. With it the image is exactly as tall as
    the text area. LaTeX needs the depth of a line on top of that, so the image opens the second page
    and leaves the first one empty; Word puts it on the first. What the reference shows is the size.
    """
    html = f'<html><body><p><img src="{_png_data_uri(300, 3000, (40, 90, 170))}"/></p></body></html>'
    assert_pages_match("image_taller_than_the_page", _html_to_pdf(test_parameters, html))
