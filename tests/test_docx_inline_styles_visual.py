"""Visual tests for inline CSS that filters/inline_styles.lua carries into a DOCX.

Each test converts HTML to DOCX through the service, the DOCX to PDF, and
compares the pages with the references (see tests/visual.py). The PDF is drawn
from the DOCX alone, so what it shows is what the DOCX holds.
"""

from __future__ import annotations

import base64
import io

from PIL import Image

from tests.test_container import TestParameters
from tests.visual import assert_pages_match

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _html_to_pdf(test_parameters: TestParameters, html: str) -> bytes:
    """Convert HTML to DOCX, then that DOCX to PDF, both through the service."""
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx", data=html.encode())
    assert docx.status_code == 200, docx.text
    pdf = session.post(f"{base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.content, DOCX_MIME)})
    assert pdf.status_code == 200, pdf.text
    return pdf.content


def _icon_data_uri(color: tuple[int, int, int]) -> str:
    """A 16 x 16 px icon of one color with a black frame, the size of Polarion's icons."""
    image = Image.new("RGB", (16, 16), (0, 0, 0))
    image.paste(color, (2, 2, 14, 14))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode()


def test_bold_font_weight_of_divs(test_parameters: TestParameters):
    """A bold font-weight on a div makes everything inside it bold; #248.

    The nested blocks inherit it as a browser renders them, until one declares a weight of its own.
    A heading is not on the page: the PDF sets every heading in bold, whatever its weight.
    """
    text = "The quick brown fox jumps over the lazy dog, and the lazy dog does not mind at all."
    html = f"""<html><body>
        <p>A paragraph without a weight. {text}</p>
        <div style="font-weight: bold;">
            <p>{text}</p>
            <ul><li>A list item inside the bold div</li><li>{text}</li></ul>
            <table><tr><td>A table cell inside the bold div</td><td>{text}</td></tr></table>
            <blockquote>{text}</blockquote>
            <div>A nested div without a weight of its own. {text}</div>
            <div style="font-weight: normal;">A nested div declared normal. {text}</div>
            <p>Bold again, with <span style="font-weight: normal;">a span declared normal</span> in it. {text}</p>
        </div>
        <div style="font-weight: 700;">A div declared 700. {text}</div>
        <div style="font-weight: 400;">A div declared 400. {text}</div>
    </body></html>"""

    assert_pages_match("bold_font_weight", _html_to_pdf(test_parameters, html))


def test_inline_image_vertical_align_and_margins(test_parameters: TestParameters):
    """vertical-align and horizontal margins of an inline image survive into the DOCX and its PDF; #410.

    A line holds each case several times over, so that a picture left on the baseline moves enough of
    the page to fail the comparison: it stands above the text, and makes its line higher than the rest.
    """
    red, green, blue = _icon_data_uri((200, 40, 40)), _icon_data_uri((40, 160, 60)), _icon_data_uri((40, 80, 200))

    def line(style: str, label: str) -> str:
        icons = " ".join(f'<img style="{style}" src="{icon}"/>{label}' for icon in (red, green, blue) * 3)
        return f"<p>{icons}</p>"

    html = f"""<html><body>
        <p>Polarion's enum icons: bottom with a 2px gap.</p>
        {line("vertical-align:bottom;border:0px;margin-right:2px;", "Draft")}
        {line("vertical-align:bottom;border:0px;margin-right:2px;", "Draft")}
        <p>The icons of linked documents, as pdf-exporter's stylesheet centers them.</p>
        {line("vertical-align:middle;margin-right:2px;", "Specification")}
        {line("vertical-align:middle;margin-right:2px;", "Specification")}
        <p>An absolute length lowers or raises an icon, a left margin keeps it off the text before it.</p>
        {line("vertical-align:-6pt;margin-left:6px;", "lowered")}
        {line("vertical-align:4pt;margin-left:6px;", "raised")}
        <p>Bold text, where the icon sits in a styled span.</p>
        <p><span style="font-weight:bold;">{"".join(f'<img style="vertical-align:bottom;margin-right:2px;" src="{icon}"/>Requirement, ' for icon in (red, green, blue) * 3)}</span></p>
        <p>A plain paragraph closes the page, so that every line above it shows where it ends.</p>
    </body></html>"""

    assert_pages_match("inline_image_layout", _html_to_pdf(test_parameters, html))
