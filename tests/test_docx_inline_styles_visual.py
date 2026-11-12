"""Visual tests for inline CSS that filters/html/inline_styles.lua carries into a DOCX.

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


def _html_to_pdf(test_parameters: TestParameters, html: str, params: dict[str, str] | None = None) -> bytes:
    """Convert HTML to DOCX, then that DOCX to PDF, both through the service."""
    session, base_url = test_parameters.request_session, test_parameters.base_url
    docx = session.post(f"{base_url}/convert/html/to/docx", data=html.encode(), params=params)
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


def _comment(number: int, text: str) -> str:
    """A comment as docx-exporter sends it: at one place, after the text it is made on."""
    return f'<span class="comment-start" id="{number}" author="Damien" date="2026-11-05T10:00:00Z">{text}</span><span class="comment-end" id="{number}"></span>'


def test_comments_in_a_styled_table_and_in_formatted_paragraphs(test_parameters: TestParameters):
    """A comment in a cell of a styled table, and in a paragraph with an indent or an alignment, is a comment of Word.

    With preserve_table_styles the comments of the cells were lost, and a paragraph gave up its indent and
    its alignment for its comment. Word reads five comments in this document, two of them in the table,
    and sets the second paragraph in by 600 twips and the third one in the middle.
    """
    cell = "border: 1px solid #000000;"
    html = f"""<html><body>
        <p>Plain paragraph with a comment after plain{_comment(0, "Comment on a plain paragraph")} text.</p>
        <div class="pandoc-para" data-indent-twips="600"><p>Indented paragraph with a comment after
            <span style="font-weight: bold;">bold{_comment(1, "Comment on an indented paragraph")}</span> text.</p></div>
        <div class="pandoc-para" data-text-align="center"><p>Centered paragraph with a comment after centered{_comment(2, "Comment on a centered paragraph")} text.</p></div>
        <table style="width: 100%; border-collapse: collapse;">
            <tr><th style="{cell} background-color: #CCCCCC;">Header</th><th style="{cell} background-color: #CCCCCC;">Cell with a comment</th></tr>
            <tr><td style="{cell}">Bold text</td><td style="{cell} background-color: #FFFF99;">A comment after
                <span style="font-weight: bold;">bold{_comment(3, 'Rich comment in a cell<br/><span style="font-weight: bold;">with bold</span>')}</span> text.</td></tr>
            <tr><td style="{cell}">Link</td><td style="{cell}">A comment on a
                <a href="https://example.com/x"><span class="comment-start" id="4" author="Damien" date="2026-11-05T10:00:00Z">Comment on a link in a cell</span>link<span class="comment-end" id="4"></span></a>
                in a cell.</td></tr>
        </table>
    </body></html>"""

    assert_pages_match("comments_in_a_styled_table", _html_to_pdf(test_parameters, html, {"preserve_table_styles": "true"}))
