"""A table nested in a table cell reaches the PDF, written as a tabular; #278.

pandoc writes a nested table as a longtable inside the minipage of its cell, which LaTeX cannot set:
the TeX run failed with "TeX capacity exceeded".
"""

from __future__ import annotations

import io

from pypdf import PdfReader

from tests.test_container import TestParameters

NESTED = '<table><tr><td>Outer</td><td>Details:<table><tr><td style="background-color: #eeeeee">Status</td><td>Approved</td></tr><tr><td>Owner</td><td>Admin</td></tr></table></td></tr></table>'


def _text(pdf: bytes) -> str:
    return " ".join(page.extract_text() for page in PdfReader(io.BytesIO(pdf)).pages)


def _post(test_parameters: TestParameters, path: str, **kwargs) -> bytes:
    response = test_parameters.request_session.post(f"{test_parameters.base_url}{path}", **kwargs)
    assert response.status_code == 200, response.text
    return response.content


def test_a_nested_table_of_html_reaches_the_pdf(test_parameters: TestParameters):
    pdf = _post(test_parameters, "/convert/html/to/pdf", data=NESTED.encode())

    text = _text(pdf)
    assert all(word in text for word in ("Outer", "Details", "Status", "Approved", "Owner", "Admin")), text


def test_a_nested_table_of_a_docx_reaches_the_pdf(test_parameters: TestParameters):
    docx = _post(test_parameters, "/convert/html/to/docx?preserve_table_styles=false", data=NESTED.encode())

    pdf = _post(test_parameters, "/convert/docx/to/pdf", files={"source": ("nested.docx", docx)})

    text = _text(pdf)
    assert all(word in text for word in ("Outer", "Details", "Status", "Approved", "Owner", "Admin")), text
    assert "\ue010" not in text and "\ue011" not in text


def test_a_nested_table_with_a_caption_or_of_the_float_class_reaches_the_pdf(test_parameters: TestParameters):
    html = '<table><tr><td>Outer</td><td><table><caption>Inner caption</caption><tr><td>Captioned</td></tr></table><table class="float"><tr><td>Floating</td></tr></table></td></tr></table>'

    pdf = _post(test_parameters, "/convert/html/to/pdf", data=html.encode())

    text = _text(pdf)
    assert all(word in text for word in ("Outer", "Inner caption", "Captioned", "Floating")), text


def _png_uri() -> str:
    import base64

    from PIL import Image

    buffer = io.BytesIO()
    Image.new("RGB", (40, 30), (200, 0, 0)).save(buffer, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode()


def _images(pdf: bytes) -> int:
    return sum(len(page.images) for page in PdfReader(io.BytesIO(pdf)).pages)


def test_what_only_a_nested_table_holds_gets_its_packages(test_parameters: TestParameters):
    """An image, struck-out text and a merged row only inside a nested table: the preamble loads their packages."""
    html = f'<table><tr><td>Outer</td><td><table><tr><td><img src="{_png_uri()}"/></td><td><del>Struck</del></td><td rowspan="2">Merged</td></tr><tr><td>Left</td><td>Right</td></tr></table></td></tr></table>'

    pdf = _post(test_parameters, "/convert/html/to/pdf", data=html.encode())

    text = _text(pdf)
    assert all(word in text for word in ("Outer", "Struck", "Merged", "Left", "Right")), text
    assert _images(pdf) == 1


def test_an_image_only_in_a_nested_table_of_a_docx_reaches_the_pdf(test_parameters: TestParameters):
    html = f'<table><tr><td>Outer</td><td><table><tr><td><img src="{_png_uri()}"/></td><td>Beside</td></tr></table></td></tr></table>'
    docx = _post(test_parameters, "/convert/html/to/docx", data=html.encode())

    pdf = _post(test_parameters, "/convert/docx/to/pdf", files={"source": ("nested.docx", docx)})

    assert "Beside" in _text(pdf)
    assert _images(pdf) == 1
