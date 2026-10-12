"""The data: images of an HTML document reach the DOCX through files and the media bag; #271.

pandoc gets each image as a file of the conversion, which filters/html_embedded_images.lua loads.
The images in a styled span and in a table with preserved styles take the {{IMG:...}} placeholder
path of filters/inline_styles.lua instead, which app.html_embedded_images restores.
"""

from __future__ import annotations

import base64
import io
import zipfile

from PIL import Image

from tests.test_container import TestParameters


def _image(color: tuple[int, int, int], image_format: str) -> bytes:
    buffer = io.BytesIO()
    Image.new("RGB", (40, 30), color).save(buffer, format=image_format)
    return buffer.getvalue()


def _uri(subtype: str, image: bytes) -> str:
    return f"data:image/{subtype};base64,{base64.b64encode(image).decode()}"


def _convert(test_parameters: TestParameters, html: str, query: str = "") -> bytes:
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/html/to/docx{query}", data=html.encode())
    assert response.status_code == 200, response.text
    return response.content


def _media(docx: bytes) -> list[bytes]:
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        return [archive.read(name) for name in archive.namelist() if name.startswith("word/media/")]


def _drawings(docx: bytes) -> int:
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        return archive.read("word/document.xml").decode().count("<wp:inline")


def test_the_images_of_a_document_reach_the_docx_as_they_are(test_parameters: TestParameters):
    red, green, blue = _image((255, 0, 0), "PNG"), _image((0, 255, 0), "GIF"), _image((0, 0, 255), "JPEG")
    html = f'<p><img src="{_uri("png", red)}"/></p><table><tr><td><img src="{_uri("gif", green)}"/></td><td><img src="{_uri("jpeg", blue)}"/></td></tr></table><p><img src="{_uri("png", red)}"/></p>'

    docx = _convert(test_parameters, html)

    assert sorted(_media(docx)) == sorted([red, green, blue])
    assert _drawings(docx) == 4


def test_an_image_in_a_styled_span_and_in_a_styled_table_reaches_the_docx(test_parameters: TestParameters):
    red, green = _image((255, 0, 0), "PNG"), _image((0, 255, 0), "PNG")
    html = f'<p><span style="color: #00ff00">Text <img src="{_uri("png", red)}"/></span></p><table><tr><td style="background-color: #eeeeee; border: 1px solid #cccccc">Cell <img src="{_uri("png", green)}"/></td></tr></table>'

    docx = _convert(test_parameters, html, "?preserve_table_styles=true")

    assert sorted(_media(docx)) == sorted([red, green])
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        assert "{{IMG:" not in archive.read("word/document.xml").decode()


def test_a_document_naming_an_image_file_itself_gets_no_image(test_parameters: TestParameters):
    """Neither a path of the container nor a name of the form the service gives the images of the conversion."""
    red = _image((255, 0, 0), "PNG")
    html = f'<p><img src="pandoc-service-image-1.png"/><img src="pandoc-service-image-0123abcd-1.png"/><img src="/etc/hostname"/></p><p><img src="{_uri("png", red)}"/></p>'

    docx = _convert(test_parameters, html)

    assert _media(docx) == [red]
    assert _drawings(docx) == 1


def test_the_images_of_a_document_with_a_base_address_reach_the_docx(test_parameters: TestParameters):
    """pandoc resolves a src against <base href>, a data: URI aside, so the filter gets the name inside an address."""
    red, green = _image((255, 0, 0), "PNG"), _image((0, 255, 0), "PNG")
    html = f'<html><head><base href="https://polarion.example/polarion/"/></head><body><p><img src="{_uri("png", red)}"/></p><p><span style="color: #00ff00">Text <img src="{_uri("png", green)}"/></span></p></body></html>'

    docx = _convert(test_parameters, html)

    assert sorted(_media(docx)) == sorted([red, green])
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        assert "{{IMG:" not in archive.read("word/document.xml").decode()
