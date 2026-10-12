"""The data: images of an HTML document, handed to pandoc as files; #271."""

from __future__ import annotations

import base64
import io
import zipfile
from typing import TYPE_CHECKING

from app import html_embedded_images

if TYPE_CHECKING:
    from pathlib import Path

PNG = b"\x89PNG\r\n\x1a\n" + b"png bytes"
GIF = b"GIF89a" + b"gif bytes"


def _uri(subtype: str, image: bytes) -> str:
    return f"data:image/{subtype};base64,{base64.b64encode(image).decode()}"


def test_an_image_is_written_into_the_directory_and_named_in_its_src(tmp_path: Path):
    source = f'<p><img alt="a" src="{_uri("png", PNG)}" style="width: 10px"/></p>'.encode()

    result = html_embedded_images.extract(source, tmp_path)

    assert result == b'<p><img alt="a" src="pandoc-service-image-1.png" style="width: 10px"/></p>'
    assert (tmp_path / "pandoc-service-image-1.png").read_bytes() == PNG


def test_equal_images_share_one_file(tmp_path: Path):
    source = f'<img src="{_uri("png", PNG)}"><img src=\'{_uri("png", PNG)}\'><img src="{_uri("gif", GIF)}">'.encode()

    result = html_embedded_images.extract(source, tmp_path)

    assert result == b'<img src="pandoc-service-image-1.png"><img src=\'pandoc-service-image-1.png\'><img src="pandoc-service-image-2.gif">'
    assert sorted(path.name for path in tmp_path.iterdir()) == ["pandoc-service-image-1.png", "pandoc-service-image-2.gif"]


def test_a_jpeg_keeps_its_bytes_under_the_jpg_extension(tmp_path: Path):
    source = f'<IMG SRC="{_uri("JPEG", b"jpeg bytes")}">'.encode()

    assert html_embedded_images.extract(source, tmp_path) == b'<IMG SRC="pandoc-service-image-1.jpg">'
    assert (tmp_path / "pandoc-service-image-1.jpg").read_bytes() == b"jpeg bytes"


def test_base64_broken_over_lines_is_read(tmp_path: Path):
    encoded = base64.encodebytes(PNG * 20).decode()
    source = f'<img src="data:image/png;base64,{encoded}">'.encode()

    html_embedded_images.extract(source, tmp_path)

    assert (tmp_path / "pandoc-service-image-1.png").read_bytes() == PNG * 20


def test_what_is_no_raster_data_image_of_an_img_stays_as_it_is(tmp_path: Path):
    source = (
        f'<img src="{_uri("svg+xml", b"<svg/>")}">'  # pandoc converts an SVG itself
        f'<img src="data:image/png,raw">'  # not base64
        f'<img src="data:image/png;base64,">'  # empty
        f'<img src="https://example.com/a.png">'
        f'<a href="{_uri("png", PNG)}">link</a>'  # not an image of the document
        f'<img data-src="{_uri("png", PNG)}">'  # not its src
    ).encode()

    assert html_embedded_images.extract(source, tmp_path) == source
    assert list(tmp_path.iterdir()) == []


def _docx(parts: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as docx:
        for name, data in parts.items():
            docx.writestr(name, data)
    return out.getvalue()


def _parts(docx: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_a_placeholder_gets_the_data_uri_of_its_image_back(tmp_path: Path):
    (tmp_path / "pandoc-service-image-1.png").write_bytes(PNG)
    (tmp_path / "pandoc-service-image-2.jpg").write_bytes(b"jpeg bytes")
    document = b"<w:t>{{IMG:10px||pandoc-service-image-1.png}} and {{IMG:|5px|pandoc-service-image-2.jpg}}</w:t>"
    footer = b"<w:t>{{IMG:||pandoc-service-image-1.png}}</w:t>"
    media = b"\x89PNG media"

    result = _parts(html_embedded_images.restore_placeholders(_docx({"word/document.xml": document, "word/footer1.xml": footer, "word/media/a.png": media}), tmp_path))

    assert result["word/document.xml"] == f"<w:t>{{{{IMG:10px||{_uri('png', PNG)}}}}} and {{{{IMG:|5px|{_uri('jpeg', b'jpeg bytes')}}}}}</w:t>".encode()
    assert result["word/footer1.xml"] == f"<w:t>{{{{IMG:||{_uri('png', PNG)}}}}}</w:t>".encode()
    assert result["word/media/a.png"] == media


def test_a_placeholder_naming_no_extracted_image_stays_as_it_is(tmp_path: Path):
    (tmp_path / "pandoc-service-image-1.png").write_bytes(PNG)
    document = b"<w:t>{{IMG:||pandoc-service-image-7.png}} {{IMG:||../secret.png}}</w:t>"

    result = _parts(html_embedded_images.restore_placeholders(_docx({"word/document.xml": document}), tmp_path))

    assert result["word/document.xml"] == document


def test_a_document_without_placeholders_is_returned_unchanged(tmp_path: Path):
    (tmp_path / "pandoc-service-image-1.png").write_bytes(PNG)
    docx = _docx({"word/document.xml": b"<w:t>text</w:t>"})

    assert html_embedded_images.restore_placeholders(docx, tmp_path) is docx


def test_the_output_of_a_conversion_which_extracted_nothing_is_not_read(tmp_path: Path):
    assert html_embedded_images.restore_placeholders(b"%PDF-1.7 not a zip", tmp_path) == b"%PDF-1.7 not a zip"


def test_a_quoted_attribute_holding_a_greater_than_sign_does_not_hide_the_src(tmp_path: Path):
    source = f'<img title="a > b" alt=\'c > d\' src="{_uri("png", PNG)}"/>'.encode()

    result = html_embedded_images.extract(source, tmp_path)

    assert result == b'<img title="a > b" alt=\'c > d\' src="pandoc-service-image-1.png"/>'


def test_a_src_inside_another_attribute_is_not_taken_for_the_src(tmp_path: Path):
    source = f'<img title=" src=&quot;x&quot;" data-x=\' src="{_uri("png", PNG)}"\' src="{_uri("gif", GIF)}">'.encode()

    result = html_embedded_images.extract(source, tmp_path)

    assert result == f'<img title=" src=&quot;x&quot;" data-x=\' src="{_uri("png", PNG)}"\' src="pandoc-service-image-1.gif">'.encode()
