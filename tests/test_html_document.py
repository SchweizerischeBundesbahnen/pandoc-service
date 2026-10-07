"""Tests for app.html_document: an HTML document is parsed to its end.

lxml's default HTML parser stops at libxml2's limits (about 10 MB for one
value, about 20 MB in total) without raising, which silently cut every
pre-processed document after a large embedded image.
"""

from __future__ import annotations

from app.html_document import parse_document
from tests.test_container import TestParameters

# Larger than libxml2's default limit for one value. "A" is valid base64 of zero bytes.
HUGE_SRC = "data:image/png;base64," + "A" * 11_000_000


def test_a_value_larger_than_10_mb_is_read_whole():
    doc = parse_document(f'<html><body><p>before</p><img src="{HUGE_SRC}"><p>after</p></body></html>'.encode())

    assert doc.find(".//img").get("src") == HUGE_SRC
    assert "after" in doc.text_content()


def test_a_document_larger_than_20_mb_is_read_to_its_end():
    image = '<img src="data:image/png;base64,' + "A" * 1_000_000 + '">'
    doc = parse_document(f"<html><body>{image * 25}<p>after</p></body></html>".encode())

    assert len(doc.findall(".//img")) == 25
    assert "after" in doc.text_content()


def _large_png_data_uri() -> str:
    """A valid PNG whose data: URI is over 10 MB: random pixels, stored without compression."""
    import base64
    import os
    import struct
    import zlib

    width, height = 1800, 1500

    def chunk(tag: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + tag + payload + struct.pack(">I", zlib.crc32(tag + payload))

    raw = b"".join(b"\x00" + os.urandom(width * 3) for _ in range(height))
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw, 0)) + chunk(b"IEND", b"")
    return "data:image/png;base64," + base64.b64encode(png).decode()


def test_text_after_a_large_image_reaches_the_docx(test_parameters: TestParameters):
    """The small un-sized image makes html_image_pre_process write the document back."""
    import base64
    import io
    import zipfile

    small = "data:image/png;base64," + base64.b64encode(bytes.fromhex("89504e470d0a1a0a0000000d4948445200000001000000010802000000907753de0000000c4944415408d763f8ffff3f0005fe02fea7d6a0010000000049454e44ae426082")).decode()
    large = _large_png_data_uri()
    assert len(large) > 10_000_000
    html = f'<html><body><p><img src="{small}"></p><p><img src="{large}"></p><p>Text after the image</p></body></html>'

    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/html/to/docx", data=html.encode())

    assert response.status_code == 200, response.text
    document_xml = zipfile.ZipFile(io.BytesIO(response.content)).read("word/document.xml").decode()
    assert "Text after the image" in document_xml
