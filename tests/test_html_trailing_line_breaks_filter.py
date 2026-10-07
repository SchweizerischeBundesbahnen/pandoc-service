"""Integration tests for ``filters/html_trailing_line_breaks.lua``.

Converts HTML to DOCX through the pandoc-service container, which applies the
filter to every HTML to DOCX conversion, and reads the paragraphs of the
produced ``document.xml``. A ``<br/>`` that ends a line of text must not reach
the DOCX, where Word would render it as an empty line the HTML does not show.
A ``<br/>`` between two lines, the one that makes an empty line on its own, and
the one in a paragraph that holds an image must stay.
"""

from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
import zlib

from tests.test_container import TestParameters


def _document_xml(test_parameters: TestParameters, body: str) -> str:
    """Convert ``body`` to DOCX via the pandoc-service container API and return its ``document.xml``."""
    url = f"{test_parameters.base_url}/convert/html/to/docx"
    response = test_parameters.request_session.post(url, data=f"<html><body>{body}</body></html>")
    if response.status_code // 100 != 2:
        raise AssertionError(f"pandoc-service returned {response.status_code}:\n{response.text}")
    return zipfile.ZipFile(io.BytesIO(response.content)).read("word/document.xml").decode()


def _paragraphs(test_parameters: TestParameters, body: str) -> list[tuple[str, int]]:
    """Text and number of ``<w:br/>`` of each paragraph, in document order."""
    document_xml = _document_xml(test_parameters, body)
    body_xml = document_xml[document_xml.index("<w:body>") :]
    result = []
    for match in re.finditer(r"<w:p\b.*?</w:p>", body_xml, re.DOTALL):
        para = match.group(0)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para)).strip()
        result.append((text, len(re.findall(r"<w:br\s*/>", para))))
    return result


def test_break_before_list_is_dropped(test_parameters: TestParameters):
    """Polarion's shape: text, a <br/>, then a list."""
    paragraphs = _paragraphs(test_parameters, "<div>Numbered list:<br/><ol><li>Item 1</li><li>Item 2</li></ol></div>")
    assert paragraphs == [("Numbered list:", 0), ("Item 1", 0), ("Item 2", 0)]


def test_break_after_span_is_dropped(test_parameters: TestParameters):
    paragraphs = _paragraphs(test_parameters, '<div><span class="fields"><span>EL-1</span> - </span>Bullets:<br/><ul><li>Bullet</li></ul></div>')
    assert paragraphs == [("EL-1 - Bullets:", 0), ("Bullet", 0)]


def test_break_at_end_of_paragraph_is_dropped(test_parameters: TestParameters):
    assert _paragraphs(test_parameters, "<p>text<br/></p>") == [("text", 0)]


def test_spaces_around_trailing_break_are_dropped(test_parameters: TestParameters):
    assert _paragraphs(test_parameters, "<p>text <br/> </p>") == [("text", 0)]


def test_break_at_end_of_styled_span_is_dropped(test_parameters: TestParameters):
    assert _paragraphs(test_parameters, '<p><span style="color: red"><strong>text<br/></strong></span></p>') == [("text", 0)]


def test_break_at_end_of_list_item_is_dropped(test_parameters: TestParameters):
    assert _paragraphs(test_parameters, "<ul><li>item<br/></li></ul>") == [("item", 0)]


def test_only_one_trailing_break_is_dropped(test_parameters: TestParameters):
    """Two breaks show one empty line in a browser, so one of them stays."""
    assert _paragraphs(test_parameters, "<p>text<br/><br/></p>") == [("text", 1)]


def test_break_between_lines_is_kept(test_parameters: TestParameters):
    assert _paragraphs(test_parameters, "<p>first<br/>second</p>") == [("firstsecond", 1)]


def test_empty_line_is_kept(test_parameters: TestParameters):
    """<p><br/></p> is how Polarion writes an empty line."""
    assert _paragraphs(test_parameters, "<p>text</p><p><br/></p>") == [("text", 0), ("", 1)]


def test_break_after_closed_paragraph_is_kept(test_parameters: TestParameters):
    """A <br/> after </p> is an empty line of its own, not the end of the paragraph's line."""
    assert _paragraphs(test_parameters, "<p>some paragraph</p><br/>") == [("some paragraph", 0), ("", 1)]


def test_break_after_closed_paragraph_before_list_is_kept(test_parameters: TestParameters):
    paragraphs = _paragraphs(test_parameters, "<div><p>some paragraph</p><br/><ul><li>Bullet</li></ul></div>")
    assert paragraphs == [("some paragraph", 0), ("", 1), ("Bullet", 0)]


def test_break_before_empty_anchor_is_dropped(test_parameters: TestParameters):
    """An empty anchor after the break opens no line in a browser, so the break shows nothing."""
    paragraphs = _paragraphs(test_parameters, '<div>Bullets:<br/><a id="anchor"></a><ul><li>Bullet</li></ul></div>')
    assert paragraphs == [("Bullets:", 0), ("Bullet", 0)]


def test_empty_anchor_after_dropped_break_stays_a_bookmark(test_parameters: TestParameters):
    document_xml = _document_xml(test_parameters, '<div>Bullets:<br/> <a id="anchor"></a> <ul><li>Bullet</li></ul></div>')
    assert re.search(r'<w:bookmarkStart [^>]*w:name="_anchor"', document_xml), document_xml


def test_break_after_closed_paragraph_before_empty_anchor_is_kept(test_parameters: TestParameters):
    paragraphs = _paragraphs(test_parameters, '<p>some paragraph</p><br/><a id="anchor"></a>')
    assert paragraphs == [("some paragraph", 0), ("", 1)]


def _png() -> str:
    """A 1x1 PNG as a data URI."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b"")
    return "data:image/png;base64," + base64.b64encode(png).decode()


def test_break_after_image_is_kept(test_parameters: TestParameters):
    assert _paragraphs(test_parameters, f'<p><img src="{_png()}"/><br/></p>') == [("", 1)]


def test_break_after_image_before_caption_is_kept(test_parameters: TestParameters):
    """Without the break the image paragraph holds nothing but the image, and a DOCX to PDF conversion pairs it with the caption into a figure."""
    caption = '<p class="polarion-rte-caption-paragraph">Figure <span class="polarion-rte-caption" data-sequence="Figure">1</span> First Picture</p>'
    paragraphs = _paragraphs(test_parameters, f'<p><img src="{_png()}"/><br/></p>{caption}')
    assert paragraphs == [("", 1), ("Figure 1 First Picture", 0)]
