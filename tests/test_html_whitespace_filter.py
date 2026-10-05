"""Integration tests for ``filters/html_whitespace.lua``.

Converts HTML to DOCX through the pandoc-service container, which applies the
filter to every HTML to DOCX conversion, and reads the text of each paragraph
of the produced ``document.xml``. Whitespace must collapse as in a browser:
one space for a run that crosses element boundaries, none at the start or the
end of a line. Non-breaking spaces stay.
"""

from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
import zlib

from tests.test_container import TestParameters


def _png() -> str:
    """A 1x1 PNG as a data URI."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b"")
    return "data:image/png;base64," + base64.b64encode(png).decode()


def _texts(test_parameters: TestParameters, body: str) -> list[str]:
    """The text of each paragraph, in document order. An image reads ``[img]``, a line break ``\\n``."""
    url = f"{test_parameters.base_url}/convert/html/to/docx"
    response = test_parameters.request_session.post(url, data=f"<html><body>{body}</body></html>")
    if response.status_code // 100 != 2:
        raise AssertionError(f"pandoc-service returned {response.status_code}:\n{response.text}")
    document_xml = zipfile.ZipFile(io.BytesIO(response.content)).read("word/document.xml").decode()
    body_xml = document_xml[document_xml.index("<w:body>") :]
    result = []
    for match in re.finditer(r"<w:p\b.*?</w:p>", body_xml, re.DOTALL):
        parts = re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>|(<w:drawing)|(<w:br\s*/>)", match.group(0))
        result.append("".join(text or ("[img]" if drawing else "\n") for text, drawing, _ in parts))
    return result


def test_polarion_linked_work_item_has_single_spaces(test_parameters: TestParameters):
    """Polarion's indented markup: whitespace on both sides of </span>, and after the opening <span>."""
    link = (
        '<span class="polarion-no-style-cleanup" style="white-space:nowrap;" title="EL-101 - User name">'
        '<a class="polarion-Hyperlink" href="#work-item-anchor-elibrary/EL-101">'
        f'<span style="white-space:nowrap;"><img src="{_png()}" class="polarion-Icons"/></span>'
        '<span style="color:#000000;">EL-101</span><span style="white-space: normal"> - User name</span>'
        "</a></span>"
    )
    body = f"""<table><tr><td>
        <span style="display:inline-block;white-space:nowrap;">
         <span class="polarion-no-style-cleanup" title="Used to link implementing Tasks and Issues">implements</span>
        </span>
        : {link}
    </td></tr></table>"""
    assert _texts(test_parameters, body) == ["implements : [img]EL-101 - User name"]


def test_run_across_span_boundaries_is_one_space(test_parameters: TestParameters):
    assert _texts(test_parameters, '<p><span class="a">one </span> <span class="b"> two</span></p>') == ["one two"]


def test_spaces_at_start_and_end_of_paragraph_are_dropped(test_parameters: TestParameters):
    assert _texts(test_parameters, "<p> <span> <a id='anchor'></a> </span> text <span> </span></p>") == ["text"]


def test_spaces_around_line_break_are_dropped(test_parameters: TestParameters):
    assert _texts(test_parameters, "<p>first <span> </span><br/> second</p>") == ["first\nsecond"]


def test_spaces_next_to_image_are_kept(test_parameters: TestParameters):
    assert _texts(test_parameters, f'<p>before <img src="{_png()}"/> after</p>') == ["before [img] after"]


def test_non_breaking_spaces_are_kept(test_parameters: TestParameters):
    assert _texts(test_parameters, "<p>a&nbsp;&nbsp;b&nbsp;</p>") == ["a\u00a0\u00a0b\u00a0"]


def test_preformatted_span_keeps_its_leading_space(test_parameters: TestParameters):
    """A browser keeps the spaces of white-space: pre, also at the start of a line."""
    assert _texts(test_parameters, '<p><span style="white-space: pre"> indented</span></p>') == [" indented"]


def test_run_across_span_boundary_in_heading_is_one_space(test_parameters: TestParameters):
    assert _texts(test_parameters, "<h2><span>Chapter </span> Title</h2>") == ["Chapter Title"]


def test_preformatted_div_keeps_spaces_of_its_spans(test_parameters: TestParameters):
    """White-space is inherited: a pre-wrap <div> keeps the spaces of the spans inside it."""
    body = '<div style="white-space: pre-wrap"><p><span>one </span> <span> two</span></p></div>'
    assert _texts(test_parameters, body) == ["one   two"]


def test_div_setting_whitespace_back_to_normal_collapses_again(test_parameters: TestParameters):
    body = '<div style="white-space: pre-wrap"><div style="white-space: normal"><p><span>one </span> <span> two</span></p></div></div>'
    assert _texts(test_parameters, body) == ["one two"]


def test_empty_preformatted_span_does_not_keep_the_space_after_it(test_parameters: TestParameters):
    """An empty element ends no whitespace run, whatever its style, so a bookmark at the start of a line leaves no space."""
    assert _texts(test_parameters, '<p><span style="white-space: pre"><a id="anchor"></a></span> text</p>') == ["text"]


def test_span_setting_whitespace_back_to_normal_collapses_its_runs(test_parameters: TestParameters):
    body = '<div style="white-space: pre-wrap"><p><span style="white-space: normal"><span>one </span> <span> two</span></span></p></div>'
    assert _texts(test_parameters, body) == ["one two"]


def test_preformatted_div_inside_normal_div_inside_preformatted_div_keeps_spaces(test_parameters: TestParameters):
    """Each style change is seen top down, so the innermost pre-wrap applies."""
    body = '<div style="white-space: pre-wrap"><div style="white-space: normal"><div style="white-space: pre-wrap"><p><span>one </span> <span> two</span></p></div></div></div>'
    assert _texts(test_parameters, body) == ["one   two"]
