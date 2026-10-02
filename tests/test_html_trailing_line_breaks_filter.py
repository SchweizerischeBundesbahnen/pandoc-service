"""Integration tests for ``filters/html_trailing_line_breaks.lua``.

Runs the real ``pandoc`` binary (html -> docx) with the filter and reads the
paragraphs of the produced ``document.xml``. A ``<br/>`` that ends a line of
text must not reach the DOCX, where Word would render it as an empty line the
HTML does not show. A ``<br/>`` between two lines, and the one that makes an
empty line on its own, must stay.
"""

from __future__ import annotations

import io
import re
import shutil
import subprocess
import zipfile

import pytest

_PANDOC = shutil.which("pandoc")
pytestmark = pytest.mark.skipif(_PANDOC is None, reason="pandoc binary not available")

_FILTER = "filters/html_trailing_line_breaks.lua"


def _paragraphs(body: str) -> list[tuple[str, int]]:
    """Text and number of ``<w:br/>`` of each paragraph, in document order."""
    html = f"<html><body>{body}</body></html>"
    completed = subprocess.run(
        [_PANDOC, "-f", "html", "-t", "docx", "--lua-filter", _FILTER, "-o", "-"],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    document_xml = zipfile.ZipFile(io.BytesIO(completed.stdout)).read("word/document.xml").decode()
    body_xml = document_xml[document_xml.index("<w:body>") :]
    result = []
    for match in re.finditer(r"<w:p\b.*?</w:p>", body_xml, re.DOTALL):
        para = match.group(0)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para)).strip()
        result.append((text, len(re.findall(r"<w:br\s*/>", para))))
    return result


def test_break_before_list_is_dropped():
    """Polarion's shape: text, a <br/>, then a list."""
    paragraphs = _paragraphs("<div>Numbered list:<br/><ol><li>Item 1</li><li>Item 2</li></ol></div>")
    assert paragraphs == [("Numbered list:", 0), ("Item 1", 0), ("Item 2", 0)]


def test_break_after_span_is_dropped():
    paragraphs = _paragraphs('<div><span class="fields"><span>EL-1</span> - </span>Bullets:<br/><ul><li>Bullet</li></ul></div>')
    assert paragraphs == [("EL-1 - Bullets:", 0), ("Bullet", 0)]


def test_break_at_end_of_paragraph_is_dropped():
    assert _paragraphs("<p>text<br/></p>") == [("text", 0)]


def test_spaces_around_trailing_break_are_dropped():
    assert _paragraphs("<p>text <br/> </p>") == [("text", 0)]


def test_break_at_end_of_styled_span_is_dropped():
    assert _paragraphs('<p><span style="color: red"><strong>text<br/></strong></span></p>') == [("text", 0)]


def test_break_at_end_of_list_item_is_dropped():
    assert _paragraphs("<ul><li>item<br/></li></ul>") == [("item", 0)]


def test_only_one_trailing_break_is_dropped():
    """Two breaks show one empty line in a browser, so one of them stays."""
    assert _paragraphs("<p>text<br/><br/></p>") == [("text", 1)]


def test_break_between_lines_is_kept():
    assert _paragraphs("<p>first<br/>second</p>") == [("firstsecond", 1)]


def test_empty_line_is_kept():
    """<p><br/></p> is how Polarion writes an empty line."""
    assert _paragraphs("<p>text</p><p><br/></p>") == [("text", 0), ("", 1)]


def test_break_after_closed_paragraph_is_kept():
    """A <br/> after </p> is an empty line of its own, not the end of the paragraph's line."""
    assert _paragraphs("<p>some paragraph</p><br/>") == [("some paragraph", 0), ("", 1)]


def test_break_after_closed_paragraph_before_list_is_kept():
    paragraphs = _paragraphs("<div><p>some paragraph</p><br/><ul><li>Bullet</li></ul></div>")
    assert paragraphs == [("some paragraph", 0), ("", 1), ("Bullet", 0)]
