"""Tests for app/docx_empty_lines_pre_process.py: the empty paragraphs of a DOCX as markers for LaTeX."""

from __future__ import annotations

import re

from app.docx_empty_lines_pre_process import EMPTY_LINES_STYLE, MAX_LINES, rewrite_part

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def _document(*paragraphs: str) -> bytes:
    return f"<w:document {W} {M}><w:body>{''.join(paragraphs)}<w:sectPr/></w:body></w:document>".encode()


def _paragraphs(*paragraphs: str) -> list[str]:
    xml, _ = rewrite_part(_document(*paragraphs))
    found = []
    for paragraph in re.findall(r"<w:p>.*?</w:p>|<w:p ?/>", xml.decode()):
        text = "".join(re.findall(r"<w:t>([^<]*)</w:t>", paragraph))
        found.append(f"{EMPTY_LINES_STYLE}:{text}" if EMPTY_LINES_STYLE in paragraph else text)
    return found


def test_an_empty_paragraph_needs_no_marker():
    source = _document("<w:p><w:r><w:t>one</w:t></w:r></w:p>", "<w:p/>", '<w:p><w:r><w:t xml:space="preserve">  </w:t></w:r></w:p>')
    assert rewrite_part(source) == (source, False)


def test_each_line_break_adds_a_line():
    assert _paragraphs('<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:br/></w:r></w:p>') == [f"{EMPTY_LINES_STYLE}:2"]
    assert _paragraphs("<w:p><w:r><w:br/><w:br/></w:r></w:p>") == [f"{EMPTY_LINES_STYLE}:3"]


def test_spaces_around_a_line_break_are_empty():
    assert _paragraphs('<w:p><w:r><w:t xml:space="preserve"> </w:t><w:br/></w:r></w:p>') == [f"{EMPTY_LINES_STYLE}:2"]


def test_the_bookmarks_of_the_paragraph_stay():
    xml, _ = rewrite_part(_document('<w:p><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:br/></w:r><w:bookmarkEnd w:id="1"/></w:p>'))
    assert 'w:name="_Toc1"' in xml.decode()
    assert "<w:bookmarkEnd" in xml.decode()


def test_the_lines_of_one_paragraph_have_a_limit():
    assert _paragraphs(f"<w:p><w:r>{'<w:br/>' * (MAX_LINES + 5)}</w:r></w:p>") == [f"{EMPTY_LINES_STYLE}:{MAX_LINES}"]


def test_a_paragraph_with_something_to_show_is_left_as_it_is():
    kept = (
        "<w:p><w:r><w:t>text</w:t><w:br/></w:r></w:p>",
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
        "<w:p><w:r><w:drawing/></w:r></w:p>",
        "<w:p><m:oMathPara/></w:p>",
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>',
        '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>',
        "<w:p><w:pPr><w:sectPr/></w:pPr></w:p>",
        "<w:p><w:r><w:tab/></w:r></w:p>",
    )
    source = _document(*kept)
    assert rewrite_part(source) == (source, False)


def test_a_paragraph_of_a_table_cell_is_left_as_it_is():
    source = _document("<w:tbl><w:tr><w:tc><w:p><w:r><w:br/></w:r></w:p></w:tc></w:tr></w:tbl>")
    assert rewrite_part(source) == (source, False)


def test_a_part_without_a_body_is_left_as_it_is():
    header = f"<w:hdr {W}><w:p><w:r><w:br/></w:r></w:p></w:hdr>".encode()
    assert rewrite_part(header) == (header, False)


def test_a_part_which_is_no_xml_is_left_as_it_is():
    assert rewrite_part(b"not xml") == (b"not xml", False)
