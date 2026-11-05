"""Tests for app/docx_page_break_pre_process.py: the page breaks of a DOCX as markers for LaTeX."""

from __future__ import annotations

import re

from app.docx_page_break_pre_process import PAGE_BREAK_STYLE, rewrite_part

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
SECTION_END = '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>'


def _text(text: str, properties: str = "") -> str:
    return f"<w:p>{properties}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _paragraphs(*paragraphs: str) -> list[str]:
    xml, _ = rewrite_part(f"<w:document {W}><w:body>{''.join(paragraphs)}<w:sectPr/></w:body></w:document>".encode())
    found = []
    for paragraph in re.findall(r"<w:p>.*?</w:p>|<w:p ?/>", xml.decode()):
        texts = "".join(re.findall(r"<w:t>([^<]*)</w:t>", paragraph))
        found.append(texts or ("section" if "sectPr" in paragraph else ""))
    return found


def test_a_paragraph_of_a_page_break_becomes_its_marker():
    assert _paragraphs(_text("one"), PAGE_BREAK, _text("two")) == ["one", PAGE_BREAK_STYLE, "two"]


def test_a_section_on_a_new_page_breaks_after_its_last_paragraph():
    end = _text("one", '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>')
    assert _paragraphs(end, _text("two")) == ["one", PAGE_BREAK_STYLE, "two"]


def test_a_section_without_a_type_starts_on_a_new_page():
    end = _text("one", "<w:pPr><w:sectPr/></w:pPr>")
    assert _paragraphs(end, _text("two")) == ["one", PAGE_BREAK_STYLE, "two"]


def test_a_continuous_section_breaks_no_page():
    end = _text("one", '<w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr>')
    assert _paragraphs(end, _text("two")) == ["one", "two"]


def test_a_page_break_before_an_empty_section_end_makes_one_new_page():
    assert _paragraphs(_text("one"), PAGE_BREAK, SECTION_END, _text("two")) == ["one", "", "section", PAGE_BREAK_STYLE, "two"]


def test_a_page_break_before_a_paragraph_breaks_before_it():
    assert _paragraphs(_text("one"), _text("two", "<w:pPr><w:pageBreakBefore/></w:pPr>")) == ["one", PAGE_BREAK_STYLE, "two"]


def test_a_page_break_before_switched_off_breaks_nothing():
    assert _paragraphs(_text("one"), _text("two", '<w:pPr><w:pageBreakBefore w:val="0"/></w:pPr>')) == ["one", "two"]


def test_a_page_break_among_text_is_left_as_it_is():
    source = f'<w:document {W}><w:body><w:p><w:r><w:t>one</w:t><w:br w:type="page"/><w:t>two</w:t></w:r></w:p></w:body></w:document>'.encode()
    assert rewrite_part(source) == (source, False)


def test_a_part_without_a_body_is_left_as_it_is():
    header = f"<w:hdr {W}><w:p/></w:hdr>".encode()
    assert rewrite_part(header) == (header, False)
    assert rewrite_part(b"<not xml") == (b"<not xml", False)
