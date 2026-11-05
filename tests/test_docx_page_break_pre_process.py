"""Tests for app/docx_page_break_pre_process.py: the page breaks of a DOCX as markers for LaTeX."""

from __future__ import annotations

import re

from app.docx_page_break_pre_process import PAGE_BREAK_STYLE, SECTION_STYLE, rewrite_part

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
SECTION_END = '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>'


def _text(text: str, properties: str = "") -> str:
    return f"<w:p>{properties}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _paragraphs(*paragraphs: str, last_section: str = "<w:sectPr/>") -> list[str]:
    xml, _ = rewrite_part(f"<w:document {W}><w:body>{''.join(paragraphs)}{last_section}</w:body></w:document>".encode())
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
    """The type of a sectPr says how its own section starts: the one after the paragraph, here the last."""
    end = _text("one", "<w:pPr><w:sectPr/></w:pPr>")
    assert _paragraphs(end, _text("two"), last_section='<w:sectPr><w:type w:val="continuous"/></w:sectPr>') == ["one", "two"]


def test_a_section_of_another_size_switches_the_page():
    end = _text("one", '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>')
    landscape = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr>'
    assert _paragraphs(end, _text("two"), last_section=landscape) == ["one", "841.9 595.3 36 72 36 72", "two"]


def test_a_section_of_the_same_page_only_breaks_it():
    letter = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    end = _text("one", f"<w:pPr>{letter}</w:pPr>")
    assert _paragraphs(end, _text("two"), last_section=letter) == ["one", PAGE_BREAK_STYLE, "two"]


def test_the_style_of_a_section_marker_is_the_section_style():
    end = _text("one", '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>')
    xml, _ = rewrite_part(f"<w:document {W}><w:body>{end}<w:p/><w:sectPr/></w:body></w:document>".encode())
    assert f'w:val="{SECTION_STYLE}"' in xml.decode()


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
