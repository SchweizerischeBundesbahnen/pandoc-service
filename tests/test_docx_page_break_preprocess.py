"""Tests for ``app/docx_page_break_pre_process.py``."""

from __future__ import annotations

from xml.etree import ElementTree as ET

from app.docx_ooxml import W_NS
from app.docx_page_break_pre_process import PAGE_BREAK, PAGE_BREAK_BEFORE, SENTINEL_CLOSE, SENTINEL_OPEN, rewrite_part

PAGE_BREAK_PARAGRAPH = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
TEXT_PARAGRAPH = "<w:p><w:r><w:t>text</w:t></w:r></w:p>"
A4_PORTRAIT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1100" w:bottom="1200" w:left="1300"/>'
A4_LANDSCAPE = '<w:pgSz w:w="16838" w:h="11906"/><w:pgMar w:top="1400" w:right="1500" w:bottom="1600" w:left="1700"/>'
LANDSCAPE_SENTINEL = f"{SENTINEL_OPEN}16838,11906,1400,1500,1600,1700,nextPage{SENTINEL_CLOSE}"


def _section_paragraph(page: str, start: str = "") -> str:
    return f"<w:p><w:pPr><w:sectPr>{start}{page}</w:sectPr></w:pPr></w:p>"


def _document(body: str) -> bytes:
    return f'<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>'.encode()


def _paragraph_texts(xml: bytes) -> list[str]:
    body = ET.fromstring(xml).find(f"{{{W_NS}}}body")
    assert body is not None
    return ["".join(text.text or "" for text in para.iter(f"{{{W_NS}}}t")) for para in body.findall(f"{{{W_NS}}}p")]


def test_page_break_becomes_a_sentinel():
    xml, changed = rewrite_part(_document(TEXT_PARAGRAPH + PAGE_BREAK_PARAGRAPH + TEXT_PARAGRAPH))
    assert changed
    assert _paragraph_texts(xml) == ["text", PAGE_BREAK, "text"]


def test_page_break_keeps_its_place_in_the_text():
    para = '<w:p><w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>'
    xml, _ = rewrite_part(_document(para))
    assert _paragraph_texts(xml) == [f"before{PAGE_BREAK}after"]


def test_line_break_stays():
    document = _document("<w:p><w:r><w:br/></w:r></w:p>")
    assert rewrite_part(document) == (document, False)


def test_run_of_a_lone_page_break_loses_its_properties():
    """A run with properties reaches LaTeX as a command around the break."""
    para = '<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:br w:type="page"/></w:r></w:p>'
    xml, _ = rewrite_part(_document(para))
    assert b"w:color" not in xml


def test_run_with_text_keeps_its_properties():
    para = '<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>text</w:t><w:br w:type="page"/></w:r></w:p>'
    xml, _ = rewrite_part(_document(para))
    assert b"w:color" in xml


def test_section_break_carries_the_page_of_the_next_section():
    """A sectPr holds the page of the section it closes, so the break opens the page of the sectPr after it."""
    body = _section_paragraph(A4_PORTRAIT) + TEXT_PARAGRAPH + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"
    xml, _ = rewrite_part(_document(body))
    assert _paragraph_texts(xml) == [LANDSCAPE_SENTINEL, "text"]


def test_section_break_stands_after_the_text_of_its_paragraph():
    para = f"<w:p><w:pPr><w:sectPr>{A4_PORTRAIT}</w:sectPr></w:pPr><w:r><w:t>text</w:t></w:r></w:p>"
    xml, _ = rewrite_part(_document(para + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"))
    assert _paragraph_texts(xml) == [f"text{LANDSCAPE_SENTINEL}"]


def test_page_break_before_a_section_break_is_one_break():
    """filters/html/page_orientation.lua writes the two for one break of the editor."""
    body = TEXT_PARAGRAPH + PAGE_BREAK_PARAGRAPH + _section_paragraph(A4_PORTRAIT) + TEXT_PARAGRAPH + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"
    xml, _ = rewrite_part(_document(body))
    assert _paragraph_texts(xml) == ["text", LANDSCAPE_SENTINEL, "text"]


def test_section_break_replaces_one_of_several_page_breaks():
    """The other breaks stand for pages of their own."""
    para = '<w:p><w:r><w:br w:type="page"/><w:br w:type="page"/></w:r></w:p>'
    body = TEXT_PARAGRAPH + para + _section_paragraph(A4_PORTRAIT) + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"
    xml, _ = rewrite_part(_document(body))
    assert _paragraph_texts(xml) == ["text", PAGE_BREAK, LANDSCAPE_SENTINEL]


def test_section_break_replaces_one_page_break_across_runs():
    para = '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:br w:type="page"/></w:r></w:p>'
    xml, _ = rewrite_part(_document(para + _section_paragraph(A4_PORTRAIT) + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"))
    assert _paragraph_texts(xml) == [PAGE_BREAK * 2, LANDSCAPE_SENTINEL]


def test_page_break_before_a_section_paragraph_with_text_stays():
    """The text stands on the page the break opens, and the section break opens one more."""
    para = f"<w:p><w:pPr><w:sectPr>{A4_PORTRAIT}</w:sectPr></w:pPr><w:r><w:t>text</w:t></w:r></w:p>"
    xml, _ = rewrite_part(_document(PAGE_BREAK_PARAGRAPH + para + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"))
    assert _paragraph_texts(xml) == [PAGE_BREAK, f"text{LANDSCAPE_SENTINEL}"]


def test_page_break_before_a_section_paragraph_with_a_picture_stays():
    para = f"<w:p><w:pPr><w:sectPr>{A4_PORTRAIT}</w:sectPr></w:pPr><w:r><w:drawing/></w:r></w:p>"
    xml, _ = rewrite_part(_document(PAGE_BREAK_PARAGRAPH + para + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"))
    assert _paragraph_texts(xml) == [PAGE_BREAK, LANDSCAPE_SENTINEL]


def test_section_paragraph_with_only_a_bookmark_still_replaces_the_page_break():
    para = f'<w:p><w:pPr><w:sectPr>{A4_PORTRAIT}</w:sectPr></w:pPr><w:bookmarkStart w:id="1" w:name="b"/><w:bookmarkEnd w:id="1"/><w:r><w:rPr><w:lang w:val="en"/></w:rPr></w:r></w:p>'
    xml, _ = rewrite_part(_document(PAGE_BREAK_PARAGRAPH + para + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"))
    assert _paragraph_texts(xml) == [LANDSCAPE_SENTINEL]


def test_page_break_holding_a_section_break_stays_before_the_next_section_break():
    """Its paragraph closes a section of its own, so it is no duplicate of the next break."""
    para = f'<w:p><w:pPr><w:sectPr>{A4_PORTRAIT}</w:sectPr></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
    xml, _ = rewrite_part(_document(para + _section_paragraph(A4_LANDSCAPE) + f"<w:sectPr>{A4_PORTRAIT}</w:sectPr>"))
    portrait = f"{SENTINEL_OPEN}11906,16838,1000,1100,1200,1300,nextPage{SENTINEL_CLOSE}"
    assert _paragraph_texts(xml) == [f"{PAGE_BREAK}{LANDSCAPE_SENTINEL}", portrait]


def test_page_break_with_text_before_a_section_break_stays():
    para = '<w:p><w:r><w:t>text</w:t><w:br w:type="page"/></w:r></w:p>'
    xml, _ = rewrite_part(_document(para + _section_paragraph(A4_PORTRAIT) + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"))
    assert _paragraph_texts(xml) == [f"text{PAGE_BREAK}", LANDSCAPE_SENTINEL]


def test_page_break_with_a_table_before_the_section_break_stays():
    body = PAGE_BREAK_PARAGRAPH + "<w:tbl/>" + _section_paragraph(A4_PORTRAIT) + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"
    xml, _ = rewrite_part(_document(body))
    assert _paragraph_texts(xml) == [PAGE_BREAK, LANDSCAPE_SENTINEL]


def test_continuous_section_opens_no_page():
    """The type of a section says how that section starts, so the sectPr after the break decides."""
    continuous = '<w:type w:val="continuous"/>'
    document = _document(_section_paragraph(A4_PORTRAIT) + f"<w:sectPr>{continuous}{A4_PORTRAIT}</w:sectPr>")
    assert rewrite_part(document) == (document, False)


def test_section_type_of_the_closed_section_does_not_count():
    body = _section_paragraph(A4_PORTRAIT, '<w:type w:val="continuous"/>') + f"<w:sectPr>{A4_LANDSCAPE}</w:sectPr>"
    xml, _ = rewrite_part(_document(body))
    assert _paragraph_texts(xml) == [LANDSCAPE_SENTINEL]


def test_odd_and_even_sections_keep_their_start():
    for start in ("oddPage", "evenPage"):
        body = _section_paragraph(A4_PORTRAIT) + f'<w:sectPr><w:type w:val="{start}"/>{A4_LANDSCAPE}</w:sectPr>'
        xml, _ = rewrite_part(_document(body))
        assert _paragraph_texts(xml) == [f"{SENTINEL_OPEN}16838,11906,1400,1500,1600,1700,{start}{SENTINEL_CLOSE}"]


def test_next_column_after_one_column_is_the_next_page():
    body = _section_paragraph(A4_PORTRAIT) + f'<w:sectPr><w:type w:val="nextColumn"/>{A4_LANDSCAPE}</w:sectPr>'
    xml, _ = rewrite_part(_document(body))
    assert _paragraph_texts(xml) == [LANDSCAPE_SENTINEL]


def test_next_column_after_columns_opens_no_page():
    columns = '<w:cols w:num="2"/>'
    document = _document(_section_paragraph(A4_PORTRAIT + columns) + f'<w:sectPr><w:type w:val="nextColumn"/>{A4_PORTRAIT}</w:sectPr>')
    assert rewrite_part(document) == (document, False)


def test_paragraph_in_a_table_cell_is_left_alone():
    document = _document(f"<w:tbl><w:tr><w:tc>{PAGE_BREAK_PARAGRAPH}</w:tc></w:tr></w:tbl>")
    assert rewrite_part(document) == (document, False)


def test_part_without_a_body_is_left_alone():
    part = f'<w:ftr xmlns:w="{W_NS}">{PAGE_BREAK_PARAGRAPH}</w:ftr>'.encode()
    assert rewrite_part(part) == (part, False)


def test_unparseable_part_is_left_alone():
    assert rewrite_part(b"<broken") == (b"<broken", False)


def test_page_break_before_a_paragraph_starts_its_text():
    """Word starts a new page before a paragraph which says so in its properties."""
    para = "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>"
    xml, changed = rewrite_part(_document(TEXT_PARAGRAPH + para))
    assert changed
    assert _paragraph_texts(xml) == ["text", f"{PAGE_BREAK_BEFORE}text"]


def test_page_break_before_turned_off_breaks_nothing():
    para = '<w:p><w:pPr><w:pageBreakBefore w:val="false"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>'
    document = _document(TEXT_PARAGRAPH + para)
    assert rewrite_part(document) == (document, False)


def test_page_break_before_from_the_style_of_the_paragraph():
    para = '<w:p><w:pPr><w:pStyle w:val="Chapter"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>'
    xml, _ = rewrite_part(_document(TEXT_PARAGRAPH + para), frozenset({"Chapter"}))
    assert _paragraph_texts(xml) == ["text", f"{PAGE_BREAK_BEFORE}text"]


def test_page_break_before_a_paragraph_in_a_table_breaks_nothing():
    cell = "<w:tbl><w:tr><w:tc><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    document = _document(TEXT_PARAGRAPH + cell)
    assert rewrite_part(document) == (document, False)
