"""Tests for app/docx_toc_pre_process.py: the TOC fields of a DOCX as markers for LaTeX."""

from __future__ import annotations

import re

from app.docx_toc_pre_process import PAGE_MARKER, rewrite_part

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _document(*paragraphs: str) -> bytes:
    return f"<w:document {W}><w:body>{''.join(paragraphs)}<w:sectPr/></w:body></w:document>".encode()


def _field(instruction: str, result: str = "") -> str:
    return f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>'


def _entry_runs(text: str, bookmark: str) -> str:
    """An entry of a list of figures as the exporter writes it: a link, a tab and a PAGEREF field."""
    return f'<w:hyperlink w:anchor="{bookmark}"><w:r><w:t>{text}</w:t></w:r><w:r><w:tab/></w:r>{_field(f" PAGEREF {bookmark} \\h ", "<w:r><w:t>1</w:t></w:r>")}</w:hyperlink>'


def _paragraph(content: str, style: str | None = None) -> str:
    properties = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{properties}{content}</w:p>"


def _paragraph_texts(xml: bytes) -> list[str]:
    return ["".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", paragraph)) for paragraph in re.findall(r"<w:p>.*?</w:p>|<w:p ?/>", xml.decode())]


def test_an_empty_table_of_contents_becomes_its_marker():
    xml, used = rewrite_part(_document(_paragraph("<w:r><w:t>Before</w:t></w:r>"), _paragraph(_field(' TOC \\o "1-9" \\h \\z \\u ')), _paragraph("<w:r><w:t>After</w:t></w:r>")))
    assert list(used) == ["PandocToc__CONTENTS_9"]
    assert _paragraph_texts(xml) == ["Before", "PandocToc__CONTENTS_9", "", "After"], "Word keeps the paragraph which ends the field"
    assert '<w:pStyle w:val="PandocToc__CONTENTS_9" />' in xml.decode()


def test_the_depth_is_the_last_level_of_the_outline():
    _, used = rewrite_part(_document(_paragraph(_field(' TOC \\o "1-3" \\h '))))
    assert list(used) == ["PandocToc__CONTENTS_3"]


def test_a_table_of_contents_without_levels_is_three_deep():
    _, used = rewrite_part(_document(_paragraph(_field(" TOC \\h \\z "))))
    assert list(used) == ["PandocToc__CONTENTS_3"]


def test_a_table_of_contents_over_several_paragraphs_becomes_one_marker():
    first = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-2" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Old entry</w:t></w:r>'
    last = '<w:r><w:t>Old entry 2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    xml, _ = rewrite_part(_document(_paragraph(first), _paragraph(last), _paragraph("<w:r><w:t>After</w:t></w:r>")))
    assert _paragraph_texts(xml) == ["PandocToc__CONTENTS_2", "", "After"]


def test_a_list_of_figures_keeps_the_text_of_its_entries_and_marks_their_page():
    first = _field(" TOC \\h \\z \\f F ", "").replace('<w:r><w:fldChar w:fldCharType="end"/></w:r>', "") + _entry_runs("Figure 1 Caption", "_Toc1")
    second = _entry_runs("Figure 2 Caption", "_Toc2")
    end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    xml, used = rewrite_part(_document(_paragraph(first, "TOC1"), _paragraph(second, "TOC1"), _paragraph(end)))
    assert list(used) == ["PandocToc__ENTRY"]
    assert _paragraph_texts(xml) == [f"Figure 1 Caption{PAGE_MARKER}_Toc1", f"Figure 2 Caption{PAGE_MARKER}_Toc2", ""]
    assert "<w:instrText" not in xml.decode(), "the PAGEREF field is gone"
    assert xml.decode().count('w:anchor="_Toc1"') == 1, "the entry keeps its link"


def test_a_list_whose_last_entry_ends_the_field_adds_no_empty_paragraph():
    first = _field(" TOC \\h \\z \\c &quot;Table&quot; ", "").replace('<w:r><w:fldChar w:fldCharType="end"/></w:r>', "") + _entry_runs("Table 1", "_Toc3")
    last = _entry_runs("Table 2", "_Toc4") + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    xml, _ = rewrite_part(_document(_paragraph(first, "TOC1"), _paragraph(last, "TOC1")))
    assert _paragraph_texts(xml) == [f"Table 1{PAGE_MARKER}_Toc3", f"Table 2{PAGE_MARKER}_Toc4"]


def test_a_bookmark_which_is_no_name_is_not_marked():
    entry = _entry_runs("Figure 1", "_Toc1").replace("PAGEREF _Toc1", "PAGEREF }bad")
    first = _field(" TOC \\f F ", "").replace('<w:r><w:fldChar w:fldCharType="end"/></w:r>', "") + entry + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    xml, used = rewrite_part(_document(_paragraph(first, "TOC1")))
    assert used == {}
    assert PAGE_MARKER not in xml.decode()


def test_a_field_with_text_before_it_is_left_as_it_is():
    source = _document(_paragraph("<w:r><w:t>Contents: </w:t></w:r>" + _field(' TOC \\o "1-3" ')))
    assert rewrite_part(source) == (source, {})


def test_other_fields_are_left_as_they_are():
    source = _document(_paragraph(_field(" PAGE ", "<w:r><w:t>1</w:t></w:r>")))
    assert rewrite_part(source) == (source, {})


def test_a_part_without_a_body_is_left_as_it_is():
    header = f"<w:hdr {W}><w:p/></w:hdr>".encode()
    assert rewrite_part(header) == (header, {})
