"""Tests for app/docx_run_spaces_pre_process.py: the edge spaces of a formatted run stay inside its formatting."""

from __future__ import annotations

import re

from app.docx_run_spaces_pre_process import SPACE, rewrite_part

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _document(*runs: str) -> bytes:
    return f"<w:document {W}><w:body><w:p>{''.join(runs)}</w:p></w:body></w:document>".encode()


def _texts(*runs: str) -> list[str]:
    xml, _ = rewrite_part(_document(*runs))
    return re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml.decode())


def test_the_edge_spaces_of_a_struck_run_become_sentinels():
    run = '<w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve"> mo du </w:t></w:r>'
    assert _texts("<w:r><w:t>a</w:t></w:r>", run) == ["a", f"{SPACE}mo du{SPACE}"]


def test_a_run_of_a_space_alone_becomes_a_sentinel():
    run = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>'
    assert _texts("<w:r><w:t>a</w:t></w:r>", run) == ["a", SPACE]


def test_the_spaces_of_a_run_without_formatting_stay():
    source = _document('<w:r><w:t xml:space="preserve"> a </w:t></w:r>', '<w:r><w:rPr><w:lang w:val="de-CH"/></w:rPr><w:t xml:space="preserve"> b </w:t></w:r>')
    assert rewrite_part(source) == (source, False)


def test_a_paragraph_of_spaces_alone_stays_empty():
    source = _document('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">  </w:t></w:r>')
    assert rewrite_part(source) == (source, False)
