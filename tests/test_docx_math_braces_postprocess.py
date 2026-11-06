"""Unit tests for ``app.docx_math_braces_post_process.apply_group_characters``."""

from __future__ import annotations

import re

from docx import Document
from docx.oxml import parse_xml
from lxml import etree

from app.docx_math_braces_post_process import apply_group_characters

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
EXPRESSION = "<m:e><m:r><m:t>a</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>b</m:t></m:r></m:e>"


def _doc(omath_body: str) -> Document:
    doc = Document()
    body = doc.element.find(f"{{{W_NS}}}body")
    body.append(parse_xml(f'<w:p xmlns:w="{W_NS}" xmlns:m="{M_NS}"><m:oMath>{omath_body}</m:oMath></w:p>'))
    return doc


def _math(doc: Document) -> str:
    omath = next(doc.element.iter(f"{{{M_NS}}}oMath"))
    return re.sub(r' xmlns:\w+="[^"]*"', "", etree.tostring(omath, encoding="unicode"))


def test_an_underbrace_becomes_a_group_character_under_its_expression():
    """With its label, the outer limit Word keeps: the brace stretches under all of a+b."""
    doc = _doc(f"<m:limLow><m:e><m:limLow>{EXPRESSION}<m:lim><m:r><m:t>⏟</m:t></m:r></m:lim></m:limLow></m:e><m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limLow>")
    apply_group_characters(doc)
    assert _math(doc) == (
        f'<m:oMath><m:limLow><m:e><m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/><m:vertJc m:val="top"/></m:groupChrPr>{EXPRESSION}</m:groupChr></m:e><m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limLow></m:oMath>'
    )


def test_an_overbrace_becomes_a_group_character_over_its_expression():
    doc = _doc(f"<m:limUpp>{EXPRESSION}<m:lim><m:r><m:t>⏞</m:t></m:r></m:lim></m:limUpp>")
    apply_group_characters(doc)
    assert '<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr>' in _math(doc)
    assert "limUpp" not in _math(doc)


def test_a_limit_which_is_more_than_a_brace_stays_a_limit():
    for limit in ("<m:r><m:t>n</m:t></m:r>", "<m:r><m:t>⏟</m:t></m:r><m:r><m:t>n</m:t></m:r>", "<m:sSub><m:e/><m:sub/></m:sSub>"):
        doc = _doc(f"<m:limLow>{EXPRESSION}<m:lim>{limit}</m:lim></m:limLow>")
        apply_group_characters(doc)
        assert "groupChr" not in _math(doc)


def test_a_brace_of_the_other_side_stays_a_limit():
    doc = _doc(f"<m:limLow>{EXPRESSION}<m:lim><m:r><m:t>⏞</m:t></m:r></m:lim></m:limLow>")
    apply_group_characters(doc)
    assert "groupChr" not in _math(doc)


def test_a_limit_without_an_expression_stays_a_limit():
    doc = _doc("<m:limLow><m:lim><m:r><m:t>⏟</m:t></m:r></m:lim></m:limLow>")
    apply_group_characters(doc)
    assert "groupChr" not in _math(doc)
