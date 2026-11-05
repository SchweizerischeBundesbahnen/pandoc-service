"""Mark the page breaks of a DOCX so LaTeX breaks the page where Word does.

pandoc's DOCX reader drops page breaks: a paragraph holding only
``<w:br w:type="page"/>``, a section which starts on a new page, and a
paragraph with ``<w:pageBreakBefore/>`` all ran on in the PDF. This
preprocessor puts a paragraph in the synthetic style ``PandocPageBreak`` at
each of them, which filters/docx_document_styles_to_latex.lua turns into
``\\clearpage``. Two breaks in a row make one new page, in LaTeX as in Word:
``\\clearpage`` on an empty page does nothing.

A page break among the text of a paragraph is left as it is.
"""

from __future__ import annotations

from xml.etree import ElementTree as ET

from .docx_ooxml import W_NS, parse_xml, serialize_tree

PAGE_BREAK_STYLE = "PandocPageBreak"
# A section without a type starts on a new page, as one of these does.
_NEW_PAGE_SECTIONS = frozenset({"nextPage", "evenPage", "oddPage"})
_OFF_VALUES = frozenset({"false", "0", "off"})


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _marker() -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    ET.SubElement(properties, _w("pStyle"), {_w("val"): PAGE_BREAK_STYLE})
    run = ET.SubElement(paragraph, _w("r"))
    ET.SubElement(run, _w("t")).text = PAGE_BREAK_STYLE
    return paragraph


def _only_a_page_break(paragraph: ET.Element) -> bool:
    """Whether the paragraph holds a page break and no text."""
    breaks = [br for br in paragraph.iter(_w("br")) if br.get(_w("type")) == "page"]
    text = "".join(t.text or "" for t in paragraph.iter(_w("t")))
    return bool(breaks) and not text.strip()


def _starts_a_new_page_after(paragraph: ET.Element) -> bool:
    """Whether the paragraph ends a section the next of which starts on a new page."""
    section = paragraph.find(f"{_w('pPr')}/{_w('sectPr')}")
    if section is None:
        return False
    kind = section.find(_w("type"))
    return (kind.get(_w("val"), "nextPage") if kind is not None else "nextPage") in _NEW_PAGE_SECTIONS


def _empty_section_end(element: ET.Element) -> bool:
    """Whether the element is a paragraph without text which ends a section on a new page."""
    text = "".join(t.text or "" for t in element.iter(_w("t")))
    return element.tag == _w("p") and not text.strip() and _starts_a_new_page_after(element)


def _breaks_before(paragraph: ET.Element) -> bool:
    flag = paragraph.find(f"{_w('pPr')}/{_w('pageBreakBefore')}")
    return flag is not None and flag.get(_w("val"), "true") not in _OFF_VALUES


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Mark the page breaks among the paragraphs of the body; returns the part and whether it changed."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        return xml_bytes, False
    body = tree.find(_w("body"))
    if body is None:
        return xml_bytes, False
    changed = False
    children = list(body)
    for index, paragraph in enumerate(children):
        if paragraph.tag != _w("p"):
            continue
        position = list(body).index(paragraph)
        following = children[index + 1] if index + 1 < len(children) else None
        if _only_a_page_break(paragraph) and following is not None and _empty_section_end(following):
            # Word makes one new page of a page break before the empty paragraph which ends a section
            continue
        if _only_a_page_break(paragraph) and not _starts_a_new_page_after(paragraph):
            body.remove(paragraph)
            body.insert(position, _marker())
            changed = True
            continue
        if _breaks_before(paragraph):
            body.insert(position, _marker())
            position += 1
            changed = True
        if _starts_a_new_page_after(paragraph):
            body.insert(position + 1, _marker())
            changed = True
    return (serialize_tree(tree), True) if changed else (xml_bytes, False)
