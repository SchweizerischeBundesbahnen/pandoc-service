"""Mark the page breaks and sections of a DOCX so LaTeX breaks the page where Word does.

pandoc's DOCX reader drops page breaks and sections: a paragraph holding only
``<w:br w:type="page"/>``, a section which starts on a new page, and a
paragraph with ``<w:pageBreakBefore/>`` all ran on in the PDF, and every page
had the size and margins of the first section. This preprocessor puts a
paragraph in a synthetic style at each of them, which
filters/docx_document_styles_to_latex.lua turns into LaTeX:

* ``PandocPageBreak``: a new page, ``\\clearpage``;
* ``PandocSection``: a section on a new page of another size or other margins,
  its text the page width, height and the top, right, bottom and left margins
  in points, which switch the page geometry.

Two breaks in a row make one new page, in LaTeX as in Word: ``\\clearpage``
on an empty page does nothing. A page break among the text of a paragraph is
left as it is.

A ``<w:sectPr>`` describes the section it ends: the one in the properties of a
paragraph the section up to that paragraph, the one of the body the last
section. Its ``<w:type>`` says how that section starts.
"""

from __future__ import annotations

from xml.etree import ElementTree as ET

from .docx_ooxml import W_NS, parse_xml, serialize_tree
from .docx_page_geometry import LETTER, ONE_INCH_MARGINS

PAGE_BREAK_STYLE = "PandocPageBreak"
SECTION_STYLE = "PandocSection"
TWIPS_PER_POINT = 20
# A section without a type starts on a new page, as one of these does.
_NEW_PAGE_SECTIONS = frozenset({"nextPage", "evenPage", "oddPage"})
_OFF_VALUES = frozenset({"false", "0", "off"})


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _marker(style: str, text: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    ET.SubElement(properties, _w("pStyle"), {_w("val"): style})
    run = ET.SubElement(paragraph, _w("r"))
    ET.SubElement(run, _w("t")).text = text
    return paragraph


def _only_a_page_break(paragraph: ET.Element) -> bool:
    """Whether the paragraph holds a page break and no text."""
    breaks = [br for br in paragraph.iter(_w("br")) if br.get(_w("type")) == "page"]
    text = "".join(t.text or "" for t in paragraph.iter(_w("t")))
    return bool(breaks) and not text.strip()


def _section_end(element: ET.Element) -> ET.Element | None:
    """The sectPr which a paragraph ends its section with."""
    return element.find(f"{_w('pPr')}/{_w('sectPr')}") if element.tag == _w("p") else None


def _starts_on_a_new_page(section: ET.Element | None) -> bool:
    if section is None:
        return False
    kind = section.find(_w("type"))
    return (kind.get(_w("val"), "nextPage") if kind is not None else "nextPage") in _NEW_PAGE_SECTIONS


def _twips(element: ET.Element | None, attribute: str, fallback: int) -> int:
    value = element.get(_w(attribute)) if element is not None else None
    try:
        return abs(int(value)) if value is not None else fallback
    except ValueError:
        return fallback


def _geometry(section: ET.Element | None) -> tuple[int, ...]:
    """The page width and height and the top, right, bottom and left margins of a section, in twips."""
    size = section.find(_w("pgSz")) if section is not None else None
    margins = section.find(_w("pgMar")) if section is not None else None
    return (
        _twips(size, "w", LETTER["w"]),
        _twips(size, "h", LETTER["h"]),
        *(_twips(margins, side, ONE_INCH_MARGINS[side]) for side in ("top", "right", "bottom", "left")),
    )


def _is_empty(element: ET.Element) -> bool:
    return not "".join(t.text or "" for t in element.iter(_w("t"))).strip()


def _breaks_before(paragraph: ET.Element) -> bool:
    flag = paragraph.find(f"{_w('pPr')}/{_w('pageBreakBefore')}")
    return flag is not None and flag.get(_w("val"), "true") not in _OFF_VALUES


def _section_start(section: ET.Element, previous: tuple[int, ...]) -> ET.Element | None:
    """The marker which starts a section: its geometry when it differs, else a page break, else none."""
    geometry = _geometry(section)
    if geometry != previous:
        return _marker(SECTION_STYLE, " ".join(f"{twips / TWIPS_PER_POINT:g}" for twips in geometry))
    return _marker(PAGE_BREAK_STYLE, PAGE_BREAK_STYLE) if _starts_on_a_new_page(section) else None


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Mark the page breaks and sections among the paragraphs of the body; returns the part and whether it changed."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        return xml_bytes, False
    body = tree.find(_w("body"))
    if body is None:
        return xml_bytes, False
    children = list(body)
    # Each section ends with its sectPr: the paragraphs which end one, then the body's own for the last
    ends = [section for section in (_section_end(child) for child in children) if section is not None]
    sections = [*ends, body.find(_w("sectPr"))]
    changed = False
    number = 0
    for index, paragraph in enumerate(children):
        if paragraph.tag != _w("p"):
            continue
        position = list(body).index(paragraph)
        following = children[index + 1] if index + 1 < len(children) else None
        ends_a_section = _section_end(paragraph) is not None
        next_section = sections[number + 1] if ends_a_section and number + 1 < len(sections) else None
        if _only_a_page_break(paragraph) and following is not None and _section_end(following) is not None and _is_empty(following):
            # Word makes one new page of a page break before the empty paragraph which ends a section
            continue
        if _only_a_page_break(paragraph) and not ends_a_section:
            body.remove(paragraph)
            body.insert(position, _marker(PAGE_BREAK_STYLE, PAGE_BREAK_STYLE))
            changed = True
            continue
        if _breaks_before(paragraph):
            body.insert(position, _marker(PAGE_BREAK_STYLE, PAGE_BREAK_STYLE))
            position += 1
            changed = True
        if ends_a_section:
            start = _section_start(next_section, _geometry(sections[number])) if next_section is not None else None
            if start is not None:
                body.insert(position + 1, start)
                changed = True
            number += 1
    return (serialize_tree(tree), True) if changed else (xml_bytes, False)
