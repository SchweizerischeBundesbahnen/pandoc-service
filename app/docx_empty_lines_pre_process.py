"""Mark the paragraphs of a DOCX which hold only line breaks, so LaTeX keeps all their lines.

Word shows a paragraph without text as an empty line, and one more line for
each line break (``<w:br/>``) in it. Polarion writes such paragraphs between
the blocks of a document. pandoc's DOCX reader drops the line breaks at the end
of a paragraph, so filters/docx/docx_empty_paragraphs_to_latex.lua saw an empty
paragraph and set one line where Word sets two or more; the pages of the PDF
broke elsewhere.

This preprocessor replaces each paragraph of the body which holds line breaks
and nothing else with a paragraph in the synthetic style ``PandocEmptyLines``.
Its text is the number of lines and the id of its own style, and its bookmarks stay. The filter sets that
many empty lines. An empty paragraph without a line break needs no marker.

A table cell keeps its own height. A paragraph with a picture, an equation, a
field, a list number, a page break or the end of a section is left as it is.

The reader drops a line break at the start or the end of a paragraph with text
as well, where Word shows an empty line. A run of the text ``PDCKEEPBREAK``
next to such a break keeps it, also in a table cell, and the filter sets the
marker as an empty line of its own.
"""

from __future__ import annotations

import re
from xml.etree import ElementTree as ET

from .docx_ooxml import W_NS, parse_xml, serialize_tree

EMPTY_LINES_STYLE = "PandocEmptyLines"
# Word sets as many lines as a paragraph holds; LaTeX gets no more than this.
MAX_LINES = 100
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
# What Word shows in a paragraph without text, or what another rewrite handles.
_CONTENT = frozenset(
    {
        f"{{{W_NS}}}drawing",
        f"{{{W_NS}}}pict",
        f"{{{W_NS}}}object",
        f"{{{W_NS}}}fldChar",
        f"{{{W_NS}}}fldSimple",
        f"{{{W_NS}}}instrText",
        f"{{{W_NS}}}sym",
        f"{{{W_NS}}}tab",
        f"{{{W_NS}}}numPr",
        f"{{{W_NS}}}sectPr",
        f"{{{M_NS}}}oMath",
        f"{{{M_NS}}}oMathPara",
    }
)
_BOOKMARKS = frozenset({f"{{{W_NS}}}bookmarkStart", f"{{{W_NS}}}bookmarkEnd"})
KEEP_BREAK = "PDCKEEPBREAK"


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _lines(paragraph: ET.Element) -> int | None:
    """The lines Word shows for a paragraph of line breaks, None for any other paragraph."""
    breaks = 0
    for element in paragraph.iter():
        if element.tag in _CONTENT:
            return None
        if element.tag == _w("t") and (element.text or "").strip():
            return None
        if element.tag == _w("br"):
            if element.get(_w("type"), "textWrapping") != "textWrapping":
                return None
            breaks += 1
    return min(breaks + 1, MAX_LINES) if breaks else None


def _base_style(paragraph: ET.Element) -> str:
    """The id of the paragraph's own style, letters and digits only; Normal without one."""
    style = paragraph.find(f"{_w('pPr')}/{_w('pStyle')}")
    value = re.sub(r"[^A-Za-z0-9]", "", style.get(_w("val")) or "") if style is not None else ""
    return value or "Normal"


def _marker(paragraph: ET.Element, lines: int) -> ET.Element:
    marker = ET.Element(_w("p"))
    properties = ET.SubElement(marker, _w("pPr"))
    ET.SubElement(properties, _w("pStyle"), {_w("val"): EMPTY_LINES_STYLE})
    for bookmark in (element for element in paragraph.iter() if element.tag in _BOOKMARKS):
        marker.append(bookmark)
    run = ET.SubElement(marker, _w("r"))
    # The style the paragraph had, whose space before and after Word keeps around its lines
    ET.SubElement(run, _w("t")).text = f"{lines} {_base_style(paragraph)}"
    return marker


def _is_break(element: ET.Element) -> bool:
    return element.tag == _w("br") and element.get(_w("type"), "textWrapping") == "textWrapping"


def _shows(element: ET.Element) -> bool:
    """Whether Word shows something for the element of a paragraph, other than a line break."""
    if element.tag == _w("t"):
        return bool((element.text or "").strip())
    return element.tag in _CONTENT and element.tag not in {_w("numPr"), _w("sectPr")}


def _keep_run(paragraph: ET.Element, br: ET.Element, *, after: bool) -> None:
    """Put a run of the marker next to the run which holds the line break."""
    parents = {child: parent for parent in paragraph.iter() for child in parent}
    run = parents[br]
    while run.tag != _w("r") and run in parents:
        run = parents[run]
    parent = parents[run]
    marker = ET.Element(_w("r"))
    ET.SubElement(marker, _w("t")).text = KEEP_BREAK
    parent.insert(list(parent).index(run) + (1 if after else 0), marker)


def _keep_edge_breaks(paragraph: ET.Element) -> bool:
    """Mark the line breaks at the start and the end of a paragraph with text; whether it has any."""
    elements = [element for element in paragraph.iter() if _is_break(element) or _shows(element)]
    shown = [index for index, element in enumerate(elements) if not _is_break(element)]
    if not shown:
        return False
    leading, trailing = elements[: shown[0]], elements[shown[-1] + 1 :]
    if trailing:
        _keep_run(paragraph, trailing[-1], after=True)
    if leading:
        _keep_run(paragraph, leading[0], after=False)
    return bool(leading or trailing)


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Mark the paragraphs of line breaks of the body; returns the part and whether it changed."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        return xml_bytes, False
    body = tree.find(_w("body"))
    if body is None:
        return xml_bytes, False
    changed = False
    for paragraph in list(body.iter(_w("p"))):
        changed = _keep_edge_breaks(paragraph) or changed
    for index, child in enumerate(list(body)):
        lines = _lines(child) if child.tag == _w("p") else None
        if lines is not None:
            body.remove(child)
            body.insert(index, _marker(child, lines))
            changed = True
    return (serialize_tree(tree), True) if changed else (xml_bytes, False)
