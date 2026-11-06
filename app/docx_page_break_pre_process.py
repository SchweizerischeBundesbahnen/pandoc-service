"""Carry the page and section breaks of a DOCX into LaTeX/PDF.

pandoc's DOCX reader drops ``<w:br w:type="page"/>`` and every ``<w:sectPr>``, so
a PDF made from a DOCX ran all its pages together, on the page of the first
section. This preprocessor runs *before* pandoc and writes each break of the body
as a sentinel in the text::

    <open><close>                                  a page break
    <open>width,height,top,right,bottom,left<close>  a section break, in twips

The numbers describe the page the break opens. A ``<w:sectPr>`` holds the page of
the section it closes, so they come from the section after it. A continuous
section break opens no page and is left alone.

The companion Lua filter (``filters/docx_page_breaks_to_latex.lua``) turns the
sentinels into LaTeX. The Private-Use-Area delimiters never collide with real
document text.

``filters/page_orientation.lua`` writes a page break right before a section break
for one break of the editor, the way ``_drop_empty_last_page`` in
:mod:`app.docx_post_process` reads them. Here the section break takes its place.

Only paragraphs directly in the body count: a table cell or a list item cannot
break a page in LaTeX.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from . import docx_page_geometry
from .docx_ooxml import W_NS, parse_xml, serialize_tree

if TYPE_CHECKING:
    from collections.abc import Mapping

logger = logging.getLogger(__name__)

# The Lua filter matches the same shape.
SENTINEL_OPEN = "\ue020"
SENTINEL_CLOSE = "\ue021"
PAGE_BREAK = f"{SENTINEL_OPEN}{SENTINEL_CLOSE}"

_BODY_TAG = f"{{{W_NS}}}body"
_P_TAG = f"{{{W_NS}}}p"
_PPR_TAG = f"{{{W_NS}}}pPr"
_SECTPR_TAG = f"{{{W_NS}}}sectPr"
_R_TAG = f"{{{W_NS}}}r"
_RPR_TAG = f"{{{W_NS}}}rPr"
_T_TAG = f"{{{W_NS}}}t"
_BR_TAG = f"{{{W_NS}}}br"
_TBL_TAG = f"{{{W_NS}}}tbl"
_SDT_TAG = f"{{{W_NS}}}sdt"
_TYPE_TAG = f"{{{W_NS}}}type"
_TYPE_ATTR = f"{{{W_NS}}}type"
_VAL_ATTR = f"{{{W_NS}}}val"
_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


def _section_sentinel(page: Mapping[str, int]) -> str:
    numbers = ",".join(str(page[name]) for name in ("paperwidth", "paperheight", "top", "right", "bottom", "left"))
    return f"{SENTINEL_OPEN}{numbers}{SENTINEL_CLOSE}"


def _is_page_break(element: ET.Element) -> bool:
    return element.tag == _BR_TAG and element.get(_TYPE_ATTR) == "page"


def _sentinel_text(marker: str) -> ET.Element:
    text = ET.Element(_T_TAG, {_SPACE_ATTR: "preserve"})
    text.text = marker
    return text


def _is_lone_page_break(para: ET.Element) -> bool:
    """Whether the paragraph holds a page break and nothing else, so it can go when a section break follows."""
    runs = [child for child in para if child.tag != _PPR_TAG]
    return bool(runs) and all(run.tag == _R_TAG and all(_is_page_break(child) or child.tag == _RPR_TAG for child in run) for run in runs) and any(_is_page_break(child) for run in runs for child in run)


def _mark_page_breaks(para: ET.Element) -> bool:
    """Replace each page break of the paragraph with its sentinel, in place, so the text around it keeps its side."""
    changed = False
    for run in para.iter(_R_TAG):
        marked = False
        for index, child in enumerate(run):
            if _is_page_break(child):
                run[index] = _sentinel_text(PAGE_BREAK)
                marked = True
        run_properties = run.find(_RPR_TAG)
        # A run with properties reaches LaTeX as a command around its text, and a page break cannot stand in one.
        if marked and run_properties is not None and len(run) == len(run.findall(_T_TAG)) + 1 and all(text.text == PAGE_BREAK for text in run.findall(_T_TAG)):
            run.remove(run_properties)
        changed = changed or marked
    return changed


def _opens_a_page(section: ET.Element | None) -> bool:
    """Whether the section starts on a page of its own. Word's default is the next page."""
    start = section.find(_TYPE_TAG) if section is not None else None
    return start is None or start.get(_VAL_ATTR) != "continuous"


def _mark_section_break(para: ET.Element, section: ET.Element | None) -> None:
    run = ET.SubElement(para, _R_TAG)
    run.append(_sentinel_text(_section_sentinel(docx_page_geometry.section_page(section))))


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Write the breaks of the body as sentinels. A part without a body has no page to break. Returns (new_bytes, changed)."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping page break preprocess")
        return xml_bytes, False
    body = tree.find(_BODY_TAG)
    if body is None:
        return xml_bytes, False

    sections = docx_page_geometry.body_sections(body)
    section_index = 0
    changed = False
    lone_page_break: ET.Element | None = None
    for block in list(body):
        if block.tag in (_TBL_TAG, _SDT_TAG):
            lone_page_break = None
        if block.tag != _P_TAG:
            continue
        is_lone_page_break = _is_lone_page_break(block)
        changed = _mark_page_breaks(block) or changed
        if block.find(f"{_PPR_TAG}/{_SECTPR_TAG}") is not None:
            section_index += 1
            if _opens_a_page(sections[section_index]):
                _mark_section_break(block, sections[section_index])
                if lone_page_break is not None:
                    body.remove(lone_page_break)
                changed = True
        lone_page_break = block if is_lone_page_break else None

    return (serialize_tree(tree), True) if changed else (xml_bytes, False)
