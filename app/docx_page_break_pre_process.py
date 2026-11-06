"""Carry the page and section breaks of a DOCX into LaTeX/PDF.

pandoc's DOCX reader drops ``<w:br w:type="page"/>`` and every ``<w:sectPr>``, so
a PDF made from a DOCX ran all its pages together, on the page of the first
section. This preprocessor runs *before* pandoc and writes each break of the body
as a sentinel in the text::

    <open><close>                                  a page break
    <open>before<close>                            a paragraph of its own: the next one starts a new page
    <open>width,height,top,right,bottom,left,start<close>  a section break

``<w:pageBreakBefore/>`` in the properties of a paragraph or of its style starts a
new page before it, but none on a page with nothing on it yet, as in Word.

The numbers are twips and describe the page the break opens. ``start`` says which
page that is: ``nextPage``, ``oddPage`` or ``evenPage``. A ``<w:sectPr>`` holds the
page and the start of the section it closes, so they come from the section after
it. A continuous section break opens no page and is left alone. ``nextColumn``
moves to the next column, which is the next page only when the closed section has
one column: the PDF sets multiple columns as one, so otherwise no page opens.

The companion Lua filter (``filters/docx/docx_page_breaks_to_latex.lua``) turns the
sentinels into LaTeX. The Private-Use-Area delimiters never collide with real
document text.

``filters/html/page_orientation.lua`` writes a page break right before a section break
for one break of the editor, the way ``_drop_empty_last_page`` in
:mod:`app.docx_post_process` reads them. Here the section break takes its place,
when its paragraph is empty as that filter writes it.

Only paragraphs directly in the body count: a table cell or a list item cannot
break a page in LaTeX.
"""

from __future__ import annotations

import logging
from itertools import pairwise
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from . import docx_page_geometry
from .docx_ooxml import W_NS, parse_xml, serialize_tree

if TYPE_CHECKING:
    from collections.abc import Iterator, Mapping

logger = logging.getLogger(__name__)

# The Lua filter matches the same shape.
SENTINEL_OPEN = "\ue020"
SENTINEL_CLOSE = "\ue021"
PAGE_BREAK = f"{SENTINEL_OPEN}{SENTINEL_CLOSE}"
# A paragraph which starts a new page before it: unlike a page break, it opens none on an empty page.
PAGE_BREAK_BEFORE = f"{SENTINEL_OPEN}before{SENTINEL_CLOSE}"

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
# What a paragraph can hold and still show nothing.
_INVISIBLE_TAGS = frozenset({_PPR_TAG, f"{{{W_NS}}}bookmarkStart", f"{{{W_NS}}}bookmarkEnd"})
_TYPE_ATTR = f"{{{W_NS}}}type"
_VAL_ATTR = f"{{{W_NS}}}val"
_PAGE_BREAK_BEFORE_TAG = f"{{{W_NS}}}pageBreakBefore"
_PSTYLE_TAG = f"{{{W_NS}}}pStyle"
_OFF = frozenset({"false", "0", "off"})
_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


def _section_sentinel(page: Mapping[str, int], start: str) -> str:
    numbers = ",".join(str(page[name]) for name in ("paperwidth", "paperheight", "top", "right", "bottom", "left"))
    return f"{SENTINEL_OPEN}{numbers},{start}{SENTINEL_CLOSE}"


def _is_page_break(element: ET.Element) -> bool:
    return element.tag == _BR_TAG and element.get(_TYPE_ATTR) == "page"


def _sentinel_text(marker: str) -> ET.Element:
    text = ET.Element(_T_TAG, {_SPACE_ATTR: "preserve"})
    text.text = marker
    return text


def _closes_a_section(para: ET.Element) -> bool:
    return para.find(f"{_PPR_TAG}/{_SECTPR_TAG}") is not None


def _shows_nothing(para: ET.Element) -> bool:
    """Whether the paragraph holds nothing but properties, bookmarks and empty runs."""
    return all(child.tag in _INVISIBLE_TAGS or (child.tag == _R_TAG and all(part.tag == _RPR_TAG for part in child)) for child in para)


def _is_lone_page_break(para: ET.Element) -> bool:
    """Whether the paragraph holds page breaks and nothing else, so a section break after it can replace one of them."""
    if _closes_a_section(para):
        return False
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


def _mark_section_break(para: ET.Element, section: ET.Element | None, start: str) -> None:
    run = ET.SubElement(para, _R_TAG)
    run.append(_sentinel_text(_section_sentinel(docx_page_geometry.section_page(section), start)))


def _drop_one_page_break(body: ET.Element, para: ET.Element) -> None:
    """Remove one marked page break of a paragraph which holds nothing else, and the paragraph with the last one."""
    breaks = [(run, text) for run in para.iter(_R_TAG) for text in run.findall(_T_TAG)]
    if len(breaks) == 1:
        body.remove(para)
        return
    run, text = breaks[-1]
    run.remove(text)


def _mark_section_end(para: ET.Element, sections: Iterator[tuple[ET.Element | None, ET.Element | None]], lone_page_break: ET.Element | None, replaced_breaks: list[ET.Element]) -> bool:
    """Mark the section break the paragraph holds. An empty paragraph replaces a page break of the paragraph before it. Returns whether it opens a page."""
    closed, opened = next(sections)
    start = docx_page_geometry.page_start(closed, opened)
    if start is None:
        return False
    # Content keeps the break: it stands on the page the break opens, and the section break opens the next one.
    if lone_page_break is not None and _shows_nothing(para):
        replaced_breaks.append(lone_page_break)
    _mark_section_break(para, opened, start)
    return True


def _turned_on(element: ET.Element | None) -> bool | None:
    """Whether an on/off property is on, None when it is not there."""
    if element is None:
        return None
    return element.get(_VAL_ATTR, "true").lower() not in _OFF


def _resolve_chain(by_id: Mapping[str | None, ET.Element], style_id: str | None, found: dict[str, bool], fallback: bool) -> None:
    """Add to FOUND whether each style of the chain from STYLE_ID breaks the page before, up to a style already in it.

    A walk along the chain, not a recursion, which a long chain would exhaust. A loop in it breaks nothing.
    """
    chain: list[str] = []
    current = style_id
    while current is not None and current in by_id and current not in found and current not in chain:
        own = _turned_on(by_id[current].find(f"{_PPR_TAG}/{_PAGE_BREAK_BEFORE_TAG}"))
        if own is not None:
            found[current] = own
            break
        chain.append(current)
        based_on = by_id[current].find(f"{{{W_NS}}}basedOn")
        current = based_on.get(_VAL_ATTR) if based_on is not None else None
    found.update(dict.fromkeys(chain, found.get(current, fallback) if current is not None else fallback))


def page_break_styles(styles: ET.Element | None) -> frozenset[str | None]:
    """The paragraph styles which start a new page before their paragraphs, themselves or by the style they are based on.

    None stands for a paragraph without a style of its own: it has the default paragraph style, else the document defaults.
    """
    if styles is None:
        return frozenset()
    by_id = {style.get(f"{{{W_NS}}}styleId"): style for style in styles.iter(f"{{{W_NS}}}style")}
    # A chain without a value of its own takes the one of the document defaults
    fallback = bool(_turned_on(styles.find(f"{{{W_NS}}}docDefaults/{{{W_NS}}}pPrDefault/{_PPR_TAG}/{_PAGE_BREAK_BEFORE_TAG}")))
    found: dict[str, bool] = {}
    for style_id in by_id:
        _resolve_chain(by_id, style_id, found, fallback)
    result: set[str | None] = {style_id for style_id, breaks in found.items() if breaks}
    # Of several default paragraph styles, Word takes the last one
    defaults = [style_id for style_id, style in by_id.items() if style.get(f"{{{W_NS}}}type") == "paragraph" and style.get(f"{{{W_NS}}}default") in ("1", "true", "on")]
    if defaults[-1] in result if defaults else fallback:
        result.add(None)
    return frozenset(result)


def _breaks_before(para: ET.Element, break_styles: frozenset[str | None]) -> bool:
    """Whether Word starts a new page before the paragraph: its own property, else its style or the default one."""
    own = _turned_on(para.find(f"{_PPR_TAG}/{_PAGE_BREAK_BEFORE_TAG}"))
    if own is not None:
        return own
    style = para.find(f"{_PPR_TAG}/{_PSTYLE_TAG}")
    return (style.get(_VAL_ATTR) if style is not None else None) in break_styles


def _page_break_before() -> ET.Element:
    """A paragraph of its own for the break: the paragraph after it keeps what it holds, blank lines among them."""
    para = ET.Element(_P_TAG)
    ET.SubElement(para, _R_TAG).append(_sentinel_text(PAGE_BREAK_BEFORE))
    return para


def _mark_body_breaks(body: ET.Element, break_styles: frozenset[str | None] = frozenset()) -> bool:
    """Write the breaks of the paragraphs directly in the body as sentinels. Returns whether there were any."""
    sections = pairwise(docx_page_geometry.body_sections(body))
    changed = False
    lone_page_break: ET.Element | None = None
    replaced_breaks: list[ET.Element] = []
    breaks_before = [block for block in body if block.tag == _P_TAG and _breaks_before(block, break_styles)]
    for block in body:
        if block.tag != _P_TAG:
            if block.tag in (_TBL_TAG, _SDT_TAG):
                lone_page_break = None
            continue
        is_lone_page_break = _is_lone_page_break(block)
        changed = _mark_page_breaks(block) or changed
        if _closes_a_section(block):
            changed = _mark_section_end(block, sections, lone_page_break, replaced_breaks) or changed
        lone_page_break = block if is_lone_page_break else None
    # Not in the loop: adding or removing a block while iterating the body visits a block twice or skips one.
    for para in breaks_before:
        body.insert(list(body).index(para), _page_break_before())
    for para in replaced_breaks:
        _drop_one_page_break(body, para)
    return changed or bool(breaks_before)


def rewrite_part(xml_bytes: bytes, break_styles: frozenset[str | None] = frozenset()) -> tuple[bytes, bool]:
    """Write the breaks of the body as sentinels. A part without a body has no page to break. Returns (new_bytes, changed).

    ``break_styles`` are the paragraph styles which start a new page before their paragraphs, see page_break_styles.
    """
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping page break preprocess")
        return xml_bytes, False
    body = tree.find(_BODY_TAG)
    if body is None or not _mark_body_breaks(body, break_styles):
        return xml_bytes, False
    return serialize_tree(tree), True
