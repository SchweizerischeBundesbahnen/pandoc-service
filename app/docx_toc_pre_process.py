"""Rewrite the tables of contents of a DOCX so LaTeX sets them as Word does.

Word fills a table of contents field (``TOC \\o``) when it opens the document;
pandoc writes it empty and its reader drops it, so the PDF had none. A list of
figures or tables (``TOC \\f``, ``TOC \\c``) keeps the entries Word or the
exporter wrote, and pandoc printed them as plain lines with the page numbers
of a layout which is not the PDF's.

This preprocessor rewrites every TOC field of the body, every paragraph from
the one which begins it to the one which ends it:

* a table of contents becomes one paragraph in the synthetic style
  ``PandocToc__CONTENTS_<depth>``, which filters/docx/docx_toc_to_latex.lua turns
  into LaTeX's table of contents of that depth, and the empty paragraph Word
  leaves after the entries it fills in;
* each entry of a list of figures or tables becomes a paragraph in the style
  ``PandocToc__ENTRY`` holding its text and a marker run
  ``PDCPAGEREF<bookmark>``, which the filter turns into a dot leader and the
  page LaTeX sets the bookmark on.

A field with other content around it in its first or last paragraph, or of
another kind, is left as it is.
"""

from __future__ import annotations

import re
from xml.etree import ElementTree as ET

from .docx_ooxml import W_NS, parse_xml, serialize_tree

STYLE_PREFIX = "PandocToc__"
CONTENTS_STYLE = STYLE_PREFIX + "CONTENTS_"
ENTRY_STYLE = STYLE_PREFIX + "ENTRY"
PAGE_MARKER = "PDCPAGEREF"
# Word has nine levels of headings; filters/docx/docx_document_styles_to_latex.lua gives LaTeX the same.
MAX_DEPTH = 9
DEFAULT_DEPTH = 3

_OUTLINE = re.compile(r'\\o\s+"\s*(\d+)\s*-\s*(\d+)\s*"')
_LIST = re.compile(r"\\[fc]\b")
# The entries of chosen styles or of a bookmarked part: LaTeX lists the headings of the whole document only
_SELECTION = re.compile(r"\\[tb]\b")
_PAGEREF = re.compile(r"^\s*PAGEREF\s+([A-Za-z0-9_]+)")


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


class _StyleSpec:
    """A synthetic paragraph style this preprocessor references."""

    def __init__(self, style_id: str) -> None:
        self.style_id = style_id


def plain_styles(*style_ids: str) -> dict[str, _StyleSpec]:
    """Synthetic paragraph styles which carry nothing but their name, for augment_styles."""
    return {style_id: _StyleSpec(style_id) for style_id in style_ids}


def build_style_element(spec: _StyleSpec) -> ET.Element:
    """The ``<w:style>`` pandoc reads the custom-style name of the paragraph from."""
    style = ET.Element(_w("style"), {_w("type"): "paragraph", _w("customStyle"): "1", _w("styleId"): spec.style_id})
    ET.SubElement(style, _w("name"), {_w("val"): spec.style_id})
    return style


def _events(paragraph: ET.Element) -> list[tuple[str, str]]:
    """The field events of a paragraph in order: begin, instr text, separate, end, and other text."""
    events: list[tuple[str, str]] = []
    for element in paragraph.iter():
        if element.tag == _w("fldChar"):
            events.append((element.get(_w("fldCharType"), ""), ""))
        elif element.tag == _w("instrText"):
            events.append(("instr", element.text or ""))
        elif element.tag == _w("t") and (element.text or "").strip():
            events.append(("text", element.text or ""))
    return events


def _toc_fields(paragraphs: list[ET.Element]) -> list[tuple[int, int, str]]:
    """The TOC fields of the paragraphs: index of the first and last paragraph and the instruction.

    Only a field which its first paragraph begins and its last paragraph ends counts.
    """
    fields: list[tuple[int, int, str]] = []
    depth = 0
    start = -1
    instruction = ""
    clean_start = False
    for index, paragraph in enumerate(paragraphs):
        events = _events(paragraph)
        for position, (kind, value) in enumerate(events):
            if kind == "begin":
                if depth == 0:
                    start, instruction = index, ""
                    clean_start = all(event != "text" for event, _ in events[:position])
                depth += 1
            elif kind == "instr" and depth == 1:
                instruction += value
            elif kind == "end" and depth > 0:
                depth -= 1
                clean_end = all(event != "text" for event, _ in events[position + 1 :])
                if depth == 0 and clean_start and clean_end and instruction.strip().startswith("TOC"):
                    fields.append((start, index, instruction.strip()))
    return fields


def _marker_paragraph(style_id: str) -> ET.Element:
    paragraph = ET.Element(_w("p"))
    properties = ET.SubElement(paragraph, _w("pPr"))
    ET.SubElement(properties, _w("pStyle"), {_w("val"): style_id})
    run = ET.SubElement(paragraph, _w("r"))
    ET.SubElement(run, _w("t")).text = style_id
    return paragraph


def _entry_paragraph(paragraph: ET.Element) -> ET.Element | None:
    """The entry of a list of figures or tables: its text up to the tab, and the marker of its page.

    None for a paragraph with no PAGEREF, the field's own begin or end.
    """
    bookmark = None
    for instruction in paragraph.iter(_w("instrText")):
        match = _PAGEREF.match(instruction.text or "")
        if match:
            bookmark = match.group(1)
    if bookmark is None:
        return None
    entry = ET.Element(_w("p"))
    properties = ET.SubElement(entry, _w("pPr"))
    ET.SubElement(properties, _w("pStyle"), {_w("val"): ENTRY_STYLE})
    started = False
    for child in paragraph:
        if child.tag == _w("pPr"):
            continue
        # The first entry starts with the runs of the TOC field itself: begin, instruction, separate
        if not started and child.tag == _w("r") and _is_field_run(child):
            continue
        started = True
        kept = _before_tab(child)
        if kept is None:
            break
        entry.append(kept)
        if kept is not child:
            break
    run = ET.SubElement(entry, _w("r"))
    ET.SubElement(run, _w("t")).text = PAGE_MARKER + bookmark
    return entry


def _is_field_run(run: ET.Element) -> bool:
    return run.find(_w("fldChar")) is not None or run.find(_w("instrText")) is not None


def _ends_text(element: ET.Element) -> bool:
    """Whether the element is a run with a tab or a field char, which ends the text of an entry."""
    return element.tag == _w("r") and (element.find(_w("tab")) is not None or _is_field_run(element))


def _before_tab(element: ET.Element) -> ET.Element | None:
    """The element cut before the first tab and field it holds: itself when it holds none, None when it starts with one.

    A hyperlink keeps the runs before its tab; a run with a tab or a field char ends the text.
    """
    if _ends_text(element):
        return None
    if element.tag != _w("hyperlink"):
        return element
    kept = ET.Element(element.tag, element.attrib)
    for child in element:
        if _ends_text(child):
            return kept
        kept.append(child)
    return element


def _whole_outline(instruction: str) -> bool:
    """Whether the field lists the headings of the whole document from the first level down, as LaTeX does."""
    outline = _OUTLINE.search(instruction)
    return not _SELECTION.search(instruction) and (outline is None or int(outline.group(1)) <= 1)


def _depth(instruction: str) -> int:
    match = _OUTLINE.search(instruction)
    if match is None:
        return DEFAULT_DEPTH
    return max(1, min(MAX_DEPTH, int(match.group(2))))


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, dict[str, _StyleSpec]]:
    """Rewrite the TOC fields among the paragraphs of the body; returns the part and the styles it references."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        return xml_bytes, {}
    body = tree.find(_w("body"))
    if body is None:
        return xml_bytes, {}
    children = list(body)
    paragraphs = [child for child in children if child.tag == _w("p")]
    fields = _toc_fields(paragraphs)
    if not fields:
        return xml_bytes, {}
    used: dict[str, _StyleSpec] = {}
    for first, last, instruction in reversed(fields):
        field_paragraphs = paragraphs[first : last + 1]
        if _LIST.search(instruction) and not _OUTLINE.search(instruction):
            replacements = [entry for entry in (_entry_paragraph(paragraph) for paragraph in field_paragraphs) if entry is not None]
            if replacements:
                used[ENTRY_STYLE] = _StyleSpec(ENTRY_STYLE)
            if replacements:
                # Word updates the field and keeps the paragraph which ends it as an empty line, in Normal
                replacements.append(ET.Element(_w("p")))
        elif not _whole_outline(instruction):
            # The entries Word stored in the field stay as they are
            continue
        else:
            style_id = f"{CONTENTS_STYLE}{_depth(instruction)}"
            # Word fills the field with its entries and keeps the paragraph which ends it as an empty line
            replacements = [_marker_paragraph(style_id), ET.Element(_w("p"))]
            used[style_id] = _StyleSpec(style_id)
        position = list(body).index(field_paragraphs[0])
        for paragraph in field_paragraphs:
            body.remove(paragraph)
        for offset, replacement in enumerate(replacements):
            body.insert(position + offset, replacement)
    return serialize_tree(tree), used
