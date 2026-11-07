"""Keep the spaces at the edges of a formatted run inside its formatting.

Word draws a space of a run with its formatting: struck through, underlined,
highlighted, and as wide as the font and size of the run make it. pandoc's DOCX
reader moves a space at the start or the end of a run out of the formatting,
so the space lost its line and took the width of the text around it: a space of
a superscript run became a space of the body size, and a line broke elsewhere.

This preprocessor writes each space at the edge of the text of a formatted run
as the sentinel U+E030. filters/docx/docx_run_spaces_to_latex.lua turns it back
into a space, inside the formatting.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from .docx_ooxml import W_NS, parse_xml, serialize_tree

if TYPE_CHECKING:
    from xml.etree import ElementTree as ET

SPACE = ""
# What a run may carry that draws or sizes nothing
_PLAIN = frozenset({"lang", "noProof", "rtl", "cs", "kern", "webHidden", "specVanish", "szCs", "bCs", "iCs", "eastAsianLayout"})


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _formatted(run: ET.Element) -> bool:
    properties = run.find(_w("rPr"))
    return properties is not None and any(child.tag.rpartition("}")[2] not in _PLAIN for child in properties)


def _sentinels(text: str) -> str:
    body = text.strip(" ")
    if not body:
        return SPACE * len(text)
    start = len(text) - len(text.lstrip(" "))
    end = len(text) - len(text.rstrip(" "))
    return SPACE * start + body + SPACE * end


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Mark the edge spaces of the formatted runs of a part; returns the part and whether it changed."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        return xml_bytes, False
    changed = False
    for paragraph in tree.iter(_w("p")):
        # A paragraph of spaces alone stays the empty paragraph pandoc reads, an empty line
        if not any((text.text or "").strip(" ") for text in paragraph.iter(_w("t"))):
            continue
        for run in paragraph.iter(_w("r")):
            if not _formatted(run):
                continue
            for text in run.iter(_w("t")):
                value = text.text or ""
                marked = _sentinels(value)
                if marked != value:
                    text.text = marked
                    changed = True
    return (serialize_tree(tree), True) if changed else (xml_bytes, False)
