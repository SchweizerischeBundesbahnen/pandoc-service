"""Carry the vertical offset and the side space of inline pictures into the docx→latex path.

Pandoc's DOCX reader drops ``<w:position>`` on the run of a picture, which
raises or lowers it, and ``<wp:effectExtent>`` on its ``<wp:inline>``, which
Word lays out as space beside it. :mod:`app.docx_post_process` writes both for
the CSS ``vertical-align`` and margins of an HTML ``<img>``. The reader keeps
the title of a picture (``<wp:docPr title>``) as the title of the Image, so this
rewrite puts both values in front of that title:

    {{PICLAYOUT:<position in half-points>|<left in EMU>|<right in EMU>}}<original title>

``filters/docx_image_layout_to_latex.lua`` reads the marker, restores the title
and sets the picture in ``\\raisebox`` and ``\\hspace``. The rewrite works on the
copy of the DOCX pandoc reads for LaTeX, never on a DOCX this service returns.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from .docx_ooxml import W_NS, parse_xml, serialize_tree

if TYPE_CHECKING:
    from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)

WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced

MARKER_PREFIX = "{{PICLAYOUT:"
MARKER_END = "}}"

_W_R = f"{{{W_NS}}}r"
_W_POSITION = f"{{{W_NS}}}rPr/{{{W_NS}}}position"
_W_VAL = f"{{{W_NS}}}val"
_WP_INLINE = f"{{{W_NS}}}drawing/{{{WP_NS}}}inline"
_WP_EFFECT_EXTENT = f"{{{WP_NS}}}effectExtent"
_WP_DOC_PR = f"{{{WP_NS}}}docPr"


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Mark the pictures of one body part that are shifted or spaced. Returns (new_bytes, changed)."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping image-layout preprocess")
        return xml_bytes, False

    changed = False
    for run in tree.iter(_W_R):
        inline = run.find(_WP_INLINE)
        doc_pr = inline.find(_WP_DOC_PR) if inline is not None else None
        if inline is None or doc_pr is None:
            continue
        position = _int(run.find(_W_POSITION), _W_VAL)
        effect_extent = inline.find(_WP_EFFECT_EXTENT)
        left = max(_int(effect_extent, "l"), 0)
        right = max(_int(effect_extent, "r"), 0)
        title = doc_pr.get("title", "")
        if not (position or left or right) or title.startswith(MARKER_PREFIX):
            continue
        doc_pr.set("title", f"{MARKER_PREFIX}{position}|{left}|{right}{MARKER_END}{title}")
        changed = True

    if not changed:
        return xml_bytes, False
    return serialize_tree(tree), True


def _int(element: ET.Element | None, attribute: str) -> int:
    """An integer attribute of an element, 0 when the element or a valid value is missing."""
    value = element.get(attribute) if element is not None else None
    try:
        return int(value) if value is not None else 0
    except ValueError:
        return 0
