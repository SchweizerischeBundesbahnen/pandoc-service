"""The page size and margins of a DOCX, for its LaTeX/PDF rendering.

pandoc's DOCX reader ignores the page geometry: a PDF made from a DOCX used
LaTeX's own paper and margins, so an image or table that fits the Word page
could run over the PDF one, and one that does not could fit. This module
reads <w:pgSz> and <w:pgMar> of the first section and hands them to LaTeX as
``geometry`` variables.

A DOCX without them (pandoc writes none) gets US Letter with 1 inch
margins, the same fallback app/docx_post_process.py sizes images by.
"""

from __future__ import annotations

import io
import zipfile
from typing import TYPE_CHECKING

from .docx_ooxml import W_NS, parse_xml

if TYPE_CHECKING:
    from xml.etree.ElementTree import Element

TWIPS_PER_INCH = 1440
LETTER = {"w": 12240, "h": 15840}
ONE_INCH_MARGINS = {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440}

# What a source this cannot read raises: not a zip, or a zip without a document in it.
# See html_lists_pre_process for why this is a name rather than an inline except-tuple
# (ruff-format / PEP 758 interaction on Python 3.14).
_UNREADABLE_DOCX = (zipfile.BadZipFile, KeyError)


def _first_section(document: Element) -> Element | None:
    """The sectPr of the first section: the first one a paragraph holds, else the body's own."""
    body = document.find(f"{{{W_NS}}}body")
    if body is None:
        return None
    first = body.find(f"{{{W_NS}}}p/{{{W_NS}}}pPr/{{{W_NS}}}sectPr")
    return first if first is not None else body.find(f"{{{W_NS}}}sectPr")


def _twips(element: Element | None, attribute: str, fallback: int) -> int:
    if element is None:
        return fallback
    value = element.get(f"{{{W_NS}}}{attribute}")
    try:
        # A negative top or bottom margin fixes the header distance; its size is what counts.
        return abs(int(value)) if value is not None else fallback
    except ValueError:
        return fallback


def _inches(twips: int) -> str:
    return f"{twips / TWIPS_PER_INCH:.4f}in"


def geometry_variables(docx_bytes: bytes) -> list[str]:
    """Pandoc ``-V geometry:...`` arguments for the page of the first section."""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as package:
            document = parse_xml(package.read("word/document.xml"))
    except _UNREADABLE_DOCX:
        document = None
    section = _first_section(document) if document is not None else None
    size = section.find(f"{{{W_NS}}}pgSz") if section is not None else None
    margins = section.find(f"{{{W_NS}}}pgMar") if section is not None else None

    values = {
        "paperwidth": _twips(size, "w", LETTER["w"]),
        "paperheight": _twips(size, "h", LETTER["h"]),
        **{side: _twips(margins, side, fallback) for side, fallback in ONE_INCH_MARGINS.items()},
    }
    variables: list[str] = []
    for name, twips in values.items():
        variables += ["-V", f"geometry:{name}={_inches(twips)}"]
    return variables
