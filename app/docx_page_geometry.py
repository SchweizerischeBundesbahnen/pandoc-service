"""The page size and margins of a DOCX, for its LaTeX/PDF rendering.

pandoc's DOCX reader ignores the page geometry: a PDF made from a DOCX used
LaTeX's own paper and margins, so an image or table that fits the Word page
could run over the PDF one, and one that does not could fit. This module
reads <w:pgSz> and <w:pgMar> of the first section and hands them to LaTeX as
``geometry`` variables. The later sections get theirs at their section break:
see app/docx_page_break_pre_process.py.

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
DOCUMENT_PART = "word/document.xml"
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


def body_sections(body: Element) -> list[Element | None]:
    """The sectPr of each section in order: the ones paragraphs hold, then the body's own, which can be missing."""
    return [*body.findall(f"{{{W_NS}}}p/{{{W_NS}}}pPr/{{{W_NS}}}sectPr"), body.find(f"{{{W_NS}}}sectPr")]


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


def section_page(section: Element | None) -> dict[str, int]:
    """The page size and margins of a section, in twips."""
    size = section.find(f"{{{W_NS}}}pgSz") if section is not None else None
    margins = section.find(f"{{{W_NS}}}pgMar") if section is not None else None
    return {
        "paperwidth": _twips(size, "w", LETTER["w"]),
        "paperheight": _twips(size, "h", LETTER["h"]),
        **{side: _twips(margins, side, fallback) for side, fallback in ONE_INCH_MARGINS.items()},
    }


def _section_text_width_twips(section: Element | None) -> int:
    page = section_page(section)
    width = page["paperwidth"] - page["left"] - page["right"]
    # Margins which leave no text are no page LaTeX can set: the Letter one stands in
    return width if width > 0 else LETTER["w"] - ONE_INCH_MARGINS["left"] - ONE_INCH_MARGINS["right"]


def text_width_twips(document: Element) -> int:
    """The width of the text of the first section, in twips."""
    return _section_text_width_twips(_first_section(document))


def table_text_widths_twips(document: Element) -> dict[Element, int]:
    """The width of the text of the section each table of the body stands in, nested tables included, in twips."""
    body = document.find(f"{{{W_NS}}}body")
    if body is None:
        return {}
    widths: dict[Element, int] = {}
    sections = iter(body_sections(body))
    width = _section_text_width_twips(next(sections))
    for block in body:
        for table in block.iter(f"{{{W_NS}}}tbl"):
            widths[table] = width
        if block.find(f"{{{W_NS}}}pPr/{{{W_NS}}}sectPr") is not None and block.tag == f"{{{W_NS}}}p":
            width = _section_text_width_twips(next(sections))
    return widths


def package_text_width_twips(entries: dict[str, bytes]) -> int:
    """The text width of the first section of a DOCX package read into its parts, in twips."""
    document = parse_xml(entries[DOCUMENT_PART]) if DOCUMENT_PART in entries else None
    return text_width_twips(document) if document is not None else LETTER["w"] - ONE_INCH_MARGINS["left"] - ONE_INCH_MARGINS["right"]


def geometry_variables(docx_bytes: bytes) -> list[str]:
    """Pandoc ``-V geometry:...`` arguments for the page of the first section."""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as package:
            document = parse_xml(package.read(DOCUMENT_PART))
    except _UNREADABLE_DOCX:
        document = None
    values = section_page(_first_section(document) if document is not None else None)
    variables: list[str] = []
    for name, twips in values.items():
        variables += ["-V", f"geometry:{name}={_inches(twips)}"]
    return variables
