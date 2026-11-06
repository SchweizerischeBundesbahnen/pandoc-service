"""The page size and margins of a DOCX, for its LaTeX/PDF rendering.

pandoc's DOCX reader ignores the page geometry: a PDF made from a DOCX used
LaTeX's own paper and margins, so an image or table that fits the Word page
could run over the PDF one, and one that does not could fit. This module
reads <w:pgSz> and <w:pgMar> of the first section and hands them to LaTeX as
``geometry`` variables.

A DOCX without them (pandoc writes none) gets US Letter with 1 inch
margins, the same fallback app/docx_post_process.py sizes images by.

Word prints a document with comments on a smaller page, with a markup area
right of it which holds the balloons of the comments. For such a document the
page is laid out on a larger paper, which ``\\mag`` shrinks to the size of the
DOCX page, and filters/docx/docx_comments_to_latex.lua draws the markup area and
the balloons where these variables place them.
"""

from __future__ import annotations

import io
import zipfile
from typing import TYPE_CHECKING

from .docx_ooxml import W_NS, parse_xml

if TYPE_CHECKING:
    from xml.etree.ElementTree import Element

TWIPS_PER_INCH = 1440
TWIPS_PER_POINT = 20
DOCUMENT_PART = "word/document.xml"
LETTER = {"w": 12240, "h": 15840}
ONE_INCH_MARGINS = {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440}

# The markup area of Word for Mac, in points of the page before Word shrinks it:
# the space it leaves left of the page and fills right of it, where the grey area
# starts left of the right edge of the page and how wide it is, where the balloons
# start and how wide they are, and where the line to a balloon turns.
MARKUP_LEFT = 1.26
MARKUP_RIGHT = 206.3
AREA_START = 65.0
AREA_WIDTH = 259.4
BALLOON_START = 41.0
BALLOON_WIDTH = 232.8
LINE_TURN = 67.3
COMMENT_REFERENCE = f"{{{W_NS}}}commentReference"

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


def text_width_twips(document: Element) -> int:
    """The width of the text of the first section, which LaTeX sets the whole PDF at, in twips."""
    section = _first_section(document)
    size = section.find(f"{{{W_NS}}}pgSz") if section is not None else None
    margins = section.find(f"{{{W_NS}}}pgMar") if section is not None else None
    left = _twips(margins, "left", ONE_INCH_MARGINS["left"])
    right = _twips(margins, "right", ONE_INCH_MARGINS["right"])
    width = _twips(size, "w", LETTER["w"]) - left - right
    # Margins which leave no text are no page LaTeX can set: the Letter one stands in
    return width if width > 0 else LETTER["w"] - ONE_INCH_MARGINS["left"] - ONE_INCH_MARGINS["right"]


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
    section = _first_section(document) if document is not None else None
    size = section.find(f"{{{W_NS}}}pgSz") if section is not None else None
    margins = section.find(f"{{{W_NS}}}pgMar") if section is not None else None

    values = {
        "paperwidth": _twips(size, "w", LETTER["w"]),
        "paperheight": _twips(size, "h", LETTER["h"]),
        **{side: _twips(margins, side, fallback) for side, fallback in ONE_INCH_MARGINS.items()},
    }
    if document is not None and next(document.iter(COMMENT_REFERENCE), None) is not None:
        return _markup_variables({name: twips / TWIPS_PER_POINT for name, twips in values.items()})
    variables: list[str] = []
    for name, twips in values.items():
        variables += ["-V", f"geometry:{name}={_inches(twips)}"]
    return variables


def _markup_variables(page: dict[str, float]) -> list[str]:
    """The page with the markup area right of it, shrunk to the paper as Word does; the values in points."""
    width, height = page["paperwidth"], page["paperheight"]
    mag = round(1000 * width / (width + MARKUP_LEFT + MARKUP_RIGHT))
    scale = mag / 1000
    top = (height / scale - height) / 2
    values = {
        "paperwidth": width / scale,
        "paperheight": height / scale,
        "top": top + page["top"],
        "bottom": top + page["bottom"],
        "left": MARKUP_LEFT + page["left"],
        "right": width / scale - MARKUP_LEFT - width + page["right"],
    }
    variables: list[str] = []
    for name, points in values.items():
        variables += ["-V", f"geometry:{name}={points:.2f}bp"]
    edge = MARKUP_LEFT + width
    area = f"{edge - AREA_START:.2f} {top:.2f} {AREA_WIDTH:g} {height:.2f}"
    balloon = f"{edge - BALLOON_START:.2f} {BALLOON_WIDTH:g} {edge - LINE_TURN:.2f}"
    return [*variables, "-M", f"docx-comment-mag={mag}", "-M", f"docx-comment-area={area}", "-M", f"docx-comment-balloon={balloon}"]
