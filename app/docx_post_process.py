import base64
import io
import logging
import math
import re
import sys
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docx import Document
from docx.oxml import parse_xml
from docx.oxml import parser as docx_parser
from docx.oxml.ns import nsdecls
from lxml import etree  # type: ignore[import-untyped]

if TYPE_CHECKING:
    from collections.abc import Iterator

    from docx.document import Document as DocumentObject
    from docx.section import Section
    from docx.table import Table, _Cell

    from app.html_table_layout import TableLayout

from app import docx_table_columns
from app.docx_math_color_post_process import apply_math_colors
from app.docx_references_post_process import add_table_of_contents_entries, enable_auto_update_fields

# Patch the python-docx parser to handle large XML documents (> 10MB)
# This enables the XML_PARSE_HUGE flag to avoid "Buffer size limit exceeded" errors
# when processing documents with large embedded content (e.g., base64-encoded images)
_huge_tree_parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False, huge_tree=True)
_huge_tree_parser.set_element_class_lookup(docx_parser.element_class_lookup)
docx_parser.oxml_parser = _huge_tree_parser

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

EMU_1_INCH = 914400  # 1 inch in docx in EMU (English Metric Units)
TWIPS_1_INCH = 1440  # 1 inch in docx in Twips (Twentieth of a Point)
DOCX_LETTER_WIDTH_EMU = 8.5 * EMU_1_INCH  # docx LETTER width = 8.5 inch
DOCX_LETTER_SIDE_MARGIN = EMU_1_INCH  # docx left & right margins = 1 inch
DOCX_LETTER_HEIGHT_EMU = 11 * EMU_1_INCH  # docx LETTER height = 11 inch
DOCX_LETTER_TOP_BOTTOM_MARGIN = EMU_1_INCH  # docx top & bottom margins = 1 inch
# An image sits on a line, and the line asks for a little more than the image: the leading above it
# and the depth below. An image given the whole text height therefore does not fit the page it was
# measured against, and the page it opens is the next one, leaving an empty page behind. A sixth of
# an inch is 12 pt, a line of the body text of the documents this produces, and it is 1.5% of the
# height of a page a reader would have to be told about to notice.
LINE_ALLOWANCE_EMU = EMU_1_INCH // 6

WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced

# Paper sizes in TWIPS (portrait orientation: width x height)
PAPER_SIZES = {
    "A5": {"width": 8419, "height": 11906},
    "A4": {"width": 11906, "height": 16838},
    "A3": {"width": 16838, "height": 23811},
    "B5": {"width": 9979, "height": 14144},
    "B4": {"width": 14144, "height": 20013},
    "JIS_B5": {"width": 10319, "height": 14572},
    "JIS_B4": {"width": 14572, "height": 20639},
    "LETTER": {"width": 12240, "height": 15840},
    "LEGAL": {"width": 12240, "height": 20160},
    "LEDGER": {"width": 15840, "height": 24480},
}
logger = logging.getLogger(__name__)

# OOXML tblPr child element order (subset of CT_TblPrBase we touch). Used to
# insert new properties at a schema-valid position so both Word and the
# stricter LibreOffice accept the output regardless of which children pandoc
# (or the inline_styles.lua table rebuild) already emitted.
_TBLPR_CHILD_ORDER = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
]


def process(docx_bytes: bytes, paper_size: str | None = None, orientation: str | None = None, table_layouts: list[TableLayout] | None = None) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    # Read before the paper size replaces the page pandoc sized its images against.
    text_width = _pandoc_text_width_emu(doc)
    _move_header_footer_references_to_first_section(doc)
    _replace_first_paragraph_styles(doc)
    _replace_size_and_orientation(doc, paper_size, orientation)
    # Before the tables, so an image in a cell is brought back to its column.
    _replace_image_placeholders(doc, text_width)
    _state_inline_distances(doc)
    _replace_table_properties(doc, table_layouts)
    _separate_adjacent_tables(doc)
    apply_math_colors(doc)
    _cap_image_heights(doc)
    # After the cap, which can make a picture lower: vertical-align: middle centers on its height.
    _apply_image_layouts(doc)
    _replace_link_placeholders(doc)
    add_table_of_contents_entries(doc)
    enable_auto_update_fields(doc)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# {{IMG:<width>|<height>|<src>}} — see image_placeholder in
# filters/inline_styles.lua. Both dimensions are always present and may be
# empty; "|" cannot occur in a validated dimension or in a data: URI's base64,
# and only the first two fields are delimited by it, so a src containing one is
# still read whole.
_IMG_PLACEHOLDER_RE = re.compile(r"\{\{IMG:([^|]*)\|([^|]*)\|(.*?)\}\}")

# CSS length unit -> EMU. A bare number is px, which is what HTML means by it.
_UNIT_TO_EMU: dict[str, float] = {
    "": EMU_1_INCH / 96,
    "px": EMU_1_INCH / 96,
    "in": float(EMU_1_INCH),
    "cm": EMU_1_INCH / 2.54,
    "mm": EMU_1_INCH / 25.4,
    "pt": EMU_1_INCH / 72,
    "pc": EMU_1_INCH / 6,
}
# A number followed by an optional unit. Written without whitespace
# quantifiers, and with the fraction as one optional group rather than
# `\d+\.?\d*`, so there is nothing for the engine to backtrack over — the
# caller strips the value instead. This mirrors _VALUE_RE in
# app/html_paragraph_pre_process.py, which carries the same note: the previous
# `^\s*...\s*...\s*$` form was already linear-time but matched SonarCloud
# S5852's "multiple \s* quantifiers" heuristic.
_DIMENSION_RE = re.compile(r"^(\d+(?:\.\d+)?)([a-z]*|%)$", re.IGNORECASE)
_HREF_PLACEHOLDER_RE = re.compile(r"\{\{HREF:(.*?)\}\}")

RELATIONSHIPS_SCHEMA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"  # NOSONAR

TWIPS_PER_POINT = 20
EMU_PER_POINT = EMU_1_INCH // 72
# The text width pandoc's DOCX writer uses when the reference document states no page: 420 pt.
PANDOC_DEFAULT_TEXT_WIDTH_EMU = 420 * EMU_PER_POINT


def _pandoc_text_width_emu(doc: DocumentObject) -> int:
    """The text width pandoc sized its images against, in EMU.

    pandoc takes the page width less the side margins from the reference
    document, whose sectPr it copies into the body, in whole points. Without
    all three it uses 420 pt. It reads a percentage as a share of this width,
    also a height, and brings a wider image back to it. Read it before the
    paper size is replaced.
    """
    sect_pr = doc.element.body.find(f"{{{SCHEMA}}}sectPr")
    if sect_pr is None:
        return PANDOC_DEFAULT_TEXT_WIDTH_EMU
    pg_sz = sect_pr.find(f"{{{SCHEMA}}}pgSz")
    pg_mar = sect_pr.find(f"{{{SCHEMA}}}pgMar")
    page_width = _int_attribute(pg_sz, "w")
    left = _int_attribute(pg_mar, "left")
    right = _int_attribute(pg_mar, "right")
    if page_width is None or left is None or right is None or page_width - left - right <= 0:
        return PANDOC_DEFAULT_TEXT_WIDTH_EMU
    return (page_width - left - right) // TWIPS_PER_POINT * EMU_PER_POINT


def _int_attribute(element: Any, name: str) -> int | None:
    """A whole-number w: attribute of an element, or None when either is missing or unreadable."""
    value = element.get(f"{{{SCHEMA}}}{name}") if element is not None else None
    if value is None or not value.lstrip("-").isdigit():
        return None
    return int(value)


def _dimension_to_emu(value: str, text_width: int | None = None) -> int | None:
    """Convert a CSS length from an image placeholder to EMU, or None.

    A percentage is a share of `text_width`, as pandoc's writer reads it. None
    covers the empty field (the node carried no such dimension), any unit this
    cannot turn into an absolute length, and a percentage without a text width.
    """
    match = _DIMENSION_RE.match(value.strip())
    if not match:
        return None
    if match.group(2) == "%":
        return round(float(match.group(1)) * text_width / 100) if text_width is not None else None
    factor = _UNIT_TO_EMU.get(match.group(2).lower())
    if factor is None:
        return None
    return round(float(match.group(1)) * factor)


def _resolve_image_extent(requested: tuple[str, str], px_width: int, px_height: int, text_width: int | None = None) -> tuple[int, int]:
    """The <wp:extent> for an image, in EMU.

    A dimension given on the node wins. When only one is given the other is
    scaled to it, which is what the DOCX writer does, so the aspect ratio is
    kept. With neither, the file's own pixel size at the 96 dpi CSS reference.
    An image wider than `text_width` is brought back to it, as the writer does.
    """
    width, height = _requested_extent(requested, px_width, px_height, text_width)
    if text_width is not None and width > text_width:
        return text_width, round(height * text_width / width)
    return width, height


def _requested_extent(requested: tuple[str, str], px_width: int, px_height: int, text_width: int | None) -> tuple[int, int]:
    """The size the node asks for, in EMU, before any limit.

    When one side alone is a percentage, pandoc's writer ignores a length on
    the other side and scales it by the file's aspect ratio, so this does too.
    """
    native_width = round(px_width * EMU_1_INCH / 96)
    native_height = round(px_height * EMU_1_INCH / 96)
    width = _dimension_to_emu(requested[0], text_width)
    height = _dimension_to_emu(requested[1], text_width)
    width_is_share = requested[0].strip().endswith("%")
    height_is_share = requested[1].strip().endswith("%")
    if width_is_share and not height_is_share and width is not None:
        height = None
    elif height_is_share and not width_is_share and height is not None:
        width = None

    if width is not None and height is not None:
        return width, height
    # Only one side given: scale the other by the file's aspect ratio, guarding
    # a degenerate image rather than dividing by zero.
    if width is not None:
        return width, round(width * native_height / native_width) if native_width else native_height
    if height is not None:
        return round(height * native_width / native_height) if native_height else native_width, height
    return native_width, native_height


def _replace_image_placeholders(doc: DocumentObject, text_width: int | None = None) -> None:
    """Replace ``{{IMG:<src>}}`` placeholders with real embedded images.

    The ``inline_styles.lua`` filter emits these markers when it rebuilds a
    styled table, or a paragraph carrying an indent or alignment, as raw
    OOXML. Images can't be embedded in raw OOXML (they need writer-level
    relationship entries), so the Lua filter writes a text placeholder and
    this function resolves it using python-docx.

    The placeholder carries the width/height the DOCX writer would have
    applied, because that is the only route by which they survive: a raw
    ``<wp:extent>`` has to be written here, and without them an
    ``<img width="100">`` of a 20px picture came out at 20px.

    ``text_width`` is the width pandoc sized its own images against (see
    ``_pandoc_text_width_emu``). A percentage is a share of it, and a wider
    image is brought back to it, so a placeholder image gets the size pandoc
    gives the same image. An image in a table cell is brought back to its
    column afterwards, by ``_replace_table_properties``.
    """
    body = doc.element.body
    doc_pr_id = max((int(dp.get("id", "0")) for dp in body.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr")), default=0)
    for t_el in body.findall(f".//{{{SCHEMA}}}t"):
        if t_el.text is None:
            continue
        match = _IMG_PLACEHOLDER_RE.search(t_el.text)
        if not match:
            continue

        requested = (match.group(1), match.group(2))
        src = match.group(3)
        run_el = t_el.getparent()
        if run_el is None or not run_el.tag.endswith("}r"):
            continue

        image_bytes = _resolve_image_src(src)
        if image_bytes is None:
            # Can't resolve — leave the placeholder text as alt-text fallback
            t_el.text = t_el.text.replace(match.group(0), "[image]")
            continue

        try:
            r_id, img = doc.part.get_or_add_image(io.BytesIO(image_bytes))
            width, height = _resolve_image_extent(requested, img.px_width, img.px_height, text_width)
            doc_pr_id += 1

            # Build the drawing XML
            drawing_xml = (
                f"<w:drawing {nsdecls('w', 'wp', 'a', 'pic', 'r')}>"
                f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
                f'<wp:extent cx="{width}" cy="{height}"/>'
                f'<wp:docPr id="{doc_pr_id}" name="Image"/>'
                f"<a:graphic>"
                f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f"<pic:pic>"
                f"<pic:nvPicPr>"
                f'<pic:cNvPr id="{doc_pr_id}" name="Image"/>'
                f"<pic:cNvPicPr/>"
                f"</pic:nvPicPr>"
                f"<pic:blipFill>"
                f'<a:blip r:embed="{r_id}"/>'
                f"<a:stretch><a:fillRect/></a:stretch>"
                f"</pic:blipFill>"
                f"<pic:spPr>"
                f"<a:xfrm>"
                f'<a:off x="0" y="0"/>'
                f'<a:ext cx="{width}" cy="{height}"/>'
                f"</a:xfrm>"
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                f"</pic:spPr>"
                f"</pic:pic>"
                f"</a:graphicData>"
                f"</a:graphic>"
                f"</wp:inline>"
                f"</w:drawing>"
            )
            drawing_el = parse_xml(drawing_xml)

            # Replace the text run with the drawing
            run_el.remove(t_el)
            run_el.append(drawing_el)

            logger.debug(f"Replaced image placeholder with embedded image ({img.px_width}x{img.px_height})")
        # A placeholder that will not embed leaves the document unchanged.
        except Exception as e:  # noqa: BLE001
            logger.warning(f"Could not embed image from placeholder: {e}")
            t_el.text = t_el.text.replace(match.group(0), "[image]")


def _state_inline_distances(doc: DocumentObject) -> None:
    """Write the distances pandoc leaves out of an inline picture as zero, which is how Word reads them.

    LibreOffice reads a missing distance as about 0.3 cm, so it sets the picture that far in from
    the text. A picture as wide as its cell then runs over the cell's edge and is cut off.
    """
    for inline in doc.element.body.iter(f"{{{WP_NS}}}inline"):
        for side in ("distT", "distB", "distL", "distR"):
            if inline.get(side) is None:
                inline.set(side, "0")


def _resolve_image_src(src: str) -> bytes | None:
    """Resolve an image src to bytes. Supports data: URIs."""
    if src.startswith("data:"):
        # data:image/gif;base64,AAAA...
        match = re.match(r"data:[^;]+;base64,(.*)", src)
        if match:
            try:
                return base64.b64decode(match.group(1))
            # Undecodable image data leaves the document unchanged.
            except Exception:  # noqa: BLE001
                logger.warning("Failed to decode base64 image data")
                return None
    if src:
        logger.warning("Unsupported image src scheme (only data: URIs are supported): %s", src[:80])
    return None


# {{IMGLAYOUT:<vertical-align>|<margin-left>|<margin-right>}} - see
# image_layout_marker in filters/inline_styles.lua. The marker is a run of its
# own, right before the run of the picture it describes.
_IMG_LAYOUT_MARKER_RE = re.compile(r"^\{\{IMGLAYOUT:([^|]*)\|([^|]*)\|([^|]*)\}\}$")
_SIGNED_DIMENSION_RE = re.compile(r"^(-?\d+(?:\.\d+)?)([a-z]*)$", re.IGNORECASE)
EMU_HALF_POINT = EMU_1_INCH // 144
# The share of the font size below the baseline, where vertical-align: bottom
# puts the bottom of a picture. The font's own descent is unknown here, so this
# is an estimate: exact for Calibri, and 0.5pt too deep at 12pt for Arial.
DESCENT_RATIO = 0.25
# The share of the font size a lowercase x takes, half of which is where
# vertical-align: middle puts the middle of a picture. An estimate as well:
# Calibri's is 0.47, Arial's 0.52.
X_HEIGHT_RATIO = 0.5
_FONT_RELATIVE_VERTICAL_ALIGN = frozenset({"bottom", "text-bottom", "middle"})
# The size Word gives text that no style sizes: 10pt.
DEFAULT_FONT_HALF_POINTS = 20
# The <w:rPr> children that come after <w:position> in the schema (CT_RPr).
_RPR_AFTER_POSITION = frozenset({"sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"})


def _signed_dimension_to_emu(value: str) -> int | None:
    """Convert a CSS length that may be negative to EMU, or None."""
    match = _SIGNED_DIMENSION_RE.match(value.strip())
    if not match:
        return None
    factor = _UNIT_TO_EMU.get(match.group(2).lower())
    if factor is None:
        return None
    emu = float(match.group(1)) * factor
    # A long enough run of digits is an infinite float, which round() refuses.
    return round(emu) if math.isfinite(emu) else None


def _apply_image_layouts(doc: DocumentObject) -> None:
    """Apply ``{{IMGLAYOUT:}}`` markers to the pictures that follow them.

    The DOCX writer drops CSS vertical-align and margins of an <img>, and Word
    puts the bottom of an inline picture on the baseline. vertical-align
    becomes <w:position> on the picture's run, which lowers or raises it. The
    margins become <wp:effectExtent>, the only extra space beside an inline
    picture that Word lays out: it ignores distL/distR and w:spacing there.
    """
    # Collected before any is removed: removing a run while the tree is walked skips the next one.
    markers = [(t_el, match) for t_el in doc.element.body.iter(f"{{{SCHEMA}}}t") if (match := _IMG_LAYOUT_MARKER_RE.match(t_el.text or ""))]
    if not markers:
        return
    styles = {style.get(f"{{{SCHEMA}}}styleId"): style for style in doc.styles.element.findall(f"{{{SCHEMA}}}style")}
    # Every marker goes before any picture is placed: a marker left in a paragraph would read as the label of the icon before it.
    pictures = [(picture, match) for t_el, match in markers if (picture := _take_marked_picture(t_el))]
    for (picture_run, inline), match in pictures:
        valign, margin_left, margin_right = match.groups()
        position = _picture_position(doc, styles, picture_run, inline, valign)
        if position:
            _set_run_position(picture_run, position)

        left = max(_signed_dimension_to_emu(margin_left) or 0, 0)
        right = max(_signed_dimension_to_emu(margin_right) or 0, 0)
        if left or right:
            _widen_effect_extent(inline, left, right)


def _take_marked_picture(t_el: Any) -> tuple[Any, Any] | None:
    """Remove a marker's run and return the run and <wp:inline> of the picture after it, or None."""
    marker_run = t_el.getparent()
    if marker_run is None or marker_run.tag != f"{{{SCHEMA}}}r":
        return None
    picture_run = marker_run.getnext()
    marker_run.getparent().remove(marker_run)
    if picture_run is None or picture_run.tag != f"{{{SCHEMA}}}r":
        return None
    inline = picture_run.find(f"{{{SCHEMA}}}drawing/{{{WP_NS}}}inline")
    return None if inline is None else (picture_run, inline)


def _picture_position(doc: DocumentObject, styles: dict[str, Any], picture_run: Any, inline: Any, valign: str) -> int:
    """The <w:position> for a picture's vertical-align, in half-points; 0 for none."""
    if valign not in _FONT_RELATIVE_VERTICAL_ALIGN:
        shift = _signed_dimension_to_emu(valign)
        return round(shift / EMU_HALF_POINT) if shift is not None else 0
    text_size = _text_size_half_points(doc, styles, picture_run)
    if valign == "middle":
        picture_height = int(inline.find(f"{{{WP_NS}}}extent").get("cy", "0")) / EMU_HALF_POINT
        return round(text_size * X_HEIGHT_RATIO / 2 - picture_height / 2)
    return -round(text_size * DESCENT_RATIO)


def _text_size_half_points(doc: DocumentObject, styles: dict[str, Any], picture_run: Any) -> int:
    """The font size of the text around a picture, in half-points.

    Read from the nearest run with text in the same paragraph, the one after
    the picture first, since a label follows its icon.
    """
    paragraph = next(picture_run.iterancestors(f"{{{SCHEMA}}}p"), None)
    if paragraph is None:
        return _default_size_half_points(doc)
    # A run of spaces says nothing of the label's size, so it is passed over.
    runs = [run for run in paragraph.iter(f"{{{SCHEMA}}}r") if run is picture_run or "".join(t.text or "" for t in run.iter(f"{{{SCHEMA}}}t")).strip()]
    index = runs.index(picture_run)
    nearest = (runs[index + 1 :] + runs[:index][::-1])[:1]
    for text_run in nearest:
        size = _half_points(text_run.find(f"{{{SCHEMA}}}rPr/{{{SCHEMA}}}sz")) or _style_size_half_points(styles, _val(text_run.find(f"{{{SCHEMA}}}rPr/{{{SCHEMA}}}rStyle")))
        if size:
            return size

    paragraph_style = _val(paragraph.find(f"{{{SCHEMA}}}pPr/{{{SCHEMA}}}pStyle"))
    if paragraph_style is None:
        paragraph_style = next((style_id for style_id, style in styles.items() if style.get(f"{{{SCHEMA}}}type") == "paragraph" and style.get(f"{{{SCHEMA}}}default") in ("1", "true")), None)
    return _style_size_half_points(styles, paragraph_style) or _default_size_half_points(doc)


def _default_size_half_points(doc: DocumentObject) -> int:
    return _half_points(doc.styles.element.find(f"{{{SCHEMA}}}docDefaults/{{{SCHEMA}}}rPrDefault/{{{SCHEMA}}}rPr/{{{SCHEMA}}}sz")) or DEFAULT_FONT_HALF_POINTS


def _style_size_half_points(styles: dict[str, Any], style_id: str | None) -> int | None:
    """The font size a style states or inherits through basedOn, in half-points."""
    seen: set[str] = set()
    while style_id is not None and style_id not in seen and style_id in styles:
        seen.add(style_id)
        style = styles[style_id]
        size = _half_points(style.find(f"{{{SCHEMA}}}rPr/{{{SCHEMA}}}sz"))
        if size:
            return size
        style_id = _val(style.find(f"{{{SCHEMA}}}basedOn"))
    return None


def _val(element: Any) -> str | None:
    return element.get(f"{{{SCHEMA}}}val") if element is not None else None


def _half_points(sz: Any) -> int | None:
    value = _val(sz)
    return int(value) if value is not None and value.isdigit() and int(value) > 0 else None


def _set_run_position(run: Any, half_points: int) -> None:
    """Give a run its <w:position>, at the place in <w:rPr> the schema wants."""
    r_pr = run.find(f"{{{SCHEMA}}}rPr")
    if r_pr is None:
        r_pr = run.makeelement(f"{{{SCHEMA}}}rPr", {})
        run.insert(0, r_pr)
    for old in r_pr.findall(f"{{{SCHEMA}}}position"):
        r_pr.remove(old)
    position = r_pr.makeelement(f"{{{SCHEMA}}}position", {f"{{{SCHEMA}}}val": str(half_points)})
    follower = next((child for child in r_pr if etree.QName(child).localname in _RPR_AFTER_POSITION), None)
    if follower is not None:
        follower.addprevious(position)
    else:
        r_pr.append(position)


def _widen_effect_extent(inline: Any, left: int, right: int) -> None:
    """Add space left and right of an inline picture, in EMU."""
    effect_extent = inline.find(f"{{{WP_NS}}}effectExtent")
    if effect_extent is None:
        effect_extent = inline.makeelement(f"{{{WP_NS}}}effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
        # The schema wants it right after <wp:extent>, which an inline always has.
        inline.find(f"{{{WP_NS}}}extent").addnext(effect_extent)
    effect_extent.set("l", str(int(effect_extent.get("l", "0")) + left))
    effect_extent.set("r", str(int(effect_extent.get("r", "0")) + right))


def _replace_link_placeholders(doc: DocumentObject) -> None:
    """Replace ``{{HREF:<url>}}`` placeholders in hyperlink tooltips with real relationships.

    The ``inline_styles.lua`` filter emits ``<w:hyperlink w:tooltip="{{HREF:url}}">``
    when it rebuilds styled tables as raw OOXML. This function registers the
    URL as a hyperlink relationship and sets the correct ``r:id``.
    """
    ns_w = f"{{{SCHEMA}}}"
    ns_r = f"{{{RELATIONSHIPS_SCHEMA}}}"
    body = doc.element.body

    for hyperlink in body.findall(f".//{ns_w}hyperlink"):
        tooltip = hyperlink.get(f"{ns_w}tooltip", "")
        match = _HREF_PLACEHOLDER_RE.search(tooltip)
        if not match:
            continue

        url = match.group(1)
        try:
            r_id = doc.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)  # NOSONAR
            hyperlink.set(f"{ns_r}id", r_id)
            hyperlink.attrib.pop(f"{ns_w}tooltip", None)
            logger.debug(f"Resolved hyperlink placeholder to {url} (r:id={r_id})")
        # An unresolvable hyperlink leaves the document unchanged.
        except Exception as e:  # noqa: BLE001
            logger.warning(f"Could not resolve hyperlink placeholder: {e}")
            hyperlink.attrib.pop(f"{ns_w}tooltip", None)


def _replace_first_paragraph_styles(doc: DocumentObject) -> None:
    """Replace pandoc's "First Paragraph" style with "Body Text".

    Pandoc's DOCX writer automatically assigns "First Paragraph" to the first
    paragraph after every heading.  This creates visual inconsistency because
    only that paragraph differs in style from subsequent ones.  Normalizing all
    such paragraphs to "Body Text" gives a uniform look.
    """
    for paragraph in doc.element.body.iter(f"{{{SCHEMA}}}p"):
        p_pr = paragraph.find(f"{{{SCHEMA}}}pPr")
        if p_pr is None:
            continue
        p_style = p_pr.find(f"{{{SCHEMA}}}pStyle")
        if p_style is not None and p_style.get(f"{{{SCHEMA}}}val") == "FirstParagraph":
            p_style.set(f"{{{SCHEMA}}}val", "BodyText")


# Elements Word skips over when it decides whether two tables touch. Anything
# else between two tables keeps them apart, so the scan below stops at it.
_TABLE_RANGE_MARKERS = frozenset(
    {
        "bookmarkStart",
        "bookmarkEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "permStart",
        "permEnd",
        "proofErr",
    }
)


def _is_table_range_marker(element: Any) -> bool:
    """Return True if Word ignores this element when pairing up adjacent tables."""
    if not isinstance(element.tag, str):
        return True  # XML comment or processing instruction
    prefix = f"{{{SCHEMA}}}"
    return element.tag.startswith(prefix) and element.tag[len(prefix) :] in _TABLE_RANGE_MARKERS


def _separate_adjacent_tables(doc: DocumentObject) -> None:
    """Insert an empty paragraph between tables Word would otherwise merge.

    Two ``w:tbl`` siblings with no ``w:p`` between them render as a single
    table. Pandoc's DOCX writer inserts the separating paragraph itself, but
    only for ``Table`` blocks that are immediate neighbours in one block list.
    Tables that a wrapping ``<div>`` puts at different AST depths - Polarion
    emits exactly that around work item fields - never look adjacent to it and
    reach the output back to back, joined by nothing but bookmarks.
    """
    tbl_tag = f"{{{SCHEMA}}}tbl"
    # findall collects the tables up front, which the insertion below needs:
    # a lazy walk would run over a tree that grows under it.
    for table_element in doc.element.body.findall(f".//{tbl_tag}"):
        sibling = table_element.getnext()
        while sibling is not None and _is_table_range_marker(sibling):
            sibling = sibling.getnext()
        if sibling is not None and sibling.tag == tbl_tag:
            table_element.addnext(parse_xml(f"<w:p {nsdecls('w')}/>"))


def _replace_size_and_orientation(doc: DocumentObject, paper_size: str | None = None, orientation: str | None = None) -> None:
    # If both parameters are None, no modifications needed
    if paper_size is None and orientation is None:
        return

    for section in doc.sections:
        # Python-docx exposes no public API for this element.
        sect_pr = section._sectPr  # noqa: SLF001
        pg_sz = sect_pr.find(".//w:pgSz", namespaces={"w": SCHEMA})

        if paper_size is not None:
            pg_sz = _set_paper_size(sect_pr, pg_sz, paper_size, orientation)

        if orientation is not None:
            _set_orientation(sect_pr, pg_sz, orientation)


def _set_paper_size(sect_pr: Any, pg_sz: Any, paper_size: str, orientation: str | None) -> Any:
    """Set the paper size for a section."""
    # Normalize paper_size to uppercase for case-insensitive lookup
    paper_size_upper = paper_size.upper()

    # Get dimensions for the specified paper size
    if paper_size_upper not in PAPER_SIZES:
        raise ValueError(f"Unsupported paper size: {paper_size}. Supported sizes: {', '.join(PAPER_SIZES.keys())}")

    page_dims = PAPER_SIZES[paper_size_upper]
    width = page_dims["width"]
    height = page_dims["height"]

    # Get existing orientation if present (to preserve it when changing paper size)
    existing_orientation = pg_sz.get(f"{{{SCHEMA}}}orient") if pg_sz is not None else None

    # Apply existing orientation if orientation parameter is not specified
    if orientation is None and existing_orientation == "landscape":
        width, height = height, width

    # Create or update pg_sz element
    if pg_sz is None:
        pg_sz = parse_xml(f'<w:pgSz {nsdecls("w")} w:w="{width}" w:h="{height}"/>')
        sect_pr.append(pg_sz)
    else:
        pg_sz.set(f"{{{SCHEMA}}}w", str(width))
        pg_sz.set(f"{{{SCHEMA}}}h", str(height))

    # Preserve existing orientation attribute if present and orientation parameter not specified
    if orientation is None and existing_orientation is not None:
        pg_sz.set(f"{{{SCHEMA}}}orient", existing_orientation)

    return pg_sz


def _set_orientation(sect_pr: Any, pg_sz: Any, orientation: str) -> None:
    """Set the orientation for a section."""
    # Ensure pg_sz exists (use LETTER as default if missing)
    if pg_sz is None:
        page_dims = PAPER_SIZES["LETTER"]
        width = page_dims["width"]
        height = page_dims["height"]
        pg_sz = parse_xml(f'<w:pgSz {nsdecls("w")} w:w="{width}" w:h="{height}"/>')
        sect_pr.append(pg_sz)

    # Get current dimensions
    current_width = int(pg_sz.get(f"{{{SCHEMA}}}w", "0"))
    current_height = int(pg_sz.get(f"{{{SCHEMA}}}h", "0"))

    # Determine current and desired orientation
    current_is_landscape = current_width > current_height
    desired_is_landscape = orientation.lower() == "landscape"

    # Swap dimensions if orientations don't match
    if current_is_landscape != desired_is_landscape:
        pg_sz.set(f"{{{SCHEMA}}}w", str(current_height))
        pg_sz.set(f"{{{SCHEMA}}}h", str(current_width))

    # Set or remove the orient attribute
    if desired_is_landscape:
        pg_sz.set(f"{{{SCHEMA}}}orient", "landscape")
    # Remove orient attribute for portrait (it's the default)
    elif f"{{{SCHEMA}}}orient" in pg_sz.attrib:
        del pg_sz.attrib[f"{{{SCHEMA}}}orient"]


def _move_header_footer_references_to_first_section(doc: DocumentObject) -> None:
    """
    Move header/footer references from the last section to the first section.

    When Lua filters insert section breaks (e.g., for page orientation changes),
    they create new sections without header/footer references. The template's
    header/footer refs end up only in the last section.

    In OOXML, sections without explicit header/footer refs inherit from previous
    sections. By moving the refs to the first section, all subsequent sections
    will inherit them automatically. This also properly handles first/odd/even
    page header/footer configurations.

    Additionally, the <w:titlePg/> element is moved along with the references.
    This element controls the "different first page" header/footer feature in Word.
    If left in the last section while refs are in the first, Word would incorrectly
    apply "first page" headers/footers to the wrong pages.
    """
    if len(doc.sections) <= 1:
        return  # Nothing to fix if there's only one section

    # Python-docx exposes no public API for this element.
    first_sect_pr = doc.sections[0]._sectPr  # noqa: SLF001
    # Python-docx exposes no public API for this element.
    last_sect_pr = doc.sections[-1]._sectPr  # noqa: SLF001

    # Check if first section already has header/footer references
    existing_headers = first_sect_pr.findall("w:headerReference", namespaces={"w": SCHEMA})
    existing_footers = first_sect_pr.findall("w:footerReference", namespaces={"w": SCHEMA})

    if existing_headers or existing_footers:
        return  # First section already has refs, nothing to do

    # Find and move header references from last to first section
    header_refs = last_sect_pr.findall("w:headerReference", namespaces={"w": SCHEMA})
    for ref in header_refs:
        last_sect_pr.remove(ref)
        first_sect_pr.insert(0, ref)

    # Find and move footer references from last to first section
    footer_refs = last_sect_pr.findall("w:footerReference", namespaces={"w": SCHEMA})
    for ref in footer_refs:
        last_sect_pr.remove(ref)
        first_sect_pr.insert(0, ref)

    # Move titlePg element if present (controls "different first page" header/footer)
    title_pg = last_sect_pr.find("w:titlePg", namespaces={"w": SCHEMA})
    if title_pg is not None:
        last_sect_pr.remove(title_pg)
        first_sect_pr.append(title_pg)


def _replace_table_properties(doc: DocumentObject, table_layouts: list[TableLayout] | None = None) -> None:  # NOSONAR  # needed by design
    # Per-table width/alignment recovered from the HTML source (see
    # app/html_table_layout.py). The list is one entry per <table> in document
    # order (depth-first, nested included) — the same order this function walks
    # tables in — so a shared iterator lines them up index-for-index. Guard on
    # an exact count match: if pandoc dropped or added a table the alignment
    # would be off, so we skip applying layouts entirely and fall back to the
    # previous 100 %/autofit default rather than mislabel tables.
    layout_iter: Iterator[TableLayout] | None = None
    if table_layouts:
        table_count = len(doc.element.body.findall(".//w:tbl", namespaces={"w": SCHEMA}))
        if table_count == len(table_layouts):
            layout_iter = iter(table_layouts)
        else:
            logger.warning("html_table_layout: %d layouts for %d tables; skipping width/alignment (fallback to defaults)", len(table_layouts), table_count)

    max_widths = [_get_available_content_width_for_section(section) for section in doc.sections]
    # Python-docx exposes no public API for this element.
    tables = {table._element: table for table in doc.tables}  # noqa: SLF001
    for element, section_index in _body_by_section(doc):
        table = tables.get(element)
        if table is not None:
            _process_table(table, 0, max_widths[min(section_index, len(max_widths) - 1)], layout_iter)


def _body_by_section(doc: DocumentObject) -> Iterator[tuple[Any, int]]:
    """Each child of the body, with the index of the section it belongs to.

    A paragraph holding a sectPr closes its section. The body's own sectPr closes the last one.
    """
    section_index = 0
    for element in doc.element.body:
        yield element, section_index
        if element.tag == f"{{{SCHEMA}}}p" and element.find(f"{{{SCHEMA}}}pPr/{{{SCHEMA}}}sectPr") is not None:
            section_index += 1


def _process_table(table: Table, parent_columns_count: int, max_width: int, layout_iter: Iterator[TableLayout] | None = None) -> None:
    # Python-docx exposes no public API for this element.
    tbl = table._element  # noqa: SLF001
    # Pull this table's layout first, before recursing into nested tables, so
    # consumption order stays depth-first and matches html_table_layout.extract.
    layout = next(layout_iter, None) if layout_iter is not None else None
    columns_count = parent_columns_count + len(table.columns)
    table_properties = tbl.find(".//w:tblPr", namespaces={"w": SCHEMA})
    if table_properties is None:
        table_properties = parse_xml(f"<w:tblPr {nsdecls('w')}/>")
        tbl.insert(0, table_properties)

    grid_is_stated = _grid_is_stated(table_properties)
    _apply_table_layout(tbl, table_properties, layout, max_width)
    styles = _table_styles(table)
    _lay_out_columns(tbl, table_properties, layout, max_width, styles, grid_is_stated=grid_is_stated)

    # Read after the layout, which decided the grid.
    column_widths = _column_widths_emu(tbl)

    # Process nested tables
    for row in table.rows:
        for cell in row.cells:
            cell_width = _cell_image_width(tbl, cell, column_widths, max_width / columns_count, max_width, styles)
            _resize_images_in_cell(cell, cell_width)
            # A nested table sits inside this cell, so nothing in it may be wider than the cell is
            for sub_table in cell.tables:
                _process_table(sub_table, columns_count, int(cell_width), layout_iter)


def _grid_is_stated(table_properties: Any) -> bool:
    """Whether the grid holds the column widths the HTML states, read before the layout replaces the table width.

    pandoc keeps the column widths of a ``<colgroup>`` in percent and then states the table width.
    Without them it splits its text width evenly and leaves the table width automatic.
    """
    tbl_w = table_properties.find("w:tblW", namespaces={"w": SCHEMA})
    return tbl_w is not None and tbl_w.get(f"{{{SCHEMA}}}type") not in (None, "auto")


def _lay_out_columns(tbl: Any, table_properties: Any, layout: TableLayout | None, max_width: int, styles: Any, *, grid_is_stated: bool) -> None:
    """Decide the width of each column and write it into the grid and into each cell.

    Images in a cell are fitted to the grid afterwards, and Word lays out the cells by their
    widths, so the two agree. A grid holding the widths the HTML states is kept, scaled to the
    table. Otherwise the widths the HTML states for the cells decide, and the content of each
    column where it states none: see :mod:`app.docx_table_columns`.
    """
    table_width = _table_width_emu(table_properties, max_width)
    stated = layout.column_widths if layout is not None else None
    grid = _column_widths_emu(tbl)
    if stated is None and grid_is_stated and grid is not None:
        widths: list[int] | None = [width * table_width // sum(grid) for width in grid]
    else:
        widths = docx_table_columns.decide(tbl, table_width, stated, lambda tc: _cell_side_margins_emu(tbl, tc, styles), styles)
    if widths:
        _write_column_widths(tbl, widths)


def _table_width_emu(table_properties: Any, max_width: int) -> int:
    """The width the table is laid out at, in EMU: a share of `max_width`, or its own absolute width."""
    tbl_w = table_properties.find("w:tblW", namespaces={"w": SCHEMA})
    width_type = tbl_w.get(f"{{{SCHEMA}}}type") if tbl_w is not None else None
    value = (tbl_w.get(f"{{{SCHEMA}}}w") or "").strip() if tbl_w is not None else ""
    if width_type == "pct":
        return max_width * min(_pct_fiftieths(value), docx_table_columns.FULL_PCT) // docx_table_columns.FULL_PCT
    if width_type == "dxa" and value.isdigit() and int(value) > 0:
        width = int(value) * TWIPS_TO_EMU
        return min(width, max_width) if max_width > 0 else width
    return max_width


def _pct_fiftieths(value: str) -> int:
    """A percentage table width in fiftieths of a percent. Strict OOXML writes it as "100%"."""
    try:
        return round(float(value[:-1]) * 50) if value.endswith("%") else int(value)
    except ValueError:
        return docx_table_columns.FULL_PCT


def _write_column_widths(tbl: Any, widths: list[int]) -> None:
    """Write the column widths into the table grid and into the preferred width of each cell."""
    twips = [max(1, round(width / TWIPS_TO_EMU)) for width in widths]
    grid = tbl.find("w:tblGrid", namespaces={"w": SCHEMA})
    if grid is None:
        grid = parse_xml(f"<w:tblGrid {nsdecls('w')}/>")
        tbl.find("w:tblPr", namespaces={"w": SCHEMA}).addnext(grid)
    for column in grid.findall("w:gridCol", namespaces={"w": SCHEMA}):
        grid.remove(column)
    for width in twips:
        grid.append(parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{width}"/>'))
    for offset, span, tc in docx_table_columns.cells(tbl, len(twips)):
        _set_cell_width(tc, sum(twips[offset : offset + span]))


def _set_cell_width(tc: Any, width_twips: int) -> None:
    """Set a cell's preferred width, where the schema puts it: first in its properties, after a cnfStyle."""
    tc_pr = tc.find("w:tcPr", namespaces={"w": SCHEMA})
    if tc_pr is None:
        tc_pr = parse_xml(f"<w:tcPr {nsdecls('w')}/>")
        tc.insert(0, tc_pr)
    for tc_w in tc_pr.findall("w:tcW", namespaces={"w": SCHEMA}):
        tc_pr.remove(tc_w)
    tc_w = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{width_twips}" w:type="dxa"/>')
    cnf_style = tc_pr.find("w:cnfStyle", namespaces={"w": SCHEMA})
    if cnf_style is not None:
        cnf_style.addnext(tc_w)
    else:
        tc_pr.insert(0, tc_w)


# Word's "Normal Table" style: 0.075 inch left and right of a cell's content.
DEFAULT_CELL_SIDE_MARGIN_TWIPS = 108
TWIPS_TO_EMU = 635


def _column_widths_emu(tbl: Any) -> list[int] | None:
    """The width of each grid column, in EMU; None when the grid states no usable width."""
    grid = tbl.find("w:tblGrid", namespaces={"w": SCHEMA})
    if grid is None:
        return None
    widths = [int(col.get(f"{{{SCHEMA}}}w") or 0) for col in grid.findall("w:gridCol", namespaces={"w": SCHEMA})]
    if not widths or min(widths) <= 0:
        return None
    return [width * TWIPS_TO_EMU for width in widths]


def _table_styles(table: Table) -> Any:
    """The styles of the document the table belongs to, or None where they cannot be read."""
    try:
        return table.part.document.styles.element  # type: ignore[attr-defined]
    # A table built outside a document part has no styles to read, and the default stands.
    except Exception:  # noqa: BLE001
        return None


def _style_cell_margins(styles: Any, tbl: Any) -> list[Any]:
    """The `tblCellMar` of the style the table names and of every style it is based on, farthest first.

    A template sets the margins of every table through its style rather than on each table, and a
    cell laid out inside margins this did not read is a cell an image is measured too wide for. A
    style states the sides it changes and leaves the rest to the style it is based on, so the whole
    chain is returned and each side is taken from the nearest style which states it.
    """
    if styles is None:
        return []
    reference = tbl.find("w:tblPr/w:tblStyle", namespaces={"w": SCHEMA})
    name = reference.get(f"{{{SCHEMA}}}val") if reference is not None else None
    chain: list[Any] = []
    seen: set[str] = set()
    while name is not None and name not in seen:
        seen.add(name)
        style = next((element for element in styles.findall("w:style", namespaces={"w": SCHEMA}) if element.get(f"{{{SCHEMA}}}styleId") == name), None)
        if style is None:
            break
        margins = style.find("w:tblPr/w:tblCellMar", namespaces={"w": SCHEMA})
        if margins is not None:
            chain.append(margins)
        based_on = style.find("w:basedOn", namespaces={"w": SCHEMA})
        name = based_on.get(f"{{{SCHEMA}}}val") if based_on is not None else None
    # Nearest last, so a style overrides the one it is based on under the "last one wins" reading
    chain.reverse()
    return chain


def _cell_side_margins_emu(tbl: Any, tc: Any, styles: Any = None) -> int:
    """Left plus right margin of a cell: its own, else the table's, else its styles', else Word's default."""
    total = 0
    scopes = [*_style_cell_margins(styles, tbl), tbl.find("w:tblPr/w:tblCellMar", namespaces={"w": SCHEMA}), tc.find("w:tcPr/w:tcMar", namespaces={"w": SCHEMA})]
    for side, alternative in (("left", "start"), ("right", "end")):
        width = DEFAULT_CELL_SIDE_MARGIN_TWIPS
        # Last one wins, so they are read from the weakest scope to the strongest
        for scope in scopes:
            if scope is None:
                continue
            margin = scope.find(f"w:{side}", namespaces={"w": SCHEMA})
            if margin is None:
                margin = scope.find(f"w:{alternative}", namespaces={"w": SCHEMA})
            if margin is not None and margin.get(f"{{{SCHEMA}}}type", "dxa") == "dxa":
                width = int(margin.get(f"{{{SCHEMA}}}w") or 0)
        total += width
    return total * TWIPS_TO_EMU


def _cell_image_width(tbl: Any, cell: _Cell, column_widths: list[int] | None, fallback: float, limit: float, styles: Any = None) -> float:
    """The widest an image in the cell may be: its columns less the cell margins.

    A cell spanning columns takes the sum of them. Without a usable grid the even share of the page
    stays the limit. `limit` is the space the table itself has, the page for a table of its own and
    the cell for a nested one: a grid may state more than that, which `_apply_table_layout` leaves
    as it is for a table laid out to its content, and an image is no wider than the space there is.
    """
    if column_widths is None:
        return min(fallback, limit)
    # Python-docx exposes no public API for this element.
    tc = cell._tc  # noqa: SLF001
    columns = column_widths[tc.grid_offset : tc.grid_offset + tc.grid_span]
    if not columns:
        return min(fallback, limit)
    return max(min(sum(columns) - _cell_side_margins_emu(tbl, tc, styles), limit), 1)


def _clamp_twips(width_twips: int, max_width_emu: int) -> int:
    """Clamp a width in twips to the available page width (given in EMU)."""
    if max_width_emu <= 0:
        return width_twips
    max_twips = max_width_emu // 635
    if width_twips > max_twips:
        logger.debug(f"Clamped table width from {width_twips} to {max_twips} twips")
        return max_twips
    return width_twips


def _clamp_existing_fixed_width(tbl: Any, table_properties: Any, max_width: int) -> None:
    """Clamp a Lua-filter-set dxa width to the page width if it overflows."""
    if max_width <= 0:
        return
    tbl_w = table_properties.find("w:tblW", namespaces={"w": SCHEMA})
    current = int(tbl_w.get(f"{{{SCHEMA}}}w", "0"))
    clamped = _clamp_twips(current, max_width)
    if clamped < current:
        tbl_w.set(f"{{{SCHEMA}}}w", str(clamped))
        _rescale_table_grid(tbl, clamped)


def _resolve_layout_width(layout: TableLayout | None) -> tuple[str, int, bool]:
    """Extract width parameters from an html_table_layout, or return defaults."""
    if layout is not None and layout.width_type is not None and layout.width_value is not None:
        return layout.width_type, layout.width_value, layout.width_type == "dxa"
    return "pct", 5000, False


def _apply_table_layout(tbl: Any, table_properties: Any, layout: TableLayout | None, max_width: int = 0) -> None:
    """Write width, alignment and indent onto a table's <w:tblPr>."""
    has_layout = layout is not None and layout.width_type is not None and layout.width_value is not None

    if not has_layout and _has_existing_fixed_width(table_properties):
        _clamp_existing_fixed_width(tbl, table_properties, max_width)
        return

    width_type, width_value, use_fixed_layout = _resolve_layout_width(layout)

    if use_fixed_layout:
        width_value = _clamp_twips(width_value, max_width)
        _rescale_table_grid(tbl, width_value)

    _set_tblpr_child(table_properties, parse_xml(f'<w:tblW {nsdecls("w")} w:w="{width_value}" w:type="{width_type}"/>'))

    layout_type = "fixed" if use_fixed_layout else "autofit"
    _set_tblpr_child(table_properties, parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="{layout_type}"/>'))

    if layout is not None and layout.jc is not None:
        _set_tblpr_child(table_properties, parse_xml(f'<w:jc {nsdecls("w")} w:val="{layout.jc}"/>'))

    if layout is not None and layout.indent_twips is not None:
        _set_tblpr_child(table_properties, parse_xml(f'<w:tblInd {nsdecls("w")} w:w="{layout.indent_twips}" w:type="dxa"/>'))


def _has_existing_fixed_width(table_properties: Any) -> bool:
    """Return True if the table already has a fixed (dxa) width from the Lua filter."""
    tbl_w = table_properties.find("w:tblW", namespaces={"w": SCHEMA})
    if tbl_w is not None and tbl_w.get(f"{{{SCHEMA}}}type") == "dxa":
        w = tbl_w.get(f"{{{SCHEMA}}}w", "0")
        return int(w) > 0
    return False


def _set_tblpr_child(table_properties: Any, new_child: Any) -> None:
    """Replace any existing child with the same tag and insert at a schema-valid position."""
    local_name = etree.QName(new_child).localname
    for existing in table_properties.findall(f"w:{local_name}", namespaces={"w": SCHEMA}):
        table_properties.remove(existing)

    order = _TBLPR_CHILD_ORDER.index(local_name)
    for child in table_properties:
        child_local = etree.QName(child).localname
        if child_local in _TBLPR_CHILD_ORDER and _TBLPR_CHILD_ORDER.index(child_local) > order:
            child.addprevious(new_child)
            return
    table_properties.append(new_child)


def _rescale_table_grid(tbl: Any, target_twips: int) -> None:
    """Scale the <w:tblGrid> column widths so they sum to target_twips.

    Under fixed layout Word derives the rendered table width from the grid
    column widths (not <w:tblW>), so an absolute table width only takes effect
    once the grid is rescaled to match it. Proportional scaling preserves the
    relative column sizing pandoc emitted; a zero/absent grid is distributed
    evenly.
    """
    grid = tbl.find("w:tblGrid", namespaces={"w": SCHEMA})
    if grid is None:
        return
    columns = grid.findall("w:gridCol", namespaces={"w": SCHEMA})
    if not columns:
        return

    width_attr = f"{{{SCHEMA}}}w"
    widths = [int(col.get(width_attr) or 0) for col in columns]
    total = sum(widths)

    if total <= 0:
        even = max(1, target_twips // len(columns))
        for col in columns:
            col.set(width_attr, str(even))
        return

    for col, width in zip(columns, widths, strict=False):
        col.set(width_attr, str(max(1, round(width * target_twips / total))))


def _margin_size(stated: int | None, fallback: int) -> int:
    """The size of a margin the document states, or the fallback where it states none.

    A margin of nothing is one the document states, so it is not replaced: a page laid out edge to
    edge keeps the whole of its width and height. A negative top or bottom margin fixes the distance
    of the header, and its size is what the text is laid out inside, which is how
    `app/docx_page_geometry.py` reads it for the PDF.
    """
    return fallback if stated is None else abs(int(stated))


def _get_available_content_width_for_section(section: Section) -> int:
    # Provide alternative 'Letter' paper size params in case if they were not set explicitly in the document
    page_width = section.page_width or DOCX_LETTER_WIDTH_EMU
    left_margin = _margin_size(section.left_margin, DOCX_LETTER_SIDE_MARGIN)
    right_margin = _margin_size(section.right_margin, DOCX_LETTER_SIDE_MARGIN)
    return int(page_width - left_margin - right_margin)


def _get_available_content_height_for_section(section: Section) -> int:
    # Provide alternative 'Letter' paper size params in case if they were not set explicitly in the document
    page_height = section.page_height or DOCX_LETTER_HEIGHT_EMU
    top_margin = _margin_size(section.top_margin, DOCX_LETTER_TOP_BOTTOM_MARGIN)
    bottom_margin = _margin_size(section.bottom_margin, DOCX_LETTER_TOP_BOTTOM_MARGIN)
    return int(page_height - top_margin - bottom_margin)


def _cap_image_heights(doc: DocumentObject) -> None:
    """Bring an image taller than its page back to the page, keeping its shape.

    pandoc brings a wide image back to the text width but leaves its height,
    so a tall one runs over several pages. The limit is the page height of
    the image's section less its top and bottom margins, and less the line
    the image is set on: see LINE_ALLOWANCE_EMU.
    """
    # A document always has a section: the body's own sectPr, or python-docx's default.
    max_heights = [_get_available_content_height_for_section(section) - LINE_ALLOWANCE_EMU for section in doc.sections]
    for element, section_index in _body_by_section(doc):
        max_height = max_heights[min(section_index, len(max_heights) - 1)]
        for extent in element.iter(f"{{{WP_NS}}}extent"):
            _cap_extent_height(extent, max_height)


def _cap_extent_height(extent: Any, max_height: int) -> None:
    """Scale one drawing down to max_height, its width by the same factor."""
    width = int(extent.get("cx", "0"))
    height = int(extent.get("cy", "0"))
    if max_height <= 0 or height <= max_height:
        return
    new_width = int(width * max_height / height)
    _set_drawing_size(extent, new_width, max_height)
    logger.debug(f"Capped image height: {width} x {height} -> {new_width} x {max_height}")


def _set_drawing_size(extent: Any, width: int, height: int) -> None:
    """Give a drawing a new size, in its <wp:extent> and in the frame of its picture.

    The picture's own frame states the size too, and Word can crop or distort a
    picture whose frame disagrees with its extent. An <a:ext> of an extension
    list carries a uri and no size, and is left alone.
    """
    extent.set("cx", str(width))
    extent.set("cy", str(height))
    for frame_ext in extent.getparent().iter(f"{{{DRAWING_NS}}}ext"):
        if frame_ext.get("cx") is not None:
            frame_ext.set("cx", str(width))
            frame_ext.set("cy", str(height))


def _resize_images_in_cell(cell: _Cell, max_image_width: float) -> None:
    # Python-docx exposes no public API for this element.
    cell_xml = cell._tc.xml  # noqa: SLF001
    # Use huge_tree parser to handle cells with large content (e.g., base64-encoded images > 10MB)
    tree = etree.fromstring(cell_xml, docx_parser.oxml_parser)

    # Find all <wp:extent> elements that define image size
    extent_elements = tree.findall(
        ".//wp:extent",
        {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"},
    )

    modified = False  # Track if any image was resized

    for extent in extent_elements:
        width = int(extent.attrib["cx"])
        height = int(extent.attrib["cy"])
        logger.debug(f"Image found, size: {width} x {height}")

        # Resize only if width exceeds max_image_width
        if width > max_image_width:
            scale_factor = max_image_width / width
            new_width = int(max_image_width)
            new_height = int(height * scale_factor)  # Maintain aspect ratio

            _set_drawing_size(extent, new_width, new_height)

            logger.debug(f"Resized to: {new_width} x {new_height}")
            modified = True

    # If any modification was made, update the cell XML
    if modified:
        # Python-docx exposes no public API for this element.
        cell._tc.clear_content()  # noqa: SLF001
        for child in tree.iterchildren():
            # Python-docx exposes no public API for this element.
            cell._tc.append(child)  # noqa: SLF001


# Command line shape for the manual entry point below.
MIN_ARGS = 2  # script name + docx path
MAX_ARGS = 4  # script name + docx path + paper_size + orientation
DOCX_PATH_ARG_INDEX = 1
PAPER_SIZE_ARG_INDEX = 2
ORIENTATION_ARG_INDEX = 3


# Just for manual test purposes. Accepts path to docx to process.
def main() -> int:
    if not (MIN_ARGS <= len(sys.argv) <= MAX_ARGS):
        logger.info("Usage: <path_to_docx> [paper_size] [orientation]")
        return 1

    docx_path = Path(sys.argv[DOCX_PATH_ARG_INDEX]).resolve()
    base_dir = Path.cwd().resolve()
    if not docx_path.is_relative_to(base_dir):
        logger.error(f"Refusing to access path outside the working directory: {docx_path}")
        return 1

    paper_size = sys.argv[PAPER_SIZE_ARG_INDEX] if len(sys.argv) > PAPER_SIZE_ARG_INDEX and sys.argv[PAPER_SIZE_ARG_INDEX] != "None" else None
    orientation = sys.argv[ORIENTATION_ARG_INDEX] if len(sys.argv) > ORIENTATION_ARG_INDEX and sys.argv[ORIENTATION_ARG_INDEX] != "None" else None

    with docx_path.open("rb") as docx_file_reader:
        result_bytes = process(docx_file_reader.read(), paper_size, orientation)

    with docx_path.open("wb") as docx_file_writer:
        docx_file_writer.write(result_bytes)

    logger.debug(f"Successfully modified table properties in {docx_path}")
    return 0


if __name__ == "__main__":
    main()
