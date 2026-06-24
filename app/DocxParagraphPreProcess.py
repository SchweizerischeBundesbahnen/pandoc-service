"""Rewrite paragraph alignment / indentation in a DOCX as paragraph styles.

Pandoc's DOCX reader drops paragraph alignment (``<w:jc>``) entirely and
coerces left indentation (``<w:ind w:left>``) into a single ``BlockQuote``
(merging distinct indent levels) before producing the AST, so no Lua filter
can recover those properties for the LaTeX/PDF writer. This preprocessor runs
*before* pandoc reads the DOCX and rewrites every paragraph that carries
alignment and/or a left indent into a reference to a synthetic *paragraph*
style named

    PandocPara__ALIGN_<align>__IND_<twips>     (each segment optional)

With ``-f docx+styles`` pandoc then surfaces those references as ``Div`` nodes
carrying ``custom-style="PandocPara__..."``, which
``filters/docx_paragraphs_to_latex.lua`` translates into a TeX group that sets
``\\leftskip`` and the matching alignment primitive (``\\centering`` /
``\\raggedleft`` / ``\\raggedright``).

This is the paragraph-level companion to :mod:`app.DocxColorPreProcess` (which
does the same trick for run-level colour via character styles); the two run
independently on the same docx→latex path and compose — a coloured run inside
an aligned paragraph yields ``Div[PandocPara] -> Para -> Span[PandocColor]``.

Alignment encoding: ``<w:jc>`` values are normalised to left/center/right;
``start``→left, ``end``→right; ``both``/``distribute`` (justified) are the
LaTeX default and carry no alignment segment (only an indent, if present).

Limitations
-----------
* Only ``<w:ind w:left>`` is carried; first-line / hanging indents are not.
* The paragraph's existing ``<w:pStyle>`` is replaced; its own properties are
  not preserved — pandoc would have dropped them anyway (only the style *name*
  survives the reader).
"""

from __future__ import annotations

import logging
from xml.etree import ElementTree as ET

from .docx_ooxml import STYLES_PART, W_NS, augment_styles, enumerate_body_parts, parse_xml, read_entries, repack, serialize_tree

logger = logging.getLogger(__name__)

STYLE_PREFIX = "PandocPara"
SEGMENT_SEPARATOR = "__"

# OOXML <w:jc> value -> normalised alignment segment, or None when the value
# needs no special handling. Only center and right are encoded:
#   * left/start is the default reading direction — LaTeX already left-aligns
#     where it matters, and encoding it wrapped every paragraph (notably the
#     left-aligned table-cell paragraphs pandoc emits) in \raggedright, which
#     needlessly perturbed rendering. Leave it to the default.
#   * both/distribute is justified, which is LaTeX's default.
_ALIGN_MAP: dict[str, str | None] = {
    "left": None,
    "start": None,
    "center": "center",
    "right": "right",
    "end": "right",
    "both": None,
    "distribute": None,
}

_P_TAG = f"{{{W_NS}}}p"
_TC_TAG = f"{{{W_NS}}}tc"
_PPR_TAG = f"{{{W_NS}}}pPr"
_PSTYLE_TAG = f"{{{W_NS}}}pStyle"
_JC_TAG = f"{{{W_NS}}}jc"
_IND_TAG = f"{{{W_NS}}}ind"
_VAL_ATTR = f"{{{W_NS}}}val"
_LEFT_ATTR = f"{{{W_NS}}}left"


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX with aligned/indented paragraphs rewritten to use synthetic
    paragraph styles. Returns the input unchanged when no such paragraphs are
    found or when the package is not a recognizable DOCX.
    """
    entries = read_entries(docx_bytes)
    if entries is None:
        logger.warning("Input is not a valid zip / DOCX; skipping paragraph preprocess")
        return docx_bytes

    if STYLES_PART not in entries:
        logger.debug("DOCX has no %s; skipping paragraph preprocess", STYLES_PART)
        return docx_bytes

    # Deduplicated by style_id across all body parts so styles.xml gets one
    # <w:style> per unique align/indent combo even if many paragraphs use it.
    needed_styles: dict[str, _StyleSpec] = {}

    for part in enumerate_body_parts(entries.keys()):
        rewritten, part_styles = _rewrite_part(entries[part])
        if part_styles:
            entries[part] = rewritten
            needed_styles.update(part_styles)

    if not needed_styles:
        return docx_bytes

    entries[STYLES_PART] = augment_styles(entries[STYLES_PART], needed_styles, _build_style_element)
    return repack(entries)


class _StyleSpec:
    """Lightweight value object describing a synthetic paragraph style."""

    __slots__ = ("align", "indent_twips", "style_id")

    def __init__(self, style_id: str, align: str | None, indent_twips: int | None) -> None:
        self.style_id = style_id
        self.align = align
        self.indent_twips = indent_twips


def _style_id(align: str | None, indent_twips: int | None) -> str:
    # Fixed ALIGN/IND ordering keeps the id deterministic: the same combination
    # always produces the same id, which is what lets the Lua filter pattern-
    # match the segments back out and what makes deduplication work.
    parts = [STYLE_PREFIX]
    if align:
        parts.append(f"ALIGN_{align}")
    if indent_twips is not None:
        parts.append(f"IND_{indent_twips}")
    return SEGMENT_SEPARATOR.join(parts)


def _extract_para_format(ppr: ET.Element) -> tuple[str | None, int | None]:
    """Read (alignment segment, left-indent twips) from a <w:pPr>.

    Returns (None, None) when there is nothing worth encoding (no alignment
    that changes rendering, no positive left indent).
    """
    align: str | None = None
    jc_el = ppr.find(_JC_TAG)
    if jc_el is not None:
        raw = jc_el.get(_VAL_ATTR)
        if raw is not None:
            align = _ALIGN_MAP.get(raw.strip().lower())

    indent: int | None = None
    ind_el = ppr.find(_IND_TAG)
    if ind_el is not None:
        left = ind_el.get(_LEFT_ATTR)
        if left is not None:
            try:
                value = int(left)
            except ValueError:
                value = 0
            if value > 0:
                indent = value

    return align, indent


def _replace_para_props(ppr: ET.Element, style_id: str) -> None:
    """Strip <w:pStyle>/<w:jc>/<w:ind> from <w:pPr> and insert a single
    <w:pStyle> reference to the synthetic style as the first child.

    Stripping <w:jc>/<w:ind> is what stops pandoc from dropping the alignment
    and coercing the indent into a BlockQuote; the style reference carries the
    information across the reader instead. <w:pStyle> must be the first child
    of <w:pPr> per the OOXML schema (CT_PPr sequence).
    """
    for tag in (_PSTYLE_TAG, _JC_TAG, _IND_TAG):
        for el in ppr.findall(tag):
            ppr.remove(el)
    ppr.insert(0, ET.Element(_PSTYLE_TAG, {_VAL_ATTR: style_id}))


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, dict[str, _StyleSpec]]:
    """Rewrite one body part. Returns (new_bytes, styles_used)."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping")
        return xml_bytes, {}

    styles_used: dict[str, _StyleSpec] = {}

    # Paragraphs inside table cells are left alone: pandoc's docx reader already
    # renders cell alignment/indent from the table structure (it emits
    # >{\centering\arraybackslash}p{…} column types), so rewriting them is
    # redundant and the per-cell wrapper changes row spacing. ElementTree has no
    # parent pointers, so collect cell paragraphs up front and skip them below.
    in_cell = {id(p) for tc in tree.iter(_TC_TAG) for p in tc.iter(_P_TAG)}

    for para in tree.iter(_P_TAG):
        if id(para) in in_cell:
            continue
        ppr = para.find(_PPR_TAG)
        if ppr is None:
            continue
        align, indent = _extract_para_format(ppr)
        if not align and not indent:
            continue
        style_id = _style_id(align, indent)
        styles_used[style_id] = _StyleSpec(style_id, align, indent)
        _replace_para_props(ppr, style_id)

    if not styles_used:
        return xml_bytes, {}

    return serialize_tree(tree), styles_used


def _build_style_element(spec: _StyleSpec) -> ET.Element:
    # w:type="paragraph" — a paragraph style (referenced by <w:pStyle>).
    # Pandoc keys its custom-style attribute off <w:name w:val=...>, so name
    # and styleId both carry the encoded segments. The <w:pPr> below keeps the
    # style usable if the intermediate DOCX is inspected manually; pandoc only
    # consumes the styleId + name.
    style = ET.Element(
        f"{{{W_NS}}}style",
        {
            f"{{{W_NS}}}type": "paragraph",
            f"{{{W_NS}}}customStyle": "1",
            f"{{{W_NS}}}styleId": spec.style_id,
        },
    )
    ET.SubElement(style, f"{{{W_NS}}}name", {f"{{{W_NS}}}val": spec.style_id})
    ppr = ET.SubElement(style, _PPR_TAG)
    if spec.indent_twips is not None:
        ET.SubElement(ppr, _IND_TAG, {_LEFT_ATTR: str(spec.indent_twips)})
    if spec.align:
        # Only center/right are ever encoded (see _ALIGN_MAP), and both map
        # straight back to the matching Word justification value.
        ET.SubElement(ppr, _JC_TAG, {_VAL_ATTR: spec.align})
    return style
