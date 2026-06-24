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

import io
import logging
import zipfile
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from defusedxml import ElementTree as DefusedET

if TYPE_CHECKING:
    from collections.abc import Iterable

logger = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced

# Canonical OOXML prefixes. ElementTree mints synthetic prefixes (ns0, ns1, …)
# for namespaces it wasn't told about, which makes pandoc's docx reader drop
# every <w:drawing> (images vanish). Registering the canonical prefixes keeps
# them on serialize. See app/DocxColorPreProcess.py for the full rationale.
_OOXML_NAMESPACES = {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
}
for _prefix, _uri in _OOXML_NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)

STYLES_PART = "word/styles.xml"
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

_FIXED_BODY_PARTS = frozenset(
    {
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    }
)

_P_TAG = f"{{{W_NS}}}p"
_TC_TAG = f"{{{W_NS}}}tc"
_PPR_TAG = f"{{{W_NS}}}pPr"
_PSTYLE_TAG = f"{{{W_NS}}}pStyle"
_JC_TAG = f"{{{W_NS}}}jc"
_IND_TAG = f"{{{W_NS}}}ind"
_VAL_ATTR = f"{{{W_NS}}}val"
_LEFT_ATTR = f"{{{W_NS}}}left"


def _enumerate_body_parts(names: Iterable[str]) -> list[str]:
    """Return zip entry names that may contain paragraphs to rewrite.

    Headers/footers live in numbered parts (``word/header1.xml`` …) whose count
    depends on the section layout, so they're matched by prefix + ``.xml``.
    """
    result = []
    for name in names:
        if name in _FIXED_BODY_PARTS or (name.startswith(("word/header", "word/footer")) and name.endswith(".xml")):
            result.append(name)
    return result


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX with aligned/indented paragraphs rewritten to use synthetic
    paragraph styles. Returns the input unchanged when no such paragraphs are
    found or when the package is not a recognizable DOCX.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
            entries = {name: zin.read(name) for name in zin.namelist()}
    except zipfile.BadZipFile:
        logger.warning("Input is not a valid zip / DOCX; skipping paragraph preprocess")
        return docx_bytes

    if STYLES_PART not in entries:
        logger.debug("DOCX has no %s; skipping paragraph preprocess", STYLES_PART)
        return docx_bytes

    body_parts = _enumerate_body_parts(entries.keys())
    if not body_parts:
        return docx_bytes

    # Deduplicated by style_id across all body parts so styles.xml gets one
    # <w:style> per unique align/indent combo even if many paragraphs use it.
    needed_styles: dict[str, _StyleSpec] = {}

    for part in body_parts:
        rewritten, part_styles = _rewrite_part(entries[part])
        if part_styles:
            entries[part] = rewritten
            needed_styles.update(part_styles)

    if not needed_styles:
        return docx_bytes

    entries[STYLES_PART] = _augment_styles(entries[STYLES_PART], needed_styles)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    return buf.getvalue()


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
    if indent_twips:
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
    try:
        tree = DefusedET.fromstring(xml_bytes)
    except ET.ParseError:
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

    new_xml = ET.tostring(tree, xml_declaration=True, encoding="UTF-8")
    return new_xml, styles_used


def _augment_styles(styles_xml: bytes, specs: dict[str, _StyleSpec]) -> bytes:
    """Insert a <w:style w:type="paragraph"> entry for each spec, idempotently."""
    try:
        tree = DefusedET.fromstring(styles_xml)
    except ET.ParseError:
        logger.warning("Unparseable %s; skipping style augmentation", STYLES_PART)
        return styles_xml

    # Idempotency: never append a second <w:style> with an existing styleId
    # (Word rejects styles.xml with duplicate styleIds).
    existing_ids = {el.get(f"{{{W_NS}}}styleId") for el in tree.findall(f"{{{W_NS}}}style")}

    for spec in specs.values():
        if spec.style_id in existing_ids:
            continue
        tree.append(_build_style_element(spec))

    return ET.tostring(tree, xml_declaration=True, encoding="UTF-8")


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
    if spec.indent_twips:
        ET.SubElement(ppr, _IND_TAG, {_LEFT_ATTR: str(spec.indent_twips)})
    if spec.align:
        # Only center/right are ever encoded (see _ALIGN_MAP), and both map
        # straight back to the matching Word justification value.
        ET.SubElement(ppr, _JC_TAG, {_VAL_ATTR: spec.align})
    return style
