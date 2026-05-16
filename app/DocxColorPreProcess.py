"""Rewrite direct run-level color formatting in a DOCX as character styles.

Pandoc's DOCX reader drops direct character formatting (``<w:color>``,
``<w:shd>``, ``<w:highlight>``) before producing the AST, so no Lua filter
can recover those properties for the LaTeX/PDF writer. This preprocessor
runs *before* pandoc reads the DOCX and converts every colored run into a
reference to a synthetic character style. With ``-f docx+styles`` pandoc
then emits ``Span`` nodes carrying ``custom-style="PandocColor__..."``
attributes, which ``filters/docx_colors_to_latex.lua`` translates into
``\\textcolor`` / ``\\colorbox`` raw LaTeX.

Style-name encoding (parseable, deterministic, segments are optional):

    PandocColor__FG_RRGGBB__BG_RRGGBB__HL_<wordHighlightName>

Limitations
-----------
* Theme colors (``<w:color w:themeColor="accent1"/>``) are not resolved
  against ``word/theme/theme1.xml``; theme-colored runs stay uncolored.
* Color applied indirectly via a paragraph or character style chain (not
  as a direct run property) is not picked up — pandoc would have dropped
  it anyway because only the style *name* survives the reader, never the
  style's character properties.
* If a run already carries an ``<w:rStyle>``, that reference is replaced
  by the synthetic one. The original style's name is therefore not
  visible to pandoc; its character properties were going to be dropped
  by pandoc anyway, so nothing colored is lost.
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

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Well-known OOXML namespaces. ElementTree's default behavior on serialize
# is to mint synthetic prefixes (ns0, ns1, ...) for every namespace it
# wasn't told about — fine for round-tripping inside ET, but DOCX readers
# (notably pandoc's docx reader) match drawing/relationship/etc. elements
# on the canonical prefix rather than the URI. When wp:/a:/pic:/r: get
# renamed to ns1:/ns2:/..., pandoc silently drops every <w:drawing> in
# the file and images disappear from the output. Registering the
# canonical prefixes globally tells ET to preserve them on serialize.
_OOXML_NAMESPACES = {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
}
for _prefix, _uri in _OOXML_NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)

STYLES_PART = "word/styles.xml"
STYLE_PREFIX = "PandocColor"
SEGMENT_SEPARATOR = "__"
HEX_COLOR_LENGTH = 6
_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")

# Parts that can contain <w:r> runs we need to rewrite. Theme / numbering
# / settings parts cannot carry runs and are skipped.
_FIXED_BODY_PARTS = frozenset(
    {
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    }
)


def _enumerate_body_parts(names: Iterable[str]) -> list[str]:
    """Return zip entry names that may contain runs to rewrite.

    Headers and footers live in numbered parts (``word/header1.xml``,
    ``word/footer2.xml``, ...) — the count and naming depend on how many
    sections the document defines — so they can't be enumerated by a
    fixed name list and have to be matched by prefix + ``.xml`` suffix.
    """
    result = []
    for name in names:
        if name in _FIXED_BODY_PARTS or (name.startswith("word/header") or name.startswith("word/footer")) and name.endswith(".xml"):
            result.append(name)
    return result


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX byte-string with colored runs rewritten to use synthetic
    character styles. Returns the input unchanged when no colored runs are
    found or when the package is not a recognizable DOCX.
    """
    # Read the entire package into memory. DOCX files we handle are well
    # under the 200MB request cap and a full in-memory dict keeps the
    # rewrite loop simple — we mutate body parts in place and then
    # re-zip from the same dict at the end.
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
            entries = {name: zin.read(name) for name in zin.namelist()}
    except zipfile.BadZipFile:
        logger.warning("Input is not a valid zip / DOCX; skipping color preprocess")
        return docx_bytes

    # Without styles.xml there's nowhere to register the synthetic styles
    # we'd want to reference, so bail out early rather than fabricate one.
    if STYLES_PART not in entries:
        logger.debug("DOCX has no %s; skipping color preprocess", STYLES_PART)
        return docx_bytes

    body_parts = _enumerate_body_parts(entries.keys())
    if not body_parts:
        return docx_bytes

    # Collected across all body parts and deduplicated by style_id so
    # styles.xml gets one <w:style> per unique fg/bg/highlight combo even
    # if many runs reference it.
    needed_styles: dict[str, _StyleSpec] = {}

    for part in body_parts:
        rewritten, part_styles = _rewrite_part(entries[part])
        if part_styles:
            entries[part] = rewritten
            needed_styles.update(part_styles)

    # Fast path: no colored runs anywhere. Returning the original bytes
    # (instead of a re-zipped equivalent) preserves the original
    # compression layout and avoids touching styles.xml unnecessarily.
    if not needed_styles:
        return docx_bytes

    entries[STYLES_PART] = _augment_styles(entries[STYLES_PART], needed_styles)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    return buf.getvalue()


class _StyleSpec:
    """Lightweight value object describing a synthetic character style."""

    __slots__ = ("bg", "fg", "highlight", "style_id")

    def __init__(self, style_id: str, fg: str | None, bg: str | None, highlight: str | None) -> None:
        self.style_id = style_id
        self.fg = fg
        self.bg = bg
        self.highlight = highlight


def _style_id(fg: str | None, bg: str | None, highlight: str | None) -> str:
    # Fixed FG/BG/HL ordering keeps the style id deterministic: the same
    # color combination always produces the same id, which is what lets
    # the Lua filter pattern-match the encoded segments back out and what
    # makes deduplication across runs work.
    parts = [STYLE_PREFIX]
    if fg:
        parts.append(f"FG_{fg}")
    if bg:
        parts.append(f"BG_{bg}")
    if highlight:
        parts.append(f"HL_{highlight}")
    return SEGMENT_SEPARATOR.join(parts)


_COLOR_TAG = f"{{{W_NS}}}color"
_SHD_TAG = f"{{{W_NS}}}shd"
_HIGHLIGHT_TAG = f"{{{W_NS}}}highlight"
_RSTYLE_TAG = f"{{{W_NS}}}rStyle"
_RPR_TAG = f"{{{W_NS}}}rPr"
_R_TAG = f"{{{W_NS}}}r"
_VAL_ATTR = f"{{{W_NS}}}val"
_FILL_ATTR = f"{{{W_NS}}}fill"


def _extract_run_colors(rpr: ET.Element) -> tuple[str | None, str | None, str | None]:
    """Read fg/bg/highlight from a <w:rPr> element, normalising and filtering
    out unusable values (theme references with no concrete value, the literal
    keyword "auto", or highlight="none").

    Three independent OOXML properties contribute:
      * ``<w:color w:val="RRGGBB"/>``    — text (foreground) color
      * ``<w:shd w:fill="RRGGBB"/>``     — paragraph/run shading background;
        the fill is on the ``w:fill`` attribute, not ``w:val`` (which carries
        the shading *pattern*, e.g. "clear")
      * ``<w:highlight w:val="yellow"/>`` — the legacy Word highlighter, whose
        value is a *named* color from a fixed palette ("yellow", "green",
        "cyan", ...), never a hex string
    """
    color_el = rpr.find(_COLOR_TAG)
    shd_el = rpr.find(_SHD_TAG)
    highlight_el = rpr.find(_HIGHLIGHT_TAG)

    fg = _normalize_hex(color_el.get(_VAL_ATTR)) if color_el is not None else None
    bg = _normalize_hex(shd_el.get(_FILL_ATTR)) if shd_el is not None else None
    highlight = highlight_el.get(_VAL_ATTR) if highlight_el is not None else None

    # "auto" means "use whatever the consumer's default is" (typically black
    # text on white background). Treating it as a real color would emit
    # \textcolor{000000} for every plain run, which is both noisy and wrong
    # when the surrounding theme picks something else.
    if fg == "AUTO":
        fg = None
    if bg == "AUTO":
        bg = None
    # highlight="none" is the OOXML idiom for "no highlight set"; some
    # editors emit it explicitly rather than omitting the element.
    if highlight == "none":
        highlight = None
    return fg, bg, highlight


def _replace_run_color_props(rpr: ET.Element, style_id: str) -> None:
    """Strip <w:color>/<w:shd>/<w:highlight> from <w:rPr> and insert a
    single <w:rStyle> reference pointing at the synthetic style. Any
    existing <w:rStyle> is replaced (see module docstring).

    Stripping the direct properties is what tells pandoc to fall back to
    the style reference: if we left ``<w:color>`` in place pandoc would
    still drop it (it doesn't honor direct run colors), but the run would
    no longer carry the synthetic-style hint either, defeating the whole
    pipeline.
    """
    # findall returns a live view in some ET implementations; materialize
    # to a list before mutating to avoid skipping siblings.
    for tag in (_COLOR_TAG, _SHD_TAG, _HIGHLIGHT_TAG, _RSTYLE_TAG):
        for el in list(rpr.findall(tag)):
            rpr.remove(el)
    new_rstyle = ET.Element(_RSTYLE_TAG, {_VAL_ATTR: style_id})
    # <w:rStyle> must be the first child of <w:rPr> per the OOXML schema
    # (CT_RPr's sequence puts rStyle ahead of every formatting element).
    # Some readers — including Word in strict-mode — reject the document
    # when this ordering is violated.
    rpr.insert(0, new_rstyle)


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, dict[str, _StyleSpec]]:
    """Rewrite one body part. Returns (new_bytes, styles_used)."""
    # defusedxml's fromstring blocks XXE / billion-laughs; we still serialize
    # with stdlib ET because defusedxml deliberately doesn't expose tostring.
    try:
        tree = DefusedET.fromstring(xml_bytes)
    except ET.ParseError:
        logger.warning("Unparseable XML in DOCX part; skipping")
        return xml_bytes, {}

    styles_used: dict[str, _StyleSpec] = {}

    # iter() walks every <w:r> descendant regardless of depth — runs can be
    # nested inside tables, structured-document-tags, content controls, etc.,
    # so a top-level findall would miss most of them.
    for run in tree.iter(_R_TAG):
        rpr = run.find(_RPR_TAG)
        if rpr is None:
            # No run-properties element means no direct color formatting.
            continue
        fg, bg, highlight = _extract_run_colors(rpr)
        if not (fg or bg or highlight):
            continue

        style_id = _style_id(fg, bg, highlight)
        styles_used[style_id] = _StyleSpec(style_id, fg, bg, highlight)
        _replace_run_color_props(rpr, style_id)

    # Skip the re-serialize roundtrip when nothing changed. ET.tostring
    # reformats namespace declarations and attribute order, so even a
    # no-op rewrite would produce different bytes — undesirable for diffs
    # and unnecessary work.
    if not styles_used:
        return xml_bytes, {}

    new_xml = ET.tostring(tree, xml_declaration=True, encoding="UTF-8")
    return new_xml, styles_used


def _normalize_hex(value: str | None) -> str | None:
    """Uppercase a 6-digit hex string. Returns 'AUTO' for the literal Word
    keyword "auto" (so the caller knows to drop the element).

    Uppercasing matters: the hex appears inside the synthetic style id
    (``PandocColor__FG_FF0000``), and two runs with ``ff0000`` vs
    ``FF0000`` must collapse onto the same style. The Lua filter also
    pattern-matches on uppercase segments.
    """
    if not value:
        return None
    stripped = value.strip()
    if stripped.lower() == "auto":
        return "AUTO"
    # Accept "RRGGBB", "#RRGGBB". The OOXML schema specifies bare
    # ``RRGGBB``, but the leading ``#`` shows up in the wild from
    # third-party tools, so accept it. Anything else (theme color refs
    # like ``accent1``, named CSS colors, malformed strings) we
    # conservatively drop — we can't produce a usable hex from them and
    # rather have no color than the wrong one.
    if stripped.startswith("#"):
        stripped = stripped[1:]
    if len(stripped) == HEX_COLOR_LENGTH and all(c in _HEX_DIGITS for c in stripped):
        return stripped.upper()
    return None


def _augment_styles(styles_xml: bytes, specs: dict[str, _StyleSpec]) -> bytes:
    """Insert <w:style> entries for each spec into word/styles.xml, idempotently.

    The fragment we add looks like::

        <w:style w:type="character" w:customStyle="1" w:styleId="PandocColor__FG_FF0000">
            <w:name w:val="PandocColor__FG_FF0000"/>
            <w:rPr>
                <w:color w:val="FF0000"/>
            </w:rPr>
        </w:style>

    Adding the matching ``<w:rPr>`` keeps the style usable when the
    intermediate DOCX is inspected manually; pandoc only consumes the
    ``w:styleId`` and ``w:name``.
    """
    try:
        tree = DefusedET.fromstring(styles_xml)
    except ET.ParseError:
        logger.warning("Unparseable %s; skipping style augmentation", STYLES_PART)
        return styles_xml

    # Collect existing styleIds to make this idempotent: if the document
    # has already been preprocessed (or genuinely defined a style with
    # the same name) we must not append a duplicate <w:style> — Word
    # rejects styles.xml with two entries sharing the same styleId.
    existing_ids = {el.get(f"{{{W_NS}}}styleId") for el in tree.findall(f"{{{W_NS}}}style")}

    for spec in specs.values():
        if spec.style_id in existing_ids:
            continue
        # Appending to the root <w:styles> element is sufficient — style
        # order in styles.xml has no semantic meaning, only the styleId
        # link from runs matters.
        tree.append(_build_style_element(spec))

    return ET.tostring(tree, xml_declaration=True, encoding="UTF-8")


def _build_style_element(spec: _StyleSpec) -> ET.Element:
    # w:type="character" — a *character* style, not paragraph/table; the
    # only kind that can be referenced by <w:rStyle> on a run.
    # w:customStyle="1" — flags this as a user/tool-defined style rather
    # than a built-in Word style, so Word's style gallery treats it
    # accordingly and doesn't try to localize the name.
    style = ET.Element(
        f"{{{W_NS}}}style",
        {
            f"{{{W_NS}}}type": "character",
            f"{{{W_NS}}}customStyle": "1",
            f"{{{W_NS}}}styleId": spec.style_id,
        },
    )
    # Pandoc keys its ``custom-style`` attribute off ``<w:name w:val=...>``,
    # not ``w:styleId``. Using the same string for both keeps the encoded
    # FG/BG/HL segments reachable to the Lua filter via either attribute.
    ET.SubElement(style, f"{{{W_NS}}}name", {f"{{{W_NS}}}val": spec.style_id})
    rpr = ET.SubElement(style, f"{{{W_NS}}}rPr")
    if spec.fg:
        ET.SubElement(rpr, f"{{{W_NS}}}color", {f"{{{W_NS}}}val": spec.fg})
    if spec.bg:
        # w:shd needs three attributes to render a solid background:
        # ``val="clear"`` (no pattern overlay), ``color="auto"`` (the
        # pattern color is irrelevant when val=clear, but the attribute
        # is required by the schema), and ``fill`` carrying the actual
        # background hex.
        ET.SubElement(
            rpr,
            f"{{{W_NS}}}shd",
            {
                f"{{{W_NS}}}val": "clear",
                f"{{{W_NS}}}color": "auto",
                f"{{{W_NS}}}fill": spec.bg,
            },
        )
    if spec.highlight:
        ET.SubElement(rpr, f"{{{W_NS}}}highlight", {f"{{{W_NS}}}val": spec.highlight})
    return style
