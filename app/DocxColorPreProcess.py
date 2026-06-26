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

import logging
from xml.etree import ElementTree as ET

from .docx_ooxml import STYLES_PART, W_NS, augment_styles, enumerate_body_parts, parse_xml, read_entries, repack, serialize_tree

logger = logging.getLogger(__name__)

STYLE_PREFIX = "PandocColor"
SEGMENT_SEPARATOR = "__"
HEX_COLOR_LENGTH = 6
_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX byte-string with colored runs rewritten to use synthetic
    character styles. Returns the input unchanged when no colored runs are
    found or when the package is not a recognizable DOCX.
    """
    # Read the entire package into memory. DOCX files we handle are well
    # under the 200MB request cap and a full in-memory dict keeps the
    # rewrite loop simple — we mutate body parts in place and then
    # re-zip from the same dict at the end.
    entries = read_entries(docx_bytes)
    if entries is None:
        logger.warning("Input is not a valid zip / DOCX; skipping color preprocess")
        return docx_bytes

    # Without styles.xml there's nowhere to register the synthetic styles
    # we'd want to reference, so bail out early rather than fabricate one.
    if STYLES_PART not in entries:
        logger.debug("DOCX has no %s; skipping color preprocess", STYLES_PART)
        return docx_bytes

    # Collected across all body parts and deduplicated by style_id so
    # styles.xml gets one <w:style> per unique fg/bg/highlight combo even
    # if many runs reference it.
    needed_styles: dict[str, _StyleSpec] = {}

    for part in enumerate_body_parts(entries.keys()):
        rewritten, part_styles = _rewrite_part(entries[part])
        if part_styles:
            entries[part] = rewritten
            needed_styles.update(part_styles)

    # Fast path: no colored runs anywhere. Returning the original bytes
    # (instead of a re-zipped equivalent) preserves the original
    # compression layout and avoids touching styles.xml unnecessarily.
    if not needed_styles:
        return docx_bytes

    entries[STYLES_PART] = augment_styles(entries[STYLES_PART], needed_styles, _build_style_element)
    return repack(entries)


class _StyleSpec:
    """Lightweight value object describing a synthetic character style."""

    __slots__ = ("bg", "fg", "highlight", "size", "style_id")

    def __init__(self, style_id: str, fg: str | None, bg: str | None, highlight: str | None, size: str | None) -> None:
        self.style_id = style_id
        self.fg = fg
        self.bg = bg
        self.highlight = highlight
        self.size = size


def _style_id(fg: str | None, bg: str | None, highlight: str | None, size: str | None) -> str:
    # Fixed FG/BG/HL/SZ ordering keeps the style id deterministic: the same
    # formatting combination always produces the same id, which is what lets
    # the Lua filter pattern-match the encoded segments back out and what
    # makes deduplication across runs work.
    parts = [STYLE_PREFIX]
    if fg:
        parts.append(f"FG_{fg}")
    if bg:
        parts.append(f"BG_{bg}")
    if highlight:
        parts.append(f"HL_{highlight}")
    if size:
        parts.append(f"SZ_{size}")
    return SEGMENT_SEPARATOR.join(parts)


_COLOR_TAG = f"{{{W_NS}}}color"
_SHD_TAG = f"{{{W_NS}}}shd"
_HIGHLIGHT_TAG = f"{{{W_NS}}}highlight"
_SZ_TAG = f"{{{W_NS}}}sz"
_SZCS_TAG = f"{{{W_NS}}}szCs"
_RSTYLE_TAG = f"{{{W_NS}}}rStyle"
_RPR_TAG = f"{{{W_NS}}}rPr"
_R_TAG = f"{{{W_NS}}}r"
_VAL_ATTR = f"{{{W_NS}}}val"
_FILL_ATTR = f"{{{W_NS}}}fill"


def _extract_run_colors(rpr: ET.Element) -> tuple[str | None, str | None, str | None, str | None]:
    """Read fg/bg/highlight/size from a <w:rPr>, normalising and filtering out
    unusable values (theme references with no concrete value, the literal
    keyword "auto", or highlight="none").

    Four independent OOXML properties contribute:
      * ``<w:color w:val="RRGGBB"/>``    — text (foreground) color
      * ``<w:shd w:fill="RRGGBB"/>``     — paragraph/run shading background;
        the fill is on the ``w:fill`` attribute, not ``w:val`` (which carries
        the shading *pattern*, e.g. "clear")
      * ``<w:highlight w:val="yellow"/>`` — the legacy Word highlighter, whose
        value is a *named* color from a fixed palette ("yellow", "green",
        "cyan", ...), never a hex string
      * ``<w:sz w:val="32"/>``           — font size in *half-points* (32 = 16pt);
        pandoc's docx reader drops it, so it rides along in the synthetic style
        like the colors do
    """
    color_el = rpr.find(_COLOR_TAG)
    shd_el = rpr.find(_SHD_TAG)
    highlight_el = rpr.find(_HIGHLIGHT_TAG)
    sz_el = rpr.find(_SZ_TAG)

    fg = _normalize_hex(color_el.get(_VAL_ATTR)) if color_el is not None else None
    bg = _normalize_hex(shd_el.get(_FILL_ATTR)) if shd_el is not None else None
    highlight = highlight_el.get(_VAL_ATTR) if highlight_el is not None else None
    size = _normalize_half_points(sz_el.get(_VAL_ATTR)) if sz_el is not None else None

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
    return fg, bg, highlight, size


def _normalize_half_points(value: str | None) -> str | None:
    """Return the run font size (OOXML half-points) as a bare integer string, or
    None when absent/unparseable. The Lua filter halves it to points."""
    if not value:
        return None
    try:
        hp = int(value)
    except ValueError:
        return None
    return str(hp) if hp > 0 else None


def _replace_run_color_props(rpr: ET.Element, style_id: str) -> None:
    """Strip <w:color>/<w:shd>/<w:highlight>/<w:sz> from <w:rPr> and insert a
    single <w:rStyle> reference pointing at the synthetic style. Any
    existing <w:rStyle> is replaced (see module docstring).

    Stripping the direct properties is what tells pandoc to fall back to
    the style reference: if we left ``<w:color>`` in place pandoc would
    still drop it (it doesn't honor direct run colors), but the run would
    no longer carry the synthetic-style hint either, defeating the whole
    pipeline.
    """
    for tag in (_COLOR_TAG, _SHD_TAG, _HIGHLIGHT_TAG, _SZ_TAG, _SZCS_TAG, _RSTYLE_TAG):
        for el in rpr.findall(tag):
            rpr.remove(el)
    new_rstyle = ET.Element(_RSTYLE_TAG, {_VAL_ATTR: style_id})
    # <w:rStyle> must be the first child of <w:rPr> per the OOXML schema
    # (CT_RPr's sequence puts rStyle ahead of every formatting element).
    # Some readers — including Word in strict-mode — reject the document
    # when this ordering is violated.
    rpr.insert(0, new_rstyle)


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, dict[str, _StyleSpec]]:
    """Rewrite one body part. Returns (new_bytes, styles_used)."""
    tree = parse_xml(xml_bytes)
    if tree is None:
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
        fg, bg, highlight, size = _extract_run_colors(rpr)
        if not (fg or bg or highlight or size):
            continue

        style_id = _style_id(fg, bg, highlight, size)
        styles_used[style_id] = _StyleSpec(style_id, fg, bg, highlight, size)
        _replace_run_color_props(rpr, style_id)

    # Skip the re-serialize roundtrip when nothing changed. ET.tostring
    # reformats namespace declarations and attribute order, so even a
    # no-op rewrite would produce different bytes — undesirable for diffs
    # and unnecessary work.
    if not styles_used:
        return xml_bytes, {}

    return serialize_tree(tree), styles_used


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
    if spec.size:
        ET.SubElement(rpr, f"{{{W_NS}}}sz", {f"{{{W_NS}}}val": spec.size})
        ET.SubElement(rpr, f"{{{W_NS}}}szCs", {f"{{{W_NS}}}val": spec.size})
    return style
