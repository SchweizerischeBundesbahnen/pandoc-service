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
* A character style's colors (``<w:rStyle w:val="Hyperlink"/>`` pointing at
  a style with ``<w:color>``, followed through ``basedOn``) are applied to a
  run that has no direct color of its own, the way Word draws it: pandoc
  keeps only the style *name*, never its character properties, so a link
  colored by the reference template's Hyperlink style would otherwise come
  out black. A direct property of the run, ``auto`` included, wins over the
  style's. Colors set by a *paragraph* style are not picked up.
* A character style based on ``Verbatim Char`` gives its color to no run:
  pandoc reads such a run as code through that style, and a synthetic style
  in its place would set a highlighted token in the font of the text.
* If a run already carries an ``<w:rStyle>``, that reference is replaced
  by the synthetic one. The original style's name is therefore not
  visible to pandoc; its character properties were going to be dropped
  by pandoc anyway, so nothing colored is lost.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from .docx_ooxml import STYLES_PART, W_NS, augment_styles, enumerate_body_parts, parse_xml, read_entries, repack, serialize_tree

if TYPE_CHECKING:
    from collections.abc import Collection, Mapping

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

    style_colors = character_style_colors(entries[STYLES_PART])
    toc_styles = table_of_contents_styles(entries[STYLES_PART])
    for part in enumerate_body_parts(entries.keys()):
        rewritten, part_styles = rewrite_part(entries[part], style_colors, toc_styles)
        if part_styles:
            entries[part] = rewritten
            needed_styles.update(part_styles)

            # Fast path: no colored runs anywhere. Returning the original bytes
            # (instead of a re-zipped equivalent) preserves the original
            # compression layout and avoids touching styles.xml unnecessarily.
    if not needed_styles:
        return docx_bytes

    entries[STYLES_PART] = augment_styles(entries[STYLES_PART], needed_styles, build_style_element)
    return repack(entries)


@dataclass(slots=True)
class _StyleSpec:
    """Lightweight value object describing a synthetic character style."""

    style_id: str
    fg: str | None
    bg: str | None
    highlight: str | None
    size: str | None
    size_cs: str | None


def _style_id(fg: str | None, bg: str | None, highlight: str | None, size: str | None, size_cs: str | None) -> str:
    # Fixed FG/BG/HL/SZ ordering keeps the style id deterministic: the same
    # formatting combination always produces the same id, which is what lets
    # the Lua filter pattern-match the encoded segments back out and what
    # makes deduplication across runs work. SZCS is included only when it
    # differs from SZ (complex-script text) so it stays part of the identity
    # without bloating the common case where the two sizes match.
    parts = [STYLE_PREFIX]
    if fg:
        parts.append(f"FG_{fg}")
    if bg:
        parts.append(f"BG_{bg}")
    if highlight:
        parts.append(f"HL_{highlight}")
    if size:
        parts.append(f"SZ_{size}")
    if size_cs and size_cs != size:
        parts.append(f"SZCS_{size_cs}")
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


def _extract_run_colors(rpr: ET.Element) -> tuple[str | None, str | None, str | None, str | None, str | None]:
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
      * ``<w:szCs w:val="32"/>``         — the complex-script (Arabic/Hebrew/…)
        font size, which can legitimately differ from ``w:sz``; preserved
        separately so it isn't silently overwritten with the ``w:sz`` value
    """
    color_el = rpr.find(_COLOR_TAG)
    shd_el = rpr.find(_SHD_TAG)
    highlight_el = rpr.find(_HIGHLIGHT_TAG)
    sz_el = rpr.find(_SZ_TAG)
    szcs_el = rpr.find(_SZCS_TAG)

    fg = _normalize_hex(color_el.get(_VAL_ATTR)) if color_el is not None else None
    bg = _normalize_hex(shd_el.get(_FILL_ATTR)) if shd_el is not None else None
    highlight = highlight_el.get(_VAL_ATTR) if highlight_el is not None else None
    size = _normalize_half_points(sz_el.get(_VAL_ATTR)) if sz_el is not None else None
    size_cs = _normalize_half_points(szcs_el.get(_VAL_ATTR)) if szcs_el is not None else None

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
    return fg, bg, highlight, size, size_cs


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


@dataclass(frozen=True, slots=True)
class StyleColors:
    """The colors a character style draws its text with, after ``basedOn``."""

    fg: str | None
    bg: str | None
    highlight: str | None


_STYLE_TAG = f"{{{W_NS}}}style"
_BASED_ON_TAG = f"{{{W_NS}}}basedOn"
_NAME_TAG = f"{{{W_NS}}}name"
_TYPE_ATTR = f"{{{W_NS}}}type"
_STYLE_ID_ATTR = f"{{{W_NS}}}styleId"
_GENERATED_STYLE_PREFIX = STYLE_PREFIX + SEGMENT_SEPARATOR
# pandoc marks a run as code when this name is on the basedOn chain of its style, whatever the case
_CODE_STYLE_NAME = "verbatim char"
# Word translates the id of a built-in style (Verzeichnis1, TM1), but never its name
_TOC_STYLE_NAME = re.compile(r"toc [1-9]|table of figures")


@dataclass(frozen=True, slots=True)
class _StyleProps:
    """What a character style says about its colors.

    ``None`` leaves a property to the style it is based on; ``""`` clears it,
    which is how ``auto`` and ``none`` read here.
    """

    fg: str | None = None
    bg: str | None = None
    highlight: str | None = None
    code: bool = False

    def over(self, parent: _StyleProps) -> _StyleProps:
        return _StyleProps(
            parent.fg if self.fg is None else self.fg,
            parent.bg if self.bg is None else self.bg,
            parent.highlight if self.highlight is None else self.highlight,
            self.code or parent.code,
        )


def _own_style_props(style: ET.Element) -> _StyleProps:
    name = style.find(_NAME_TAG)
    code = name is not None and name.get(_VAL_ATTR, "").lower() == _CODE_STYLE_NAME
    rpr = style.find(_RPR_TAG)
    if rpr is None:
        return _StyleProps(code=code)
    fg, bg, highlight = _extract_run_colors(rpr)[:3]
    return _StyleProps(
        fg or ("" if rpr.find(_COLOR_TAG) is not None else None),
        bg or ("" if rpr.find(_SHD_TAG) is not None else None),
        highlight or ("" if rpr.find(_HIGHLIGHT_TAG) is not None else None),
        code,
    )


def _resolve_based_on(own: Mapping[str, _StyleProps], based_on: Mapping[str, str]) -> dict[str, _StyleProps]:
    """Merge each style with the styles it is based on, the nearest value first.

    Every style is visited a fixed number of times, however long the chains: an
    uploaded document decides how many styles there are, and this runs before
    pandoc, in a conversion slot.
    """
    resolved: dict[str, _StyleProps] = {}
    for style_id in own:
        chain: list[str] = []
        position: dict[str, int] = {}
        current = style_id
        while current in own and current not in resolved and current not in position:
            position[current] = len(chain)
            chain.append(current)
            current = based_on.get(current, "")
        # A hand-edited template may chain styles into a cycle. Going round it
        # twice lets each of its styles see all the others, the nearest first.
        start = position.get(current, len(chain))
        props = resolved.get(current, _StyleProps())
        for link in reversed(chain[:start] + chain[start:] * 2):
            props = own[link].over(props)
            resolved[link] = props
    return resolved


def character_style_colors(styles_xml: bytes) -> dict[str, StyleColors]:
    """Map each character style of ``styles.xml`` that draws with a color to
    those colors, following its ``basedOn`` chain.

    A style's own property wins over the one it inherits, as in Word, and that
    includes clearing it with ``auto`` or ``none``. The concrete ``w:val`` is
    used even when the color names a theme color too (Word writes both), so
    ``accent1`` styles resolve without reading the theme. Styles without any
    color are left out, and so are the styles of code (see the module
    docstring) and the synthetic styles of an earlier pass, whose runs are
    rewritten already.
    """
    root = parse_xml(styles_xml)
    if root is None:
        return {}
    own: dict[str, _StyleProps] = {}
    based_on: dict[str, str] = {}
    for style in root.iter(_STYLE_TAG):
        if style.get(_TYPE_ATTR) != "character":
            continue
        style_id = style.get(_STYLE_ID_ATTR)
        if not style_id or style_id.startswith(_GENERATED_STYLE_PREFIX):
            continue
        own[style_id] = _own_style_props(style)
        parent = style.find(_BASED_ON_TAG)
        if parent is not None and parent.get(_VAL_ATTR):
            based_on[style_id] = parent.get(_VAL_ATTR, "")

    colors: dict[str, StyleColors] = {}
    for style_id, props in _resolve_based_on(own, based_on).items():
        if not props.code and (props.fg or props.bg or props.highlight):
            colors[style_id] = StyleColors(props.fg or None, props.bg or None, props.highlight or None)
    return colors


def table_of_contents_styles(styles_xml: bytes) -> frozenset[str]:
    """The ids of the paragraph styles of ``styles.xml`` for the entries of a
    table of contents, figures or tables, read from their names."""
    root = parse_xml(styles_xml)
    if root is None:
        return frozenset()
    style_ids: set[str] = set()
    for style in root.iter(_STYLE_TAG):
        name = style.find(_NAME_TAG)
        if style.get(_TYPE_ATTR) == "paragraph" and name is not None and _TOC_STYLE_NAME.fullmatch(name.get(_VAL_ATTR, "").lower()):
            style_ids.add(style.get(_STYLE_ID_ATTR, ""))
    return frozenset(style_ids)


def _inherit_style_colors(rpr: ET.Element, style_colors: Mapping[str, StyleColors], fg: str | None, bg: str | None, highlight: str | None) -> tuple[str | None, str | None, str | None]:
    """Fill in the colors a run takes from its character style.

    Only a property the run does not set itself is taken: a direct
    ``<w:color w:val="auto"/>`` still means "no color" for that run, as in
    Word, even though it reads as ``None`` here.
    """
    rstyle = rpr.find(_RSTYLE_TAG)
    inherited = style_colors.get(rstyle.get(_VAL_ATTR, "")) if rstyle is not None else None
    if inherited is None:
        return fg, bg, highlight
    if rpr.find(_COLOR_TAG) is None:
        fg = inherited.fg
    if rpr.find(_SHD_TAG) is None:
        bg = inherited.bg
    if rpr.find(_HIGHLIGHT_TAG) is None:
        highlight = inherited.highlight
    return fg, bg, highlight


_P_TAG = f"{{{W_NS}}}p"
_PSTYLE_PATH = f"{{{W_NS}}}pPr/{{{W_NS}}}pStyle"
_TABLE_OF_FIGURES_STYLE = "TableofFigures"


def _table_of_contents_runs(tree: ET.Element, toc_styles: Collection[str]) -> set[int]:
    """The ids of the runs in the entries of a table of contents, figures or tables.

    Their text carries the Hyperlink style, but a table of contents is not drawn
    as a list of links, so the style's color is not taken over there. The ids
    themselves count too: ``docx_references_post_process`` writes ``TOC1``
    whether or not the template defines that style.
    """
    runs: set[int] = set()
    for paragraph in tree.iter(_P_TAG):
        pstyle = paragraph.find(_PSTYLE_PATH)
        style_id = pstyle.get(_VAL_ATTR, "") if pstyle is not None else ""
        if style_id in toc_styles or (style_id.startswith("TOC") and style_id[3:].isdigit()) or style_id == _TABLE_OF_FIGURES_STYLE:
            runs.update(id(run) for run in paragraph.iter(_R_TAG))
    return runs


def rewrite_part(xml_bytes: bytes, style_colors: Mapping[str, StyleColors] | None = None, toc_styles: Collection[str] = ()) -> tuple[bytes, dict[str, _StyleSpec]]:
    """Rewrite one body part. Returns (new_bytes, styles_used).

    ``style_colors`` maps a character style id to the colors it draws with
    (see :func:`character_style_colors`); a run referencing such a style takes
    every color it does not set directly from it. The entries of a table of
    contents, in ``toc_styles`` (see :func:`table_of_contents_styles`), do not.
    """
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping")
        return xml_bytes, {}

    styles_used: dict[str, _StyleSpec] = {}
    toc_runs = _table_of_contents_runs(tree, toc_styles) if style_colors else set()

    # iter() walks every <w:r> descendant regardless of depth — runs can be
    # nested inside tables, structured-document-tags, content controls, etc.,
    # so a top-level findall would miss most of them.
    for run in tree.iter(_R_TAG):
        rpr = run.find(_RPR_TAG)
        if rpr is None:
            # No run-properties element means no direct color formatting.
            continue
        fg, bg, highlight, size, size_cs = _extract_run_colors(rpr)
        if style_colors and id(run) not in toc_runs:
            fg, bg, highlight = _inherit_style_colors(rpr, style_colors, fg, bg, highlight)
        if not (fg or bg or highlight or size or size_cs):
            continue

        style_id = _style_id(fg, bg, highlight, size, size_cs)
        styles_used[style_id] = _StyleSpec(style_id, fg, bg, highlight, size, size_cs)
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
    stripped = stripped.removeprefix("#")
    if len(stripped) == HEX_COLOR_LENGTH and all(c in _HEX_DIGITS for c in stripped):
        return stripped.upper()
    return None


def build_style_element(spec: _StyleSpec) -> ET.Element:
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
    if spec.size_cs:
        ET.SubElement(rpr, f"{{{W_NS}}}szCs", {f"{{{W_NS}}}val": spec.size_cs})
    return style
