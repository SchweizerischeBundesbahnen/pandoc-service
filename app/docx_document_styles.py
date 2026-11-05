"""The styles of a DOCX, for its LaTeX/PDF rendering.

pandoc's DOCX reader drops the formatting of the styles: a PDF made from a
DOCX set its text in Latin Modern 10 pt, justified and hyphenated, with 6 pt
between paragraphs, and its title and headings in the sizes of LaTeX. Word
sets the same document in the font of its styles, Aptos 12 pt for pandoc's
own template, ragged right and unhyphenated, with the spacing of the Body
Text style, and its title and headings in the font, size and colour of their
styles.

This module reads the body text style (Body Text, then Normal, then the
document defaults), the Title and Heading 1 to 5 styles, the theme fonts and
the hyphenation setting, and hands them to pandoc: the font and size of the
body as LaTeX variables, the rest as metadata which
filters/docx_document_styles_to_latex.lua turns into LaTeX.

A font reaches LaTeX only when the container has a family of exactly that
name. Any other name, which a document may set to anything, keeps Latin
Modern: the name is never written into the LaTeX source. A colour reaches it
only as six hexadecimal digits.
"""

from __future__ import annotations

import functools
import io
import logging
import re
import subprocess
import zipfile
from typing import TYPE_CHECKING, Any

from .docx_ooxml import STYLES_PART, W_NS, parse_xml

if TYPE_CHECKING:
    from xml.etree.ElementTree import Element

logger = logging.getLogger(__name__)

DOCUMENT_PART = "word/document.xml"
THEME_PART = "word/theme/theme1.xml"
SETTINGS_PART = "word/settings.xml"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced

# The style pandoc gives body paragraphs, then the styles it inherits from.
BODY_STYLE = "BodyText"
# The styles of the title and of the headings LaTeX sets as \section to \subparagraph.
DISPLAY_STYLES = ("Title", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5")
# The sizes LaTeX's article class sets its text in. Word's 12 pt is the default of pandoc's template.
LATEX_FONT_SIZES = (10, 11, 12)
TWIPS_PER_POINT = 20
# Word lays a paragraph out justified only for these values of <w:jc>.
JUSTIFIED = frozenset({"both", "distribute"})
ALIGNMENTS = {"left": "left", "start": "left", "center": "center", "right": "right", "end": "right", "both": "left", "distribute": "left"}
HEX_COLOR = re.compile(r"[0-9A-Fa-f]{6}")
# The separators of a style in the metadata: a family which holds one is left out.
FIELD_SEPARATORS = frozenset(",:")
OFF_VALUES = frozenset({"false", "0", "off"})
# The single line of a font LaTeX sets when the font is not known: Word's comes from the font.
DEFAULT_LINE_RATIO = 1.2
# <w:spacing w:line> of a multiple counts in 240ths of a line.
LINE_UNITS = 240
# The colour scheme slot of each w:themeColor, as pandoc's and Word's own templates map them.
THEME_SLOTS = {
    "dark1": "dk1",
    "text1": "dk1",
    "light1": "lt1",
    "background1": "lt1",
    "dark2": "dk2",
    "text2": "dk2",
    "light2": "lt2",
    "background2": "lt2",
    **{f"accent{number}": f"accent{number}" for number in range(1, 7)},
    "hyperlink": "hlink",
    "followedHyperlink": "folHlink",
}
HEX_BYTE = re.compile(r"[0-9A-Fa-f]{2}")
# The colour filters/docx_document_styles_to_latex.lua defines as the colour of the Hyperlink style.
LINK_COLOR_NAME = "pdcLink"

_UNREADABLE_DOCX = (zipfile.BadZipFile, KeyError)


def document_style_options(docx_bytes: bytes) -> list[str]:
    """pandoc arguments which set the PDF in the styles Word sets the DOCX in."""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as package:
            parts = {name: package.read(name) for name in (STYLES_PART, THEME_PART, SETTINGS_PART, DOCUMENT_PART) if name in package.namelist()}
    except _UNREADABLE_DOCX:
        return []
    styles = parse_xml(parts[STYLES_PART]) if STYLES_PART in parts else None
    if styles is None:
        return []
    theme = parse_xml(parts[THEME_PART]) if THEME_PART in parts else None
    settings = parse_xml(parts[SETTINGS_PART]) if SETTINGS_PART in parts else None
    document = parse_xml(parts[DOCUMENT_PART]) if DOCUMENT_PART in parts else None

    chain = _style_chain(styles, BODY_STYLE)
    options: list[str] = []
    font = _font(chain, styles, theme)
    if font is not None:
        # Word kerns no text unless a style turns kerning on, which pandoc's template does not
        options += ["-V", f"mainfont={font}", "-V", "mainfontoptions=Kerning=Off"]
    size = _half_points(chain, styles)
    if size is not None:
        # The class sets the sizes of its footnotes and the like from the nearest size it knows
        nearest = min(LATEX_FONT_SIZES, key=lambda points: abs(points - size / 2))
        options += ["-V", f"fontsize={nearest}pt"]
        options += ["-M", f"docx-body-size={size / 2:g}", "-M", f"docx-body-baseline={_baseline(chain, styles, size / 2, font):.4g}"]
    skip = _paragraph_skip(chain, styles)
    options += ["-M", f"docx-body-paragraph-skip={skip:g}"]
    options += ["-M", f"docx-body-justified={'true' if _justified(chain, styles) else 'false'}"]
    options += ["-M", f"docx-body-hyphenated={'true' if _hyphenated(settings) else 'false'}"]
    link_chain = _style_chain(styles, "Hyperlink", fallback=None)
    link_color = _color(_first(link_chain, styles, "rPr", "color"), theme) if link_chain else None
    if link_color is not None:
        # The template sets hyperref up after the preamble of the filters, which defines pdcLink
        options += ["-M", f"docx-link-color={link_color}", "-V", "colorlinks=true"]
        options += [item for kind in ("link", "url", "file", "cite") for item in ("-V", f"{kind}color={LINK_COLOR_NAME}")]
    if document is not None and not _has_header_or_footer(document):
        options += ["-M", "docx-page-numbers=false"]
    for style_id in DISPLAY_STYLES:
        fields = _display_style(styles, theme, style_id, skip)
        if fields:
            options += ["-M", f"docx-style-{style_id}={fields}"]
    return options


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _styles_by_id(styles: Element) -> dict[str | None, Element]:
    return {style.get(_w("styleId")): style for style in styles.iter(_w("style"))}


def _style_chain(styles: Element, style_id: str, fallback: str | None = "Normal") -> list[Element]:
    """The style and the styles it is based on, nearest first. A loop in basedOn ends the chain."""
    by_id = _styles_by_id(styles)
    chain: list[Element] = []
    current = by_id.get(style_id)
    if current is None and fallback is not None:
        current = by_id.get(fallback)
    while current is not None and current not in chain:
        chain.append(current)
        based_on = current.find(_w("basedOn"))
        current = by_id.get(based_on.get(_w("val"))) if based_on is not None else None
    return chain


def _defaults(styles: Element, kind: str) -> Element | None:
    """The run (rPr) or paragraph (pPr) properties of the document defaults."""
    return styles.find(f"{_w('docDefaults')}/{_w(kind + 'Default')}/{_w(kind)}")


def _first(chain: list[Element], styles: Element, kind: str, child: str) -> Element | None:
    """The first element ``child`` of the ``kind`` properties along the chain, then in the defaults."""
    for style in chain:
        found = style.find(f"{_w(kind)}/{_w(child)}")
        if found is not None:
            return found
    defaults = _defaults(styles, kind)
    return defaults.find(_w(child)) if defaults is not None else None


def _owners(chain: list[Element], styles: Element, kind: str) -> list[Element]:
    """The ``kind`` properties of each style of the chain, nearest first, then of the defaults."""
    properties = [style.find(_w(kind)) for style in chain]
    return [owner for owner in [*properties, _defaults(styles, kind)] if owner is not None]


def _attribute(chain: list[Element], styles: Element, kind: str, child: str, attributes: tuple[str, ...]) -> tuple[str, str] | None:
    """The first of ``attributes`` of ``child`` which a style of the chain sets, with its name.

    A style may carry the element with other attributes only, <w:rFonts w:eastAsiaTheme> or
    <w:spacing w:line>, and inherit the one asked for.
    """
    for owner in _owners(chain, styles, kind):
        element = owner.find(_w(child))
        if element is None:
            continue
        for attribute in attributes:
            value = element.get(_w(attribute))
            if value is not None:
                return attribute, value
    return None


def _font(chain: list[Element], styles: Element, theme: Element | None) -> str | None:
    """The family of the text of the style, if the container has it."""
    found = _attribute(chain, styles, "rPr", "rFonts", ("ascii", "asciiTheme"))
    if found is None:
        return None
    attribute, value = found
    family = value if attribute == "ascii" else _theme_font(theme, value)
    if not family:
        return None
    if family not in installed_families():
        logger.info("docx_document_styles: the font %r is not installed, keeping the LaTeX font", family)
        return None
    return family


def _theme_font(theme: Element | None, slot: str | None) -> str | None:
    """The Latin family of a theme slot: minorHAnsi is the minor font, majorHAnsi the major one."""
    if theme is None or slot not in ("minorHAnsi", "majorHAnsi"):
        return None
    scheme = "minorFont" if slot == "minorHAnsi" else "majorFont"
    latin = theme.find(f".//{{{A_NS}}}{scheme}/{{{A_NS}}}latin")
    return latin.get("typeface") if latin is not None else None


@functools.cache
def installed_families() -> frozenset[str]:
    """The font families fontconfig knows, which tectonic finds its fonts through."""
    try:
        result = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True, check=True, timeout=30)  # noqa: S607 - fc-list of the image
    except OSError, subprocess.SubprocessError:
        logger.warning("docx_document_styles: fc-list failed, no font of a document is used")
        return frozenset()
    return frozenset(family.strip() for line in result.stdout.splitlines() for family in line.split(",") if family.strip())


def _half_points(chain: list[Element], styles: Element) -> int | None:
    size = _first(chain, styles, "rPr", "sz")
    try:
        return int(size.get(_w("val"), "")) if size is not None else None
    except ValueError:
        return None


def _paragraph_skip(chain: list[Element], styles: Element) -> float:
    """The space between two body paragraphs in points: the larger of after and before, as Word lays them out."""
    return max(_spacing_twips(chain, styles, side) for side in ("before", "after")) / TWIPS_PER_POINT


def _spacing_twips(chain: list[Element], styles: Element, side: str) -> int:
    """One side of <w:spacing>, from the nearest style which sets it: a style may set the line and inherit the space."""
    found = _attribute(chain, styles, "pPr", "spacing", (side,))
    try:
        return int(found[1]) if found is not None else 0
    except ValueError:
        return 0


def _justified(chain: list[Element], styles: Element) -> bool:
    alignment = _first(chain, styles, "pPr", "jc")
    return alignment is not None and alignment.get(_w("val")) in JUSTIFIED


def _hyphenated(settings: Element | None) -> bool:
    if settings is None:
        return False
    hyphenation = settings.find(_w("autoHyphenation"))
    return hyphenation is not None and hyphenation.get(_w("val"), "true") not in OFF_VALUES


def _toggle(chain: list[Element], styles: Element, child: str) -> bool:
    """A toggle property of the run, <w:b> or <w:i>, as the nearest style sets it."""
    element = _first(chain, styles, "rPr", child)
    return element is not None and element.get(_w("val"), "true") not in OFF_VALUES


def _display_style(styles: Element, theme: Element | None, style_id: str, skip: float) -> str:
    """The title or a heading style as ``name:value`` fields, empty when the document lacks the style.

    The space before and after is the whole space next to the heading: Word puts the larger of the
    after of one paragraph and the before of the next between them, and titlesec adds no paragraph
    skip around a heading.
    """
    chain = _style_chain(styles, style_id, fallback=None)
    if not chain:
        return ""
    fields: dict[str, str] = {}
    font = _font(chain, styles, theme)
    if font is not None and not FIELD_SEPARATORS & set(font):
        fields["font"] = font
    size = _half_points(chain, styles)
    if size is not None:
        fields["size"] = f"{size / 2:g}"
        fields["line"] = f"{_baseline(chain, styles, size / 2, font):.4g}"
    fields["bold"] = "true" if _toggle(chain, styles, "b") else "false"
    fields["italic"] = "true" if _toggle(chain, styles, "i") else "false"
    color = _color(_first(chain, styles, "rPr", "color"), theme)
    if color is not None:
        fields["color"] = color
    for side in ("before", "after"):
        fields[side] = f"{max(_spacing_twips(chain, styles, side) / TWIPS_PER_POINT, skip):g}"
    alignment = _first(chain, styles, "pPr", "jc")
    fields["align"] = ALIGNMENTS.get(alignment.get(_w("val"), "") if alignment is not None else "", "left")
    return ",".join(f"{name}:{value}" for name, value in fields.items())


@functools.cache
def line_ratio(family: str | None) -> float:
    """The single line of a family as a share of its size, as Word sets it: hhea ascent, descent and line gap.

    Word for Mac reads the hhea table, and so does Word for Windows for a font which asks for its typographic
    metrics, Aptos among them. The family is one fontconfig knows: see installed_families.
    """
    if family is None:
        return DEFAULT_LINE_RATIO
    try:
        result = subprocess.run(["fc-match", "-f", "%{file}", f"{family}:style=Regular"], capture_output=True, text=True, check=True, timeout=30)  # noqa: S603, S607 - fc-match of the image, given an installed family as one argument
        from fontTools.ttLib import TTFont  # noqa: PLC0415 - only a document in an installed font needs it

        font = TTFont(result.stdout.strip(), lazy=True, fontNumber=0)
        # fontTools makes the fields of a table at run time, which a type checker does not see
        hhea: Any = font["hhea"]
        head: Any = font["head"]
        return (hhea.ascent - hhea.descent + hhea.lineGap) / head.unitsPerEm
    except Exception:  # noqa: BLE001 - a font fontTools cannot read gets LaTeX's line
        logger.info("docx_document_styles: no line metrics for %r, using %s", family, DEFAULT_LINE_RATIO)
        return DEFAULT_LINE_RATIO


def _baseline(chain: list[Element], styles: Element, size: float, font: str | None) -> float:
    """The distance of two lines of the style in points: the single line of its font, times its line spacing."""
    single = size * line_ratio(font)
    found = _attribute(chain, styles, "pPr", "spacing", ("line",))
    rule = _attribute(chain, styles, "pPr", "spacing", ("lineRule",))
    try:
        line = int(found[1]) if found is not None else LINE_UNITS
    except ValueError:
        line = LINE_UNITS
    match rule[1] if rule is not None else "auto":
        case "exact":
            return line / TWIPS_PER_POINT
        case "atLeast":
            return max(single, line / TWIPS_PER_POINT)
        case _:
            return single * line / LINE_UNITS


def _theme_colors(theme: Element | None) -> dict[str, str]:
    """The colour scheme of the theme: each slot as six hex digits."""
    if theme is None:
        return {}
    scheme = theme.find(f".//{{{A_NS}}}clrScheme")
    colors: dict[str, str] = {}
    for slot in scheme if scheme is not None else []:
        name = slot.tag.removeprefix(f"{{{A_NS}}}")
        rgb = slot.find(f"{{{A_NS}}}srgbClr")
        system = slot.find(f"{{{A_NS}}}sysClr")
        value = rgb.get("val", "") if rgb is not None else system.get("lastClr", "") if system is not None else ""
        if HEX_COLOR.fullmatch(value):
            colors[name] = value.upper()
    return colors


def _color(element: Element | None, theme: Element | None) -> str | None:
    """The colour of a <w:color> as six hex digits: the theme colour when it names one, as Word reads it, else its value."""
    if element is None:
        return None
    slot = THEME_SLOTS.get(element.get(_w("themeColor"), ""))
    base = _theme_colors(theme).get(slot) if slot is not None else None
    if base is None:
        value = element.get(_w("val"), "")
        return value.upper() if HEX_COLOR.fullmatch(value) else None
    channels = [int(base[index : index + 2], 16) for index in (0, 2, 4)]
    shade, tint = element.get(_w("themeShade"), ""), element.get(_w("themeTint"), "")
    if HEX_BYTE.fullmatch(shade):
        channels = [channel * int(shade, 16) // 255 for channel in channels]
    if HEX_BYTE.fullmatch(tint):
        channels = [channel + (255 - channel) * (255 - int(tint, 16)) // 255 for channel in channels]
    return "".join(f"{channel:02X}" for channel in channels)


def _has_header_or_footer(document: Element) -> bool:
    """Whether a section of the document references a header or a footer, which Word prints the page number in."""
    return any(True for tag in ("headerReference", "footerReference") for _ in document.iter(_w(tag)))
