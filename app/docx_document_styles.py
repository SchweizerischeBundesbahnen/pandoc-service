"""The styles of a DOCX, for its LaTeX/PDF rendering.

pandoc's DOCX reader drops the formatting of the styles: a PDF made from a
DOCX set its text in Latin Modern 10 pt, justified and hyphenated, with 6 pt
between paragraphs, and its title and headings in the sizes of LaTeX. Word
sets the same document in the font of its styles, Aptos 12 pt for pandoc's
own template, ragged right and unhyphenated, with the spacing of the Body
Text style, and its title and headings in the font, size and colour of their
styles.

This module reads the body text style (Body Text, then Normal, then the
document defaults), the Title and Heading 1 to 9 styles, the theme fonts and
the hyphenation setting, and hands them to pandoc: the font and size of the
body as LaTeX variables, the rest as metadata which
filters/docx/docx_document_styles_to_latex.lua turns into LaTeX.

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

from .docx_numbers import whole_number
from .docx_ooxml import STYLES_PART, W_NS, parse_xml

if TYPE_CHECKING:
    from xml.etree.ElementTree import Element

logger = logging.getLogger(__name__)

DOCUMENT_PART = "word/document.xml"
THEME_PART = "word/theme/theme1.xml"
SETTINGS_PART = "word/settings.xml"
NUMBERING_PART = "word/numbering.xml"
# The LaTeX counter command of each number format of Word a heading may carry.
NUMBER_FORMATS = {"decimal": "arabic", "lowerLetter": "alph", "upperLetter": "Alph", "lowerRoman": "roman", "upperRoman": "Roman"}
# The text of a number: its level placeholders, digits and the punctuation around them.
_NUMBER_TEXT = re.compile(r"^(?:%[1-9]|[0-9.\-() ])*$")
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced

# The font Word sets the equations (OMML) of every document in.
MATH_FONT = "Cambria Math"
# The open families the image has in place of fonts of Word, each as wide and as high as the font it stands for:
# fonts/metric_fonts.py makes the Metric families, the others are metric compatible as they come
SUBSTITUTES = {
    "aptos": "Metric Sans",
    "aptos display": "Metric Sans Display",
    "segoe ui": "Metric UI",
    "georgia": "Metric Serif",
    "calibri": "Carlito",
    "arial": "Liberation Sans",
    "times new roman": "Liberation Serif",
    "courier new": "Liberation Mono",
}
# The hhea ascent of the fonts Word sets its bullets in, which no open font stands for: Symbol and Wingdings
BULLET_ASCENTS = {"symbol": 1.005, "wingdings": 0.899}
# The style pandoc gives body paragraphs, then the styles it inherits from.
BODY_STYLE = "BodyText"
# The style pandoc gives the paragraphs of a table cell and of a tight list.
CELL_STYLE = "Compact"
# The styles of the title, of the nine levels of headings, and of captions.
DISPLAY_STYLES = ("Title", *(f"Heading{level}" for level in range(1, 10)), "Caption")
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

_UNREADABLE_DOCX = (zipfile.BadZipFile, KeyError)


def _body_line(chain: list[Element], styles: Element, size: float, font: str | None) -> list[str]:
    """The size and the line of the body text, with where a line starts and ends when the font is known."""
    options = ["-M", f"docx-body-size={size:g}", "-M", f"docx-body-baseline={_baseline(chain, styles, size, font):.4g}"]
    ascent = ascent_ratio(font)
    if ascent is not None:
        # The balloon of a comment starts at the top of its line, the line to it at the bottom
        descent = line_ratio(font) - ascent
        options += ["-M", f"docx-body-ascent={size * ascent:.4g}", "-M", f"docx-body-descent={size * descent:.4g}"]
        if gap_ratio(font) > 0:
            # A picture lowered below the descent of the font takes the line gap below it as well
            options += ["-M", f"docx-body-line-gap={size * gap_ratio(font):.4g}"]
    return options


def _paragraph_spacing(styles: Element, document: Element | None) -> str:
    """The space before and after the paragraph styles of the body, as ``id:name:before:after`` joined by ``;``.

    Word puts the larger of the space after one paragraph and before the next between them, LaTeX the
    paragraph skip. Normal, the style of a paragraph without one, is always there. A heading, the title
    and a caption have spaces of their own.
    """
    used = {"Normal"}
    if document is not None:
        used |= {style.get(_w("val"), "") for style in document.iter(_w("pStyle"))}
    entries = []
    for style_id in sorted(used):
        chain = _style_chain(styles, style_id, fallback=None)
        if not chain or style_id in DISPLAY_STYLES or not re.fullmatch(r"[A-Za-z0-9]+", style_id):
            continue
        name_element = chain[0].find(_w("name"))
        # The reader names a style without a name by its id
        name = name_element.get(_w("val"), style_id) if name_element is not None else style_id
        if not re.fullmatch(r"[A-Za-z0-9 ]+", name):
            continue
        before, after = (_spacing_twips(chain, styles, side) / TWIPS_PER_POINT for side in ("before", "after"))
        entries.append(f"{style_id}:{name}:{before:g}:{after:g}")
    return ";".join(entries)


def _cell_options(styles: Element) -> list[str]:
    """The spaces of the paragraphs of a table cell and of a list item, which Word sets in Compact or Normal."""
    options: list[str] = []
    cell_chain = _style_chain(styles, CELL_STYLE, fallback=None)
    if cell_chain:
        # A cell holds its paragraph with the space before and after it
        cell_space = sum(_spacing_twips(cell_chain, styles, side) for side in ("before", "after")) / TWIPS_PER_POINT
        options += ["-M", f"docx-table-cell-space={cell_space:g}"]
        # The items of a tight list are paragraphs of Compact: Word puts the larger of its spaces between them
        item_space = max(_spacing_twips(cell_chain, styles, side) for side in ("before", "after")) / TWIPS_PER_POINT
        options += ["-M", f"docx-list-item-space={item_space:g}"]
    # A cell paragraph without a style is in Normal, with the space of Normal
    plain_chain = _style_chain(styles, "Normal", fallback=None)
    plain_space = sum(_spacing_twips(plain_chain, styles, side) for side in ("before", "after")) / TWIPS_PER_POINT
    options += ["-M", f"docx-table-plain-cell-space={plain_space:g}"]
    return options


def _font_options(font: str | None) -> list[str]:
    """The font of the body, and the math font when it is installed."""
    options: list[str] = []
    if font is not None:
        # Word kerns no text unless a style turns kerning on, which pandoc's template does not, never
        # shrinks a space to fit one more word on the line, as TeX does,
        # nor turns a straight quote of the text into a curly one, nor sets the ligatures of TeX and the font
        options += ["-V", f"mainfont={font}", "-V", "mainfontoptions=Kerning=Off", "-V", "mainfontoptions=WordSpace={1,1,0}"]
        options += ["-V", "mainfontoptions=Ligatures={TeXOff,NoCommon}"]
    if MATH_FONT in installed_families():
        options += ["-V", f"mathfont={MATH_FONT}"]
    return options


def _size_options(chain: list[Element], styles: Element, font: str | None, numbering: Element | None, document: Element | None) -> list[str]:
    """The size and the line of the body text, and the bullets of its lists, when the styles give a size."""
    size = _half_points(chain, styles)
    if size is None:
        return []
    # The class sets the sizes of its footnotes and the like from the nearest size it knows
    nearest = min(LATEX_FONT_SIZES, key=lambda points: abs(points - size / 2))
    options = ["-V", f"fontsize={nearest}pt", *_body_line(chain, styles, size / 2, font)]
    bullets = _bullet_ascents(numbering, document, size / 2)
    if any(bullets):
        options += ["-M", "docx-list-bullet-ascents=" + ",".join(f"{ascent:g}" for ascent in bullets)]
    return options


def _display_options(styles: Element, theme: Element | None, numbering: Element | None) -> list[str]:
    """The fields of the title, the headings and the caption, each style that has any."""
    options: list[str] = []
    for style_id in DISPLAY_STYLES:
        fields = _display_style(styles, theme, style_id, numbering)
        if fields:
            options += ["-M", f"docx-style-{style_id}={fields}"]
    return options


def document_style_options(docx_bytes: bytes) -> list[str]:
    """pandoc arguments which set the PDF in the styles Word sets the DOCX in."""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as package:
            parts = {name: package.read(name) for name in (STYLES_PART, THEME_PART, SETTINGS_PART, DOCUMENT_PART, NUMBERING_PART) if name in package.namelist()}
    except _UNREADABLE_DOCX:
        return []
    styles = parse_xml(parts[STYLES_PART]) if STYLES_PART in parts else None
    if styles is None:
        return []
    theme = parse_xml(parts[THEME_PART]) if THEME_PART in parts else None
    settings = parse_xml(parts[SETTINGS_PART]) if SETTINGS_PART in parts else None
    document = parse_xml(parts[DOCUMENT_PART]) if DOCUMENT_PART in parts else None
    numbering = parse_xml(parts[NUMBERING_PART]) if NUMBERING_PART in parts else None

    chain = _style_chain(styles, BODY_STYLE)
    font = _font(chain, styles, theme)
    options = _font_options(font) + _size_options(chain, styles, font, numbering, document)
    indents = _list_indents(numbering, document)
    if indents:
        options += ["-M", "docx-list-indents=" + ",".join(f"{left:g}:{hanging:g}" for left, hanging in indents)]
    skip = _paragraph_skip(chain, styles)
    options += ["-M", f"docx-body-paragraph-skip={skip:g}"]
    options += ["-M", f"docx-body-justified={'true' if _justified(chain, styles) else 'false'}"]
    options += ["-M", f"docx-body-hyphenated={'true' if _hyphenated(settings) else 'false'}"]
    cell_margin = _cell_margin(styles)
    if cell_margin is not None:
        options += ["-M", f"docx-table-cell-margin={cell_margin:g}"]
    options += _cell_options(styles)
    options += ["-M", f"docx-paragraph-spacing={_paragraph_spacing(styles, document)}"]
    if document is not None and not _has_header_or_footer(document):
        options += ["-M", "docx-page-numbers=false"]
    return options + _display_options(styles, theme, numbering)


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _styles_by_id(styles: Element) -> dict[str | None, Element]:
    return {style.get(_w("styleId")): style for style in styles.iter(_w("style"))}


def _builtin_style(styles: Element, style_id: str) -> Element | None:
    """The style of the id given, else the built-in style of that name: a localised Word names Heading1 "heading 1"."""
    found = _styles_by_id(styles).get(style_id)
    if found is not None:
        return found
    name = re.sub(r"(?<=[a-z])(?=[A-Z0-9])", " ", style_id).lower()
    for style in styles.iter(_w("style")):
        style_name = style.find(_w("name"))
        if style_name is not None and (style_name.get(_w("val")) or "").lower() == name:
            return style
    return None


def _style_chain(styles: Element, style_id: str, fallback: str | None = "Normal") -> list[Element]:
    """The style and the styles it is based on, nearest first. A loop in basedOn ends the chain."""
    by_id = _styles_by_id(styles)
    chain: list[Element] = []
    current = _builtin_style(styles, style_id)
    if current is None and fallback is not None:
        current = _builtin_style(styles, fallback)
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
    installed = installed_family(family)
    if installed is None:
        logger.info("docx_document_styles: the font %r is not installed, keeping the LaTeX font", family)
    return installed


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


def installed_family(name: str | None) -> str | None:
    """The installed family of a name, matched as fontconfig matches it, regardless of case; None without one.

    A run of Polarion names its font as CSS does, in lower case: "segoe ui". A font of Word the container
    lacks gets its substitute, see SUBSTITUTES.
    """
    if not name:
        return None
    wanted = name.strip().casefold()
    families = sorted(installed_families())
    found = next((family for family in families if family.casefold() == wanted), None)
    substitute = SUBSTITUTES.get(wanted)
    if found is None and substitute in families:
        return substitute
    return found


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


def _display_style(styles: Element, theme: Element | None, style_id: str, numbering: Element | None = None) -> str:
    """The title or a heading style as ``name:value`` fields, empty when the document lacks the style.

    The space before and after is the style's own: Word puts the larger of the after of one paragraph
    and the before of the next between them, which filters/docx/docx_document_styles_to_latex.lua works
    out from the neighbours of each heading.
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
        ascent = ascent_ratio(font)
        if ascent is not None:
            fields["ascent"] = f"{size / 2 * ascent:.4g}"
    fields["bold"] = "true" if _toggle(chain, styles, "b") else "false"
    fields["italic"] = "true" if _toggle(chain, styles, "i") else "false"
    color = word_color(_first(chain, styles, "rPr", "color"), theme)
    if color is not None:
        fields["color"] = color
    for side in ("before", "after"):
        fields[side] = f"{_spacing_twips(chain, styles, side) / TWIPS_PER_POINT:g}"
    alignment = _first(chain, styles, "pPr", "jc")
    fields["align"] = ALIGNMENTS.get(alignment.get(_w("val"), "") if alignment is not None else "", "left")
    if style_id.startswith("Heading"):
        fields.update(_heading_number(chain, styles, numbering))
    return ",".join(f"{name}:{value}" for name, value in fields.items())


def _number_levels(numbering: Element, num_id: str) -> dict[str, Element]:
    """The levels of the abstract numbering a numbering instance refers to, by their index."""
    abstract_id = next(
        (ref.get(_w("val")) for num in numbering.iter(_w("num")) if num.get(_w("numId")) == num_id and (ref := num.find(_w("abstractNumId"))) is not None),
        None,
    )
    abstract = next((item for item in numbering.iter(_w("abstractNum")) if item.get(_w("abstractNumId")) == abstract_id), None)
    if abstract is None:
        return {}
    return {level.get(_w("ilvl"), ""): level for level in abstract.iter(_w("lvl"))}


def _bullet_levels(numbering: Element | None, document: Element | None) -> dict[str, Element]:
    """The levels of the numbering of the first bulleted list of the document, by their index."""
    if numbering is None or document is None:
        return {}
    for num_id in document.iter(_w("numId")):
        levels = _number_levels(numbering, num_id.get(_w("val"), ""))
        format_ = levels["0"].find(_w("numFmt")) if "0" in levels else None
        if format_ is not None and format_.get(_w("val")) == "bullet":
            return levels
    return {}


def _bullet_ascents(numbering: Element | None, document: Element | None, size: float) -> list[float]:
    """How far above the baseline the bullet of each of the first four levels of a bulleted list reaches, in points.

    Word sets a line as high as the highest font in it, the bullet's among them. The levels are those of
    the first bulleted list of the document; a level whose font is not installed gets the ascent of
    BULLET_ASCENTS, else 0.
    """
    levels = _bullet_levels(numbering, document)
    ascents = []
    for index in range(4):
        fonts = levels[str(index)].find(f"{_w('rPr')}/{_w('rFonts')}") if str(index) in levels else None
        family = fonts.get(_w("ascii")) if fonts is not None else None
        ratio = ascent_ratio(installed_family(family))
        if ratio is None and family is not None:
            ratio = BULLET_ASCENTS.get(family.strip().casefold())
        ascents.append(round(size * ratio, 2) if ratio is not None else 0)
    return ascents


def _list_indents(numbering: Element | None, document: Element | None) -> list[tuple[float, float]]:
    """The indent of each of the first four levels of the first list of the document, in points.

    Each level as its left indent past the level above it and its hanging indent: Word sets the label at the
    start of the hanging indent and the text at the left indent. Empty when a level has no indent of its own.
    """
    if numbering is None or document is None:
        return []
    num_id = next((element.get(_w("val"), "") for element in document.iter(_w("numId"))), None)
    levels = _number_levels(numbering, num_id) if num_id is not None else {}
    indents: list[tuple[float, float]] = []
    outer = 0
    for index in range(4):
        indent = levels[str(index)].find(f"{_w('pPr')}/{_w('ind')}") if str(index) in levels else None
        if indent is None:
            return []
        left = whole_number(indent.get(_w("left"), indent.get(_w("start"))))
        hanging = whole_number(indent.get(_w("hanging"), "0"))
        if left is None or hanging is None or left <= outer:
            return []
        indents.append(((left - outer) / TWIPS_PER_POINT, hanging / TWIPS_PER_POINT))
        outer = left
    return indents


def _heading_level(chain: list[Element], styles: Element, numbering: Element | None) -> tuple[dict[str, Element], Element] | None:
    """The levels of the numbering of a heading style and the level it takes, or None without numbering."""
    num_pr = _first(chain, styles, "pPr", "numPr")
    num_id = num_pr.find(_w("numId")) if num_pr is not None else None
    if num_pr is None or num_id is None or numbering is None:
        return None
    level_index = num_pr.find(_w("ilvl"))
    levels = _number_levels(numbering, num_id.get(_w("val"), ""))
    level = levels.get(level_index.get(_w("val"), "0") if level_index is not None else "0")
    return (levels, level) if level is not None else None


def _number_text(text: str, levels: dict[str, Element]) -> str | None:
    """The text of a number with each level as ``{level:format}``, or None for a format LaTeX has no counter for."""
    parts = []
    for match in re.finditer(r"%([1-9])|[^%]+", text):
        if match.group(1) is None:
            parts.append(match.group(0))
            continue
        number_format = levels.get(str(int(match.group(1)) - 1))
        number_format = number_format.find(_w("numFmt")) if number_format is not None else None
        latex = NUMBER_FORMATS.get(number_format.get(_w("val"), "") if number_format is not None else "")
        if latex is None:
            return None
        parts.append(f"{{{match.group(1)}:{latex}}}")
    return "".join(parts)


def _heading_number(chain: list[Element], styles: Element, numbering: Element | None) -> dict[str, str]:
    """The number the numbering of a heading style gives it: its text, each level in its format, and its hanging indent.

    The text names the level of each part as ``{level:format}``, the format one of LaTeX's counter commands.
    Empty for a style without numbering, or with a format or text this cannot set.
    """
    found = _heading_level(chain, styles, numbering)
    if found is None:
        return {}
    levels, level = found
    text = level.find(_w("lvlText"))
    if text is None or not _NUMBER_TEXT.match(text.get(_w("val"), "")):
        return {}
    number = _number_text(text.get(_w("val"), ""), levels)
    if number is None:
        return {}
    indent = level.find(f"{_w('pPr')}/{_w('ind')}")
    hanging = whole_number(indent.get(_w("hanging"), "0") if indent is not None else "0") or 0
    return {"number": number, "numberhang": f"{hanging / TWIPS_PER_POINT:g}"}


def _hhea(family: str) -> tuple[float, float, float] | None:
    """The ascent, the single line and the line gap of a family as shares of its size, from its hhea table; None when unreadable."""
    try:
        # fc-match of the image, given an installed family as one argument
        result = subprocess.run(["fc-match", "-f", "%{file}", f"{family}:style=Regular"], capture_output=True, text=True, check=True, timeout=30)  # noqa: S603, S607
        from fontTools.ttLib import TTFont  # noqa: PLC0415 - only a document in an installed font needs it

        font = TTFont(result.stdout.strip(), lazy=True, fontNumber=0)
        # fontTools makes the fields of a table at run time, which a type checker does not see
        hhea: Any = font["hhea"]
        head: Any = font["head"]
        return hhea.ascent / head.unitsPerEm, (hhea.ascent - hhea.descent + hhea.lineGap) / head.unitsPerEm, hhea.lineGap / head.unitsPerEm
    except Exception:  # noqa: BLE001 - a font fontTools cannot read gets LaTeX's line
        logger.info("docx_document_styles: no line metrics for %r, using %s", family, DEFAULT_LINE_RATIO)
        return None


@functools.cache
def line_ratio(family: str | None) -> float:
    """The single line of a family as a share of its size, as Word sets it: hhea ascent, descent and line gap.

    Word for Mac reads the hhea table, and so does Word for Windows for a font which asks for its typographic
    metrics, Aptos among them. The family is one fontconfig knows: see installed_families.
    """
    metrics = _hhea(family) if family is not None else None
    return metrics[1] if metrics is not None else DEFAULT_LINE_RATIO


@functools.cache
def ascent_ratio(family: str | None) -> float | None:
    """The hhea ascent of a family as a share of its size: Word sets the first line of a page that far down."""
    metrics = _hhea(family) if family is not None else None
    return metrics[0] if metrics is not None else None


@functools.cache
def gap_ratio(family: str | None) -> float:
    """The hhea line gap of a family as a share of its size: Word adds it below the deepest thing of a line."""
    metrics = _hhea(family) if family is not None else None
    return metrics[2] if metrics is not None else 0


def _baseline(chain: list[Element], styles: Element, size: float, font: str | None) -> float:
    """The distance of two lines of the style in points: the single line of its font, times its line spacing."""
    single = size * line_ratio(font)
    found = _attribute(chain, styles, "pPr", "spacing", ("line",))
    rule = _attribute(chain, styles, "pPr", "spacing", ("lineRule",))
    try:
        line = int(found[1]) if found is not None else LINE_UNITS
    except ValueError:
        line = LINE_UNITS
    kind = rule[1] if rule is not None else "auto"
    if kind == "exact":
        return line / TWIPS_PER_POINT
    if kind == "atLeast":
        return max(single, line / TWIPS_PER_POINT)
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
        value = ""
        if rgb is not None:
            value = rgb.get("val", "")
        elif system is not None:
            value = system.get("lastClr", "")
        if HEX_COLOR.fullmatch(value):
            colors[name] = value.upper()
    return colors


def word_color(element: Element | None, theme: Element | None) -> str | None:
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


def _cell_margin(styles: Element) -> float | None:
    """The left and right margin of a cell of the default table style, in points, when both are the same.

    pandoc writes its tables in the Table style, which Word lays out with 5.4 pt on each side of a cell
    where LaTeX puts \\tabcolsep, 6 pt.
    """
    table_styles = [style for style in styles.iter(_w("style")) if style.get(_w("type")) == "table"]
    default = next((style for style in table_styles if style.get(_w("default")) in ("1", "true", "on")), None)
    margins = default.find(f"{_w('tblPr')}/{_w('tblCellMar')}") if default is not None else None
    if margins is None:
        return None
    sides = [margins.find(_w(side)) for side in ("left", "right")]
    widths = {_twips_of(side) for side in sides}
    width = widths.pop() if len(widths) == 1 else None
    return width / TWIPS_PER_POINT if width is not None else None


def _twips_of(element: Element | None) -> int | None:
    """The width of a margin in twips: dxa only, the unit Word writes them in."""
    if element is None or element.get(_w("type"), "dxa") != "dxa":
        return None
    try:
        return int(element.get(_w("w"), ""))
    except ValueError:
        return None
