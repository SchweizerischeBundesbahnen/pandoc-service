"""The body text of a DOCX, for its LaTeX/PDF rendering.

pandoc's DOCX reader drops the font and the paragraph layout of the styles: a
PDF made from a DOCX set its text in Latin Modern 10 pt, justified and
hyphenated, with 6 pt between paragraphs. Word sets the same document in the
font of its styles, Aptos 12 pt for pandoc's own template, ragged right and
unhyphenated, with the spacing of the Body Text style.

This module reads the body text style (Body Text, then Normal, then the
document defaults), the theme fonts and the hyphenation setting, and hands
them to pandoc: the font and size as LaTeX variables, the paragraph layout as
metadata which filters/docx_body_style_to_latex.lua turns into LaTeX.

A font reaches LaTeX only when the container has a family of exactly that
name. Any other name, which a document may set to anything, keeps Latin
Modern: the name is never written into the LaTeX source.
"""

from __future__ import annotations

import functools
import io
import logging
import subprocess
import zipfile
from typing import TYPE_CHECKING

from .docx_ooxml import STYLES_PART, W_NS, parse_xml

if TYPE_CHECKING:
    from xml.etree.ElementTree import Element

logger = logging.getLogger(__name__)

THEME_PART = "word/theme/theme1.xml"
SETTINGS_PART = "word/settings.xml"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced

# The style pandoc gives body paragraphs, then the styles it inherits from.
BODY_STYLE = "BodyText"
# The sizes LaTeX's article class sets its text in. Word's 12 pt is the default of pandoc's template.
LATEX_FONT_SIZES = (10, 11, 12)
TWIPS_PER_POINT = 20
# Word lays a paragraph out justified only for these values of <w:jc>.
JUSTIFIED = frozenset({"both", "distribute"})

_UNREADABLE_DOCX = (zipfile.BadZipFile, KeyError)


def body_style_options(docx_bytes: bytes) -> list[str]:
    """pandoc arguments which set the body text of the PDF as Word sets the DOCX."""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as package:
            parts = {name: package.read(name) for name in (STYLES_PART, THEME_PART, SETTINGS_PART) if name in package.namelist()}
    except _UNREADABLE_DOCX:
        return []
    styles = parse_xml(parts[STYLES_PART]) if STYLES_PART in parts else None
    if styles is None:
        return []
    theme = parse_xml(parts[THEME_PART]) if THEME_PART in parts else None
    settings = parse_xml(parts[SETTINGS_PART]) if SETTINGS_PART in parts else None

    chain = _style_chain(styles, BODY_STYLE)
    options: list[str] = []
    font = _font(chain, styles, theme)
    if font is not None:
        options += ["-V", f"mainfont={font}"]
    size = _half_points(chain, styles)
    if size is not None:
        nearest = min(LATEX_FONT_SIZES, key=lambda points: abs(points - size / 2))
        options += ["-V", f"fontsize={nearest}pt"]
    options += ["-M", f"docx-body-paragraph-skip={_paragraph_skip(chain, styles):g}"]
    options += ["-M", f"docx-body-justified={'true' if _justified(chain, styles) else 'false'}"]
    options += ["-M", f"docx-body-hyphenated={'true' if _hyphenated(settings) else 'false'}"]
    return options


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _style_chain(styles: Element, style_id: str) -> list[Element]:
    """The style and the styles it is based on, nearest first. A loop in basedOn ends the chain."""
    by_id = {style.get(_w("styleId")): style for style in styles.iter(_w("style"))}
    chain: list[Element] = []
    current = by_id.get(style_id)
    if current is None:
        current = by_id.get("Normal")
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


def _font(chain: list[Element], styles: Element, theme: Element | None) -> str | None:
    """The family of the body text, if the container has it."""
    fonts = _first(chain, styles, "rPr", "rFonts")
    if fonts is None:
        return None
    family = fonts.get(_w("ascii"))
    if family is None:
        family = _theme_font(theme, fonts.get(_w("asciiTheme")))
    if not family:
        return None
    if family not in installed_families():
        logger.info("docx_body_style: the font %r of the body text is not installed, keeping the LaTeX font", family)
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
        logger.warning("docx_body_style: fc-list failed, no font of a document is used")
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
    defaults = _defaults(styles, "pPr")
    owners = [*(style.find(_w("pPr")) for style in chain), defaults]
    for owner in owners:
        spacing = owner.find(_w("spacing")) if owner is not None else None
        value = spacing.get(_w(side)) if spacing is not None else None
        if value is not None:
            try:
                return int(value)
            except ValueError:
                return 0
    return 0


def _justified(chain: list[Element], styles: Element) -> bool:
    alignment = _first(chain, styles, "pPr", "jc")
    return alignment is not None and alignment.get(_w("val")) in JUSTIFIED


def _hyphenated(settings: Element | None) -> bool:
    if settings is None:
        return False
    hyphenation = settings.find(_w("autoHyphenation"))
    return hyphenation is not None and hyphenation.get(_w("val"), "true") not in ("false", "0", "off")
