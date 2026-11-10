"""Make open fonts which set text as wide and as high as the fonts of Word they stand for.

Word sets a document in fonts the image may not ship: Aptos, Segoe UI, Georgia. A PDF set in another
font breaks its lines and pages elsewhere. Each Metric family here takes the outlines of an open font
and the metrics of a font of Word: the advance width of every character, the space, and the vertical
metrics Word sets a line with. app/docx_document_styles.py maps the name of the font of Word to it.

Metric Math stands for Cambria Math, which Word sets an equation in. An equation is laid out by the heights of
its glyphs as well, by their larger forms and by the constants of the MATH table: Metric Math takes those of
Cambria Math, as numbers, and draws each glyph with the outline of STIX Two Math fitted into them.

The open fonts are under the SIL Open Font License, which allows a modified font under another name.
The metrics of Word's fonts are numbers in fonts/word_metrics.json; no outline of them is in it.

Two commands:

    python fonts/metric_fonts.py extract WORD_FONTS SOURCES fonts/word_metrics.json
        reads the metrics from the fonts of Word, for the characters the open fonts have (once, by hand)
    python fonts/metric_fonts.py build SOURCES fonts/word_metrics.json OUT
        writes the Metric families into OUT (in the image build)

SOURCES holds the open fonts as the Dockerfile downloads them.
"""

from __future__ import annotations

import copy
import io
import json
import math
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTCollection, TTFont
from fontTools.ttLib.tables import otTables
from fontTools.varLib.instancer import instantiateVariableFont

SELAWIK = "Selawik_Release.zip"
# The tables of the vertical metrics of a font, with the fields Word reads, the heights of x and the capitals,
# which fontspec scales another font to, and the size and the shift of the scripts, which Word sets them by
SCRIPTS = ("ySuperscriptYSize", "ySuperscriptYOffset", "ySubscriptYSize", "ySubscriptYOffset")
VERTICAL = {"hhea": ("ascent", "descent", "lineGap"), "OS/2": ("sTypoAscender", "sTypoDescender", "sTypoLineGap", "usWinAscent", "usWinDescent", "sxHeight", "sCapHeight", *SCRIPTS)}
USE_TYPO_METRICS = 1 << 7
ITALIC_BITS = 0b1100001  # italic, bold and regular of OS/2 fsSelection
# The tables of hinting, which no longer fit the changed outlines
HINTING = ("fpgm", "prep", "cvt ", "gasp", "hdmx", "LTSH", "VDMX")
NAME_IDS = (1, 2, 3, 4, 6, 16, 17, 21, 22, 25)
WINDOWS_ENGLISH = (3, 1, 0x409)
# The tables of OpenType, whose classes fontTools makes at run time, which a type checker does not see
OT: Any = otTables


@dataclass(frozen=True)
class Style:
    """One font of a Metric family: the font of Word it measures and the open font it draws."""

    family: str
    style: str
    word: str
    source: str
    axes: dict[str, float] = field(default_factory=dict)
    # Degrees the outlines lean to the right: an italic the open font lacks
    slant: float = 0


# The width axis makes each open font about as wide as the font of Word, so the advance of a character changes little
STYLES = (
    Style("Metric Sans", "Regular", "Aptos.ttf", "Archivo[wdth,wght].ttf", {"wght": 400, "wdth": 96.88}),
    Style("Metric Sans", "Bold", "Aptos-Bold.ttf", "Archivo[wdth,wght].ttf", {"wght": 700, "wdth": 94.78}),
    Style("Metric Sans", "Italic", "Aptos-Italic.ttf", "Archivo-Italic[wdth,wght].ttf", {"wght": 400, "wdth": 96.37}),
    Style("Metric Sans", "Bold Italic", "Aptos-Bold-Italic.ttf", "Archivo-Italic[wdth,wght].ttf", {"wght": 700, "wdth": 94.43}),
    Style("Metric Sans Display", "Regular", "AptosDisplay-Regular.ttf", "Archivo[wdth,wght].ttf", {"wght": 400, "wdth": 89.82}),
    Style("Metric Sans Display", "Bold", "AptosDisplay-Bold.ttf", "Archivo[wdth,wght].ttf", {"wght": 700, "wdth": 87.96}),
    Style("Metric UI", "Regular", "SegoeUI-Regular.ttf", f"{SELAWIK}/selawk.ttf"),
    Style("Metric UI", "Bold", "SegoeUI-Bold.ttf", f"{SELAWIK}/selawkb.ttf"),
    Style("Metric UI", "Italic", "SegoeUI-Italic.ttf", f"{SELAWIK}/selawk.ttf", slant=12),
    Style("Metric UI", "Bold Italic", "SegoeUI-Bold-Italic.ttf", f"{SELAWIK}/selawkb.ttf", slant=12),
    Style("Metric Serif", "Regular", "Georgia.ttf", "Gelasio[wght].ttf", {"wght": 400}),
    Style("Metric Serif", "Bold", "Georgia Bold.ttf", "Gelasio[wght].ttf", {"wght": 700}),
    Style("Metric Serif", "Italic", "Georgia Italic.ttf", "Gelasio-Italic[wght].ttf", {"wght": 400}),
    Style("Metric Serif", "Bold Italic", "Georgia Bold Italic.ttf", "Gelasio-Italic[wght].ttf", {"wght": 700}),
    # A collection names the font it holds after the #
    Style("Metric Math", "Regular", "Cambria.ttc#Cambria Math", "STIXTwoMath-Regular.ttf"),
)


def _table(font: TTFont, tag: str) -> Any:
    """A table of the font: fontTools makes its fields at run time, which a type checker does not see."""
    return font[tag]


def _cmap(font: TTFont) -> dict[int, str]:
    return font.getBestCmap() or {}


def file_name(style: Style) -> str:
    return f"{style.family}-{style.style}".replace(" ", "") + ".ttf"


def open_source(sources: Path, style: Style) -> TTFont:
    """The open font of a style at its axes. A path inside the Selawik archive is read from it."""
    archive, _, member = style.source.partition("/")
    if member:
        with zipfile.ZipFile(sources / archive) as release:
            font = TTFont(io.BytesIO(release.read(member)))
    else:
        font = TTFont(sources / style.source)
    if "fvar" in font:
        font = instantiateVariableFont(font, style.axes, updateFontNames=False)
    return font


def open_word(word_fonts: Path, style: Style) -> TTFont:
    """The font of Word of a style. A collection holds it under the family after the #."""
    path, _, family = style.word.partition("#")
    if not family:
        return TTFont(word_fonts / path)
    return next(font for font in TTCollection(word_fonts / path).fonts if font["name"].getDebugName(1) == family)


def word_metrics(word: TTFont, source: TTFont) -> dict[str, Any]:
    """The metrics of a font of Word, for the characters the open font has as well."""
    characters = _cmap(source)
    advances = _table(word, "hmtx").metrics
    return {
        "unitsPerEm": _table(word, "head").unitsPerEm,
        "vertical": {table: {name: getattr(_table(word, table), name) for name in names if hasattr(_table(word, table), name)} for table, names in VERTICAL.items()},
        "useTypoMetrics": bool(_table(word, "OS/2").fsSelection & USE_TYPO_METRICS),
        "weight": _table(word, "OS/2").usWeightClass,
        "fsSelection": _table(word, "OS/2").fsSelection & ITALIC_BITS,
        "macStyle": _table(word, "head").macStyle,
        "advances": {str(code): advances[glyph][0] for code, glyph in sorted(_cmap(word).items()) if code in characters},
    } | ({"math": math_metrics(word, source)} if "MATH" in word and "MATH" in source else {})


def _redraw(font: TTFont, glyph_name: str, transform: tuple[float, float, float, float, float, float]) -> Any:
    """The glyph drawn through a transformation, its components resolved."""
    recording = DecomposingRecordingPen(font.getGlyphSet())
    font.getGlyphSet()[glyph_name].draw(recording)
    pen = TTGlyphPen(None)
    recording.replay(TransformPen(pen, transform))
    return pen.glyph()


def apply_metrics(font: TTFont, metrics: dict[str, Any], slant: float = 0) -> None:
    """Give the font the metrics of the font of Word: each character as wide, the line as high, the outlines leaning by SLANT."""
    scale = _table(font, "head").unitsPerEm / metrics["unitsPerEm"]
    for table, values in metrics["vertical"].items():
        for name, value in values.items():
            setattr(_table(font, table), name, round(value * scale))
    os2 = _table(font, "OS/2")
    os2.fsSelection = (os2.fsSelection & ~(ITALIC_BITS | USE_TYPO_METRICS)) | metrics["fsSelection"] | (USE_TYPO_METRICS if metrics["useTypoMetrics"] else 0)
    os2.version = max(os2.version, 4)
    os2.usWeightClass = metrics["weight"]
    _table(font, "head").macStyle = metrics["macStyle"]
    shear = math.tan(math.radians(slant))
    if slant:
        _table(font, "post").italicAngle = -slant
    glyf, hmtx = _table(font, "glyf"), _table(font, "hmtx")
    # Every glyph is drawn anew from the old outlines before any of them changes, since composites refer to others
    targets: dict[str, int] = {}
    for code, glyph_name in _cmap(font).items():
        advance = metrics["advances"].get(str(code))
        if advance is not None and glyph_name not in targets:
            targets[glyph_name] = round(advance * scale)
    drawn = {}
    for glyph_name in font.getGlyphOrder():
        old = hmtx.metrics[glyph_name][0]
        width = targets.get(glyph_name, old)
        # Each glyph keeps its place in its advance: it is as much wider or narrower as its advance is
        drawn[glyph_name] = (_redraw(font, glyph_name, (width / old if old else 1, 0, shear, 1, 0, 0)), width)
    for glyph_name, (glyph, width) in drawn.items():
        glyf[glyph_name] = glyph
        glyph.recalcBounds(glyf)
        hmtx.metrics[glyph_name] = (width, getattr(glyph, "xMin", 0))
    for table in HINTING:
        if table in font:
            del font[table]
    _table(font, "head").flags &= ~(1 << 4)  # no hinting changes the advances anymore


def _script_forms(font: TTFont) -> dict[str, list[str]]:
    """Each glyph with its forms for the scripts and the scripts of scripts, from the ssty lookups."""
    forms: dict[str, list[str]] = {}
    if "GSUB" not in font:
        return forms
    gsub = _table(font, "GSUB").table
    for record in gsub.FeatureList.FeatureRecord:
        if record.FeatureTag != "ssty":
            continue
        for index in record.Feature.LookupListIndex:
            for wrapped in gsub.LookupList.Lookup[index].SubTable:
                subtable = getattr(wrapped, "ExtSubTable", wrapped)
                for glyph, alternates in getattr(subtable, "alternates", {}).items():
                    forms.setdefault(glyph, list(alternates))
                for glyph, alternate in getattr(subtable, "mapping", {}).items():
                    forms.setdefault(glyph, [alternate])
    return forms


def _larger_forms(font: TTFont) -> dict[str, Any]:
    """Each glyph with the construction of its larger forms, from the vertical variants of the MATH table."""
    variants = _table(font, "MATH").table.MathVariants
    if variants.VertGlyphCoverage is None:
        return {}
    return dict(zip(variants.VertGlyphCoverage.glyphs, variants.VertGlyphConstruction, strict=True))


def _covered(info: Any, coverage: str, values: str) -> dict[str, Any]:
    """A table of the MATH glyph info as a dictionary from the glyph to its record."""
    if info is None or getattr(info, coverage) is None:
        return {}
    return dict(zip(getattr(info, coverage).glyphs, getattr(info, values), strict=True))


def _shape(font: TTFont, glyph: str) -> dict[str, Any]:
    """The advance, the height and the depth of a glyph, and its italic correction where it has one."""
    glyf_glyph = _table(font, "glyf")[glyph]
    italics = _covered(_table(font, "MATH").table.MathGlyphInfo.MathItalicsCorrectionInfo, "Coverage", "ItalicsCorrection")
    shape = {"advance": _table(font, "hmtx").metrics[glyph][0], "bottom": getattr(glyf_glyph, "yMin", 0), "top": getattr(glyf_glyph, "yMax", 0)}
    return shape | ({"italic": italics[glyph].Value} if glyph in italics else {})


def math_metrics(word: TTFont, source: TTFont) -> dict[str, Any]:
    """What Word lays an equation out by, for the characters the open font has as well.

    The constants of the MATH table, and for each character the height and the depth of its glyph, its italic
    correction and where an accent sits on it, and the same of its forms for scripts and of its larger forms.
    """
    table = _table(word, "MATH").table
    constants = {}
    for converter in table.MathConstants.getConverters():
        value = getattr(table.MathConstants, converter.name)
        constants[converter.name] = getattr(value, "Value", value)
    accents = _covered(table.MathGlyphInfo.MathTopAccentAttachment, "TopAccentCoverage", "TopAccentAttachment")
    scripts, larger = _script_forms(word), _larger_forms(word)
    characters = _cmap(source)
    glyphs = {}
    for code, glyph in sorted(_cmap(word).items()):
        if code not in characters:
            continue
        entry = _shape(word, glyph) | ({"accent": accents[glyph].Value} if glyph in accents else {})
        if glyph in scripts:
            entry["scripts"] = [_shape(word, form) for form in scripts[glyph]]
        if glyph in larger:
            records = larger[glyph].MathGlyphVariantRecord[1:]
            entry["larger"] = [_shape(word, record.VariantGlyph) | {"measure": record.AdvanceMeasurement} for record in records]
        glyphs[str(code)] = entry
    return {"constants": constants, "glyphs": glyphs}


class _Fit(DecomposingRecordingPen):
    """Records a glyph with its outline stretched to a height above and a depth below the baseline, and to a width."""

    def __init__(self, glyph_set: Any, width: float, up: float, down: float) -> None:
        super().__init__(glyph_set)
        self.scale = (width, up, down)

    def _point(self, point: tuple[float, float]) -> tuple[float, float]:
        width, up, down = self.scale
        return (point[0] * width, point[1] * (up if point[1] > 0 else down))

    def glyph(self) -> Any:
        pen = TTGlyphPen(None)
        for operator, points in self.value:
            getattr(pen, operator)(*[self._point(point) if point is not None else None for point in points])
        return pen.glyph()


def _fitted(font: TTFont, glyph: str, shape: dict[str, Any], scale: float) -> tuple[Any, int]:
    """The glyph drawn as wide, as high and as deep as the shape of Word, in the units of the font."""
    old = _table(font, "glyf")[glyph]
    old.recalcBounds(_table(font, "glyf"))
    top, bottom = getattr(old, "yMax", 0), getattr(old, "yMin", 0)
    up = shape["top"] * scale / top if top > 0 and shape["top"] > 0 else 1
    down = shape["bottom"] * scale / bottom if bottom < 0 and shape["bottom"] < 0 else 1
    advance = _table(font, "hmtx").metrics[glyph][0]
    target = round(shape["advance"] * scale)
    pen = _Fit(font.getGlyphSet(), target / advance if advance else 1, up, down)
    font.getGlyphSet()[glyph].draw(pen)
    return pen.glyph(), target


def _set_covered(font: TTFont, info: Any, coverage: str, values: str, records: dict[str, Any]) -> None:
    """Write a table of the MATH table back from a dictionary, in the order of the glyphs, as its coverage wants."""
    glyphs = sorted(records, key=font.getGlyphID)
    if getattr(info, coverage) is None:
        setattr(info, coverage, OT.Coverage())
    getattr(info, coverage).glyphs = glyphs
    setattr(info, values, [records[glyph] for glyph in glyphs])
    count = {"ItalicsCorrection": "ItalicsCorrectionCount", "TopAccentAttachment": "TopAccentAttachmentCount", "VertGlyphConstruction": "VertGlyphCount"}[values]
    setattr(info, count, len(glyphs))


def _set_constants(table: Any, constants: dict[str, int], scale: float) -> None:
    """The constants of the MATH table of Word in the units of the font; a percentage stays as it is."""
    for name, value in constants.items():
        target = getattr(table.MathConstants, name)
        if hasattr(target, "Value"):
            target.Value = round(value * scale)
        else:
            setattr(table.MathConstants, name, value if "Percent" in name else round(value * scale))


def _targets(font: TTFont, metrics: dict[str, Any], scale: float) -> tuple[dict[str, dict[str, Any]], dict[str, tuple[str, dict[str, Any]]]]:
    """The shape of Word of each glyph and of its forms for scripts, and the new glyphs of the larger forms of Word.

    A new glyph is drawn from the larger form of the open font nearest in size: the open font may have fewer of
    them, and TeX takes the first one which is large enough.
    """
    scripts, larger, cmap = _script_forms(font), _larger_forms(font), _cmap(font)
    shapes: dict[str, dict[str, Any]] = {}
    new: dict[str, tuple[str, dict[str, Any]]] = {}
    for code, entry in metrics["glyphs"].items():
        glyph = cmap.get(int(code))
        if glyph is None:
            continue
        shapes.setdefault(glyph, entry)
        for form, shape in zip(scripts.get(glyph, []), entry.get("scripts", []), strict=False):
            shapes.setdefault(form, shape)
        if glyph in larger and entry.get("larger"):
            own = [record.VariantGlyph for record in larger[glyph].MathGlyphVariantRecord]
            for number, shape in enumerate(entry["larger"], start=1):
                height = (shape["top"] - shape["bottom"]) * scale
                new[f"{glyph}.metric{number}"] = (min(own, key=lambda other, height=height: abs(_height(font, other) - height)), shape)
    return shapes, new


def _set_glyph_info(font: TTFont, shapes: dict[str, dict[str, Any]], scale: float) -> None:
    """The italic corrections of Word, and where an accent sits on a glyph of both fonts."""
    info = _table(font, "MATH").table.MathGlyphInfo
    italics = _covered(info.MathItalicsCorrectionInfo, "Coverage", "ItalicsCorrection")
    accents = _covered(info.MathTopAccentAttachment, "TopAccentCoverage", "TopAccentAttachment")
    for glyph, shape in shapes.items():
        if "italic" in shape:
            record = italics.setdefault(glyph, OT.MathValueRecord())
            record.Value = round(shape["italic"] * scale)
            record.DeviceTable = getattr(record, "DeviceTable", None)
        if "accent" in shape and glyph in accents:
            accents[glyph].Value = round(shape["accent"] * scale)
    if italics:
        if info.MathItalicsCorrectionInfo is None:
            info.MathItalicsCorrectionInfo = OT.MathItalicsCorrectionInfo()
            info.MathItalicsCorrectionInfo.Coverage = None
        _set_covered(font, info.MathItalicsCorrectionInfo, "Coverage", "ItalicsCorrection", italics)


def _set_larger_forms(font: TTFont, metrics: dict[str, Any], scale: float) -> None:
    """The larger forms of each glyph as Word has them: the new glyphs, each measured as in Word."""
    larger, cmap = _larger_forms(font), _cmap(font)
    for code, entry in metrics["glyphs"].items():
        glyph = cmap.get(int(code))
        if glyph not in larger or not entry.get("larger"):
            continue
        first = larger[glyph].MathGlyphVariantRecord[0]
        first.AdvanceMeasurement = round((entry["top"] - entry["bottom"]) * scale)
        records = [first]
        for number, shape in enumerate(entry["larger"], start=1):
            record = copy.deepcopy(first)
            record.VariantGlyph = f"{glyph}.metric{number}"
            record.AdvanceMeasurement = round(shape["measure"] * scale)
            records.append(record)
        larger[glyph].MathGlyphVariantRecord = records
        larger[glyph].VariantCount = len(records)


def apply_math(font: TTFont, metrics: dict[str, Any], units_per_em: int) -> None:
    """Fit each glyph of the open math font, its forms for scripts and its larger forms into the shapes of Word."""
    scale = _table(font, "head").unitsPerEm / units_per_em
    _set_constants(_table(font, "MATH").table, metrics["constants"], scale)
    shapes, new = _targets(font, metrics, scale)
    # The new glyphs are drawn from the old outlines, before any glyph changes
    drawn = {name: _fitted(font, source, shape, scale) for name, (source, shape) in new.items()}
    drawn |= {glyph: _fitted(font, glyph, shape, scale) for glyph, shape in shapes.items()}
    glyf, hmtx = _table(font, "glyf"), _table(font, "hmtx")
    font.setGlyphOrder(font.getGlyphOrder() + list(new))
    glyf.setGlyphOrder(font.getGlyphOrder())
    for glyph, (outline, advance) in drawn.items():
        glyf[glyph] = outline
        outline.recalcBounds(glyf)
        hmtx.metrics[glyph] = (advance, getattr(outline, "xMin", 0))
    _set_glyph_info(font, shapes | {name: shape for name, (_, shape) in new.items()}, scale)
    _set_larger_forms(font, metrics, scale)


def _height(font: TTFont, glyph: str) -> float:
    outline = _table(font, "glyf")[glyph]
    outline.recalcBounds(_table(font, "glyf"))
    return getattr(outline, "yMax", 0) - getattr(outline, "yMin", 0)


def rename(font: TTFont, family: str, style: str) -> None:
    """Name the font after its family alone: the license of the open font reserves its name for the font as it is."""
    names = _table(font, "name")
    original = names.getDebugName(4) or names.getDebugName(1) or "an open font"
    for name_id in NAME_IDS:
        names.removeNames(nameID=name_id)
    full = family if style == "Regular" else f"{family} {style}"
    postscript = f"{family}-{style}".replace(" ", "")
    for name_id, text in ((1, family), (2, style), (3, postscript), (4, full), (6, postscript)):
        names.setName(text, name_id, *WINDOWS_ENGLISH)
    names.setName(f"Modified from {original}: the widths and line metrics of a font of Word, for pandoc-service.", 10, *WINDOWS_ENGLISH)


def extract(word_fonts: Path, sources: Path, target: Path) -> None:
    result = {}
    for style in STYLES:
        result[file_name(style)] = word_metrics(open_word(word_fonts, style), open_source(sources, style))
    target.write_text(json.dumps(result, indent=1, sort_keys=True) + "\n")


def build(sources: Path, metrics_file: Path, out: Path) -> None:
    metrics = json.loads(metrics_file.read_text())
    out.mkdir(parents=True, exist_ok=True)
    for style in STYLES:
        font = open_source(sources, style)
        apply_metrics(font, metrics[file_name(style)], style.slant)
        if "math" in metrics[file_name(style)]:
            apply_math(font, metrics[file_name(style)]["math"], metrics[file_name(style)]["unitsPerEm"])
        rename(font, style.family, style.style)
        font.save(out / file_name(style))


def main(arguments: list[str]) -> int:
    match arguments:
        case ["extract", word_fonts, sources, target]:
            extract(Path(word_fonts), Path(sources), Path(target))
        case ["build", sources, metrics_file, out]:
            build(Path(sources), Path(metrics_file), Path(out))
        case _:
            print(__doc__, file=sys.stderr)  # noqa: T201 - a command line tool
            return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
