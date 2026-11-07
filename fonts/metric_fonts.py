"""Make open fonts which set text as wide and as high as the fonts of Word they stand for.

Word sets a document in fonts the image may not ship: Aptos, Segoe UI, Georgia. A PDF set in another
font breaks its lines and pages elsewhere. Each Metric family here takes the outlines of an open font
and the metrics of a font of Word: the advance width of every character, the space, and the vertical
metrics Word sets a line with. app/docx_document_styles.py maps the name of the font of Word to it.

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
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

SELAWIK = "Selawik_Release.zip"
# The tables of the vertical metrics of a font, with the fields Word reads, and the heights of x and the capitals,
# which fontspec scales another font to
VERTICAL = {"hhea": ("ascent", "descent", "lineGap"), "OS/2": ("sTypoAscender", "sTypoDescender", "sTypoLineGap", "usWinAscent", "usWinDescent", "sxHeight", "sCapHeight")}
USE_TYPO_METRICS = 1 << 7
ITALIC_BITS = 0b1100001  # italic, bold and regular of OS/2 fsSelection
# The tables of hinting, which no longer fit the changed outlines
HINTING = ("fpgm", "prep", "cvt ", "gasp", "hdmx", "LTSH", "VDMX")
NAME_IDS = (1, 2, 3, 4, 6, 16, 17, 21, 22, 25)
WINDOWS_ENGLISH = (3, 1, 0x409)


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
    }


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
        result[file_name(style)] = word_metrics(TTFont(word_fonts / style.word), open_source(sources, style))
    target.write_text(json.dumps(result, indent=1, sort_keys=True) + "\n")


def build(sources: Path, metrics_file: Path, out: Path) -> None:
    metrics = json.loads(metrics_file.read_text())
    out.mkdir(parents=True, exist_ok=True)
    for style in STYLES:
        font = open_source(sources, style)
        apply_metrics(font, metrics[file_name(style)], style.slant)
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
