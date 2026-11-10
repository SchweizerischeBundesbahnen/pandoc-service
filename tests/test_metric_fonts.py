"""Tests for ``fonts/metric_fonts.py``."""

from __future__ import annotations

import io
import json
import zipfile
from typing import TYPE_CHECKING

import pytest
from fontTools.feaLib.builder import addOpenTypeFeaturesFromString
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTCollection, TTFont, newTable
from fontTools.ttLib.tables import otTables, ttProgram

from fonts import metric_fonts

if TYPE_CHECKING:
    from pathlib import Path


def _font(family: str, upm: int, advances: dict[str, int], ascent: int, descent: int) -> TTFont:
    """A font with a box glyph per character: 'a' and 'b', the space, and a composite 'c' made of 'a'."""
    builder = FontBuilder(upm, isTTF=True)
    order = [".notdef", "space", "a", "b", "c"]
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap({32: "space", 97: "a", 98: "b", 99: "c"})
    glyphs = {}
    for name in order:
        pen = TTGlyphPen(glyphs)
        if name in ("a", "b", ".notdef"):
            pen.moveTo((50, 0))
            pen.lineTo((50, 500))
            pen.lineTo((advances.get(name, 500) - 50, 500))
            pen.lineTo((advances.get(name, 500) - 50, 0))
            pen.closePath()
        elif name == "c":
            pen.addComponent("a", (1, 0, 0, 1, 0, 0))
        glyphs[name] = pen.glyph()
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (advances.get(name, 500), 50) for name in order})
    builder.setupHorizontalHeader(ascent=ascent, descent=-descent)
    builder.setupNameTable({"familyName": family, "styleName": "Regular"})
    builder.setupOS2(sTypoAscender=ascent, sTypoDescender=-descent, usWinAscent=ascent, usWinDescent=descent, sxHeight=ascent // 2, sCapHeight=ascent * 3 // 4)
    builder.setupPost()
    font = builder.font
    buffer = io.BytesIO()
    font.save(buffer)
    buffer.seek(0)
    return TTFont(buffer)


@pytest.fixture
def word() -> TTFont:
    return _font("Word Font", 2048, {"space": 400, "a": 1000, "b": 1200, "c": 1000}, 1900, 600)


@pytest.fixture
def source() -> TTFont:
    return _font("Open Font", 1000, {"space": 250, "a": 520, "b": 560, "c": 520}, 800, 200)


def test_the_metrics_of_word_hold_only_the_characters_of_the_open_font(word: TTFont, source: TTFont):
    del source["cmap"].tables[0].cmap[98]
    metrics = metric_fonts.word_metrics(word, source)
    assert metrics["advances"] == {"32": 400, "97": 1000, "99": 1000}
    assert metrics["unitsPerEm"] == 2048
    assert metrics["vertical"]["hhea"] == {"ascent": 1900, "descent": -600, "lineGap": 0}


def test_each_character_gets_the_advance_of_word_in_the_units_of_the_open_font(word: TTFont, source: TTFont):
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    advances = {name: source["hmtx"].metrics[name][0] for name in ("space", "a", "b", "c")}
    assert advances == {"space": round(400 * 1000 / 2048), "a": round(1000 * 1000 / 2048), "b": round(1200 * 1000 / 2048), "c": round(1000 * 1000 / 2048)}


def test_a_glyph_is_scaled_into_its_new_advance(word: TTFont, source: TTFont):
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    glyph = source["glyf"]["b"]
    scale = round(1200 * 1000 / 2048) / 560
    assert glyph.xMin == round(50 * scale)
    assert glyph.xMax == round(510 * scale)


def test_a_composite_glyph_keeps_the_outline_of_its_component_before_the_change(word: TTFont, source: TTFont):
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    assert not source["glyf"]["c"].isComposite()
    assert source["glyf"]["c"].xMax == source["glyf"]["a"].xMax


def test_the_line_of_the_open_font_is_the_line_of_word(word: TTFont, source: TTFont):
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    assert source["hhea"].ascent == round(1900 * 1000 / 2048)
    assert source["hhea"].descent == round(-600 * 1000 / 2048)
    assert source["OS/2"].usWinAscent == round(1900 * 1000 / 2048)
    # fontspec scales another font to the height of x of the main font
    assert source["OS/2"].sxHeight == round(950 * 1000 / 2048)


def test_a_slant_leans_the_outlines_and_sets_the_italic_angle(word: TTFont, source: TTFont):
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source), slant=12)
    assert source["post"].italicAngle == -12
    glyph = source["glyf"]["a"]
    assert glyph.xMax > round(470 * round(1000 * 1000 / 2048) / 520)


def test_the_hinting_goes_with_the_old_outlines(word: TTFont, source: TTFont):
    prep = newTable("prep")
    prep.program = ttProgram.Program()
    prep.program.fromBytecode(b"")
    source["prep"] = prep
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    assert "prep" not in source


def test_the_advances_no_longer_depend_on_hinting(word: TTFont, source: TTFont):
    """Bit 4 of the flags says instructions may alter the advance widths; bit 3, integer sizes, stays as it was."""
    source["head"].flags |= (1 << 4) | (1 << 3)
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    assert not source["head"].flags & (1 << 4)
    assert source["head"].flags & (1 << 3)


def test_the_font_takes_the_name_of_its_family_alone(source: TTFont):
    metric_fonts.rename(source, "Metric Sans", "Bold Italic")
    names = source["name"]
    assert (names.getDebugName(1), names.getDebugName(2), names.getDebugName(4), names.getDebugName(6)) == ("Metric Sans", "Bold Italic", "Metric Sans Bold Italic", "MetricSans-BoldItalic")
    assert "Open Font" in names.getDebugName(10)


def test_build_writes_each_style_from_its_source(tmp_path: Path, word: TTFont, source: TTFont, monkeypatch: pytest.MonkeyPatch):
    sources = tmp_path / "sources"
    sources.mkdir()
    source.save(sources / "open.ttf")
    with zipfile.ZipFile(sources / "release.zip", "w") as release:
        buffer = io.BytesIO()
        source.save(buffer)
        release.writestr("inner.ttf", buffer.getvalue())
    styles = (metric_fonts.Style("Metric Test", "Regular", "word.ttf", "open.ttf"), metric_fonts.Style("Metric Test", "Italic", "word.ttf", "release.zip/inner.ttf", slant=12))
    monkeypatch.setattr(metric_fonts, "STYLES", styles)
    words = tmp_path / "word"
    words.mkdir()
    word.save(words / "word.ttf")
    metrics = tmp_path / "metrics.json"
    assert metric_fonts.main(["extract", str(words), str(sources), str(metrics)]) == 0
    assert set(json.loads(metrics.read_text())) == {"MetricTest-Regular.ttf", "MetricTest-Italic.ttf"}
    assert metric_fonts.main(["build", str(sources), str(metrics), str(tmp_path / "out")]) == 0
    italic = TTFont(tmp_path / "out" / "MetricTest-Italic.ttf")
    assert italic["name"].getDebugName(4) == "Metric Test Italic"
    assert italic["post"].italicAngle == -12


def test_without_a_command_it_shows_how_to_call_it():
    assert metric_fonts.main([]) == 2


def test_the_scripts_of_the_open_font_are_the_scripts_of_word(word: TTFont, source: TTFont):
    """Word sizes and shifts a superscript and a subscript by the OS/2 table of its font."""
    os2 = word["OS/2"]
    os2.ySuperscriptYSize, os2.ySuperscriptYOffset, os2.ySubscriptYSize, os2.ySubscriptYOffset = 1229, 717, 1229, 154
    metric_fonts.apply_metrics(source, metric_fonts.word_metrics(word, source))
    scripts = tuple(getattr(source["OS/2"], name) for name in metric_fonts.SCRIPTS)
    assert scripts == tuple(round(value * 1000 / 2048) for value in (1229, 717, 1229, 154))


def _math_font(family: str, upm: int, shapes: dict[str, tuple[int, int, int]], larger: list[str], italic: int) -> TTFont:
    """A math font of box glyphs, each (advance, bottom, top): 'a' with a script form 'a.s1', and a bracket with its larger forms."""
    order = [".notdef", "a", "a.s1", "paren", *larger]
    builder = FontBuilder(upm, isTTF=True)
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap({97: "a", 40: "paren"})
    glyphs = {}
    for name in order:
        advance, bottom, top = shapes.get(name, (500, 0, 500))
        pen = TTGlyphPen(None)
        pen.moveTo((50, bottom))
        pen.lineTo((50, top))
        pen.lineTo((advance - 50, top))
        pen.lineTo((advance - 50, bottom))
        pen.closePath()
        glyphs[name] = pen.glyph()
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (shapes.get(name, (500, 0, 500))[0], 50) for name in order})
    builder.setupHorizontalHeader(ascent=upm * 3 // 4, descent=-upm // 4)
    builder.setupNameTable({"familyName": family, "styleName": "Regular"})
    builder.setupOS2(sTypoAscender=upm * 3 // 4, sTypoDescender=-upm // 4, usWinAscent=upm * 3 // 4, usWinDescent=upm // 4)
    builder.setupPost()
    addOpenTypeFeaturesFromString(builder.font, "feature ssty { sub a from [a.s1]; } ssty;")
    table = otTables.MATH()
    table.Version = 0x00010000
    constants = otTables.MathConstants()
    for converter in constants.getConverters():
        if "Percent" in converter.name:
            setattr(constants, converter.name, 70)
        elif converter.name in ("DelimitedSubFormulaMinHeight", "DisplayOperatorMinHeight"):
            setattr(constants, converter.name, upm)
        else:
            record = otTables.MathValueRecord()
            record.Value, record.DeviceTable = upm // 4, None
            setattr(constants, converter.name, record)
    table.MathConstants = constants
    info = otTables.MathGlyphInfo()
    info.MathItalicsCorrectionInfo = otTables.MathItalicsCorrectionInfo()
    info.MathItalicsCorrectionInfo.Coverage = otTables.Coverage()
    info.MathItalicsCorrectionInfo.Coverage.glyphs = ["a"]
    record = otTables.MathValueRecord()
    record.Value, record.DeviceTable = italic, None
    info.MathItalicsCorrectionInfo.ItalicsCorrection = [record]
    info.MathItalicsCorrectionInfo.ItalicsCorrectionCount = 1
    info.MathTopAccentAttachment = info.ExtendedShapeCoverage = info.MathKernInfo = None
    table.MathGlyphInfo = info
    variants = otTables.MathVariants()
    variants.MinConnectorOverlap = 0
    variants.VertGlyphCoverage = otTables.Coverage()
    variants.VertGlyphCoverage.glyphs = ["paren"]
    construction = otTables.MathGlyphConstruction()
    construction.GlyphAssembly = None
    construction.MathGlyphVariantRecord = []
    for name in ("paren", *larger):
        variant = otTables.MathGlyphVariantRecord()
        variant.VariantGlyph, variant.AdvanceMeasurement = name, shapes[name][2] - shapes[name][1]
        construction.MathGlyphVariantRecord.append(variant)
    construction.VariantCount = len(construction.MathGlyphVariantRecord)
    variants.VertGlyphConstruction = [construction]
    variants.VertGlyphCount = 1
    variants.HorizGlyphCoverage, variants.HorizGlyphConstruction, variants.HorizGlyphCount = None, [], 0
    table.MathVariants = variants
    builder.font["MATH"] = newTable("MATH")
    builder.font["MATH"].table = table
    buffer = io.BytesIO()
    builder.font.save(buffer)
    buffer.seek(0)
    return TTFont(buffer)


@pytest.fixture
def word_math() -> TTFont:
    shapes = {"a": (1000, -20, 900), "a.s1": (800, -16, 700), "paren": (600, -500, 1500), "paren.v1": (650, -800, 1800), "paren.v2": (700, -1400, 2400)}
    return _math_font("Word Math", 2048, shapes, ["paren.v1", "paren.v2"], italic=60)


@pytest.fixture
def source_math() -> TTFont:
    shapes = {"a": (520, -10, 460), "a.s1": (450, -10, 400), "paren": (300, -250, 750), "paren.v1": (330, -600, 1100)}
    return _math_font("Open Math", 1000, shapes, ["paren.v1"], italic=10)


def _built(word: TTFont, source: TTFont) -> TTFont:
    metrics = metric_fonts.word_metrics(word, source)
    metric_fonts.apply_metrics(source, metrics)
    metric_fonts.apply_math(source, metrics["math"], metrics["unitsPerEm"])
    buffer = io.BytesIO()
    source.save(buffer)
    buffer.seek(0)
    return TTFont(buffer)


def test_the_math_metrics_of_word_hold_its_constants_and_the_shapes_of_each_character(word_math: TTFont, source_math: TTFont):
    math = metric_fonts.word_metrics(word_math, source_math)["math"]
    assert math["constants"]["AxisHeight"] == 512
    assert math["constants"]["ScriptPercentScaleDown"] == 70
    assert math["glyphs"]["97"] == {"advance": 1000, "bottom": -20, "top": 900, "italic": 60, "scripts": [{"advance": 800, "bottom": -16, "top": 700}]}
    assert [shape["top"] for shape in math["glyphs"]["40"]["larger"]] == [1800, 2400]
    assert math["glyphs"]["40"]["larger"][1]["measure"] == 3800


def test_a_math_glyph_is_as_high_and_as_deep_as_the_glyph_of_word(word_math: TTFont, source_math: TTFont):
    """TeX sets a script, a radical and a fraction by the heights of the glyphs, which the outlines give."""
    glyph = _built(word_math, source_math)["glyf"]["a"]
    assert (glyph.yMin, glyph.yMax) == (round(-20 * 1000 / 2048), round(900 * 1000 / 2048))


def test_the_form_of_a_script_takes_the_width_and_the_height_of_word(word_math: TTFont, source_math: TTFont):
    font = _built(word_math, source_math)
    assert font["hmtx"].metrics["a.s1"][0] == round(800 * 1000 / 2048)
    assert font["glyf"]["a.s1"].yMax == round(700 * 1000 / 2048)


def test_each_larger_form_of_word_gets_a_glyph_of_its_own(word_math: TTFont, source_math: TTFont):
    """The open font has one larger bracket, Word two: TeX takes the first one large enough, so each size of Word is there."""
    font = _built(word_math, source_math)
    records = font["MATH"].table.MathVariants.VertGlyphConstruction[0].MathGlyphVariantRecord
    assert [record.VariantGlyph for record in records] == ["paren", "paren.metric1", "paren.metric2"]
    largest = font["glyf"]["paren.metric2"]
    assert (largest.yMin, largest.yMax) == (round(-1400 * 1000 / 2048), round(2400 * 1000 / 2048))
    assert records[2].AdvanceMeasurement == round(3800 * 1000 / 2048)


def test_the_math_font_takes_the_constants_and_the_italic_corrections_of_word(word_math: TTFont, source_math: TTFont):
    table = _built(word_math, source_math)["MATH"].table
    assert table.MathConstants.AxisHeight.Value == round(512 * 1000 / 2048)
    assert table.MathConstants.ScriptPercentScaleDown == 70
    italics = table.MathGlyphInfo.MathItalicsCorrectionInfo
    assert italics.ItalicsCorrection[italics.Coverage.glyphs.index("a")].Value == round(60 * 1000 / 2048)


def test_a_collection_gives_the_font_of_word_by_its_family(tmp_path: Path, word: TTFont, word_math: TTFont):
    collection = TTCollection()
    collection.fonts = [word, word_math]
    collection.save(tmp_path / "fonts.ttc")
    style = metric_fonts.Style("Metric Math", "Regular", "fonts.ttc#Word Math", "open.ttf")
    assert metric_fonts.open_word(tmp_path, style)["name"].getDebugName(1) == "Word Math"
