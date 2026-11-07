"""Tests for ``fonts/metric_fonts.py``."""

from __future__ import annotations

import io
import json
import zipfile
from typing import TYPE_CHECKING

import pytest
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables import ttProgram

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
