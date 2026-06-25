"""Tests for SvgProcessor (with a real Chromium browser) and the HTML SVG preprocessing hook.

The SVG-conversion tests drive a real ChromiumManager (started/stopped per test),
mirroring weasyprint-service. The pure dimension/parsing helpers need no browser.
The controller lifespan helpers (_start_chromium/_stop_chromium) are verified with a
manager-level mock, since they test the controller wiring rather than the browser.
"""

import base64
import os
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from bs4 import BeautifulSoup
from defusedxml import ElementTree as DET

from app import PandocController
from app.chromium_manager import ChromiumManager
from app.svg_processor import SvgProcessor


def _svg_data_url(svg: str) -> str:
    b64 = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{b64}"


# ---------------- Dimension parsing (no browser) ----------------


@pytest.mark.parametrize(
    "svg_content,dimension,expected",
    [
        ('<svg width="100px"></svg>', "width", ("100", "px")),
        ('<svg height="50"></svg>', "height", ("50", None)),
        ("<svg></svg>", "width", (None, None)),
        ('<svg width="10em"></svg>', "width", ("10", "em")),
    ],
)
def test_parse_svg_dimension(svg_content, dimension, expected):
    svg = DET.fromstring(svg_content)
    assert SvgProcessor().get_svg_dimension(svg, dimension) == expected


@pytest.mark.parametrize(
    "svg_content,expected",
    [
        ('<svg viewBox="0 0 800 600"></svg>', (800.0, 600.0)),
        ("<svg></svg>", (None, None)),
        ('<svg viewBox="0 0 800"></svg>', (None, None)),
        ('<svg viewBox="0,0,800,600"></svg>', (800.0, 600.0)),
        ('<svg viewBox="0 0 abc 600"></svg>', (None, None)),
    ],
)
def test_parse_viewbox(svg_content, expected):
    assert SvgProcessor().parse_viewbox(DET.fromstring(svg_content)) == expected


@pytest.mark.parametrize(
    "svg_content,expected_width,expected_height",
    [
        ('<svg height="200px" width="100px"></svg>', 100, 200),
        ('<svg viewBox="0 0 300 150"></svg>', 300, 150),
        ("<svg></svg>", None, None),
        ('<svg width="100px" viewBox="0 0 400 200"></svg>', 100, 200),
        ('<svg width="abc" height="xyz"></svg>', None, None),
    ],
)
def test_extract_svg_dimensions(svg_content, expected_width, expected_height):
    width, height, _ = SvgProcessor().extract_svg_dimensions_as_px(DET.fromstring(svg_content))
    assert (width, height) == (expected_width, expected_height)


@pytest.mark.parametrize(
    "svg_content,expected_error",
    [
        ('<svg width="100vw" height="100vh"></svg>', "vw units require a viewBox to be defined"),
        ('<svg width="100%" height="100%"></svg>', "% units require a viewBox to be defined"),
    ],
)
def test_extract_svg_dimensions_relative_units_error(svg_content, expected_error):
    with pytest.raises(ValueError, match=expected_error):
        SvgProcessor().extract_svg_dimensions_as_px(DET.fromstring(svg_content))


@pytest.mark.parametrize(
    "svg_content,expected_width,expected_height",
    [
        ('<svg width="100vw" height="100vh" viewBox="0 0 800 600"></svg>', 800, 600),
        ('<svg width="50%" height="25%" viewBox="0 0 800 600"></svg>', 400, 150),
    ],
)
def test_extract_svg_dimensions_relative_units(svg_content, expected_width, expected_height):
    processor = SvgProcessor()
    width, height, updated = processor.extract_svg_dimensions_as_px(DET.fromstring(svg_content))
    assert (width, height) == (expected_width, expected_height)
    content = processor.svg_to_string(updated)
    assert f'width="{width}px"' in content
    assert f'height="{height}px"' in content


@pytest.mark.parametrize(
    "content_type,content_base64,expected_content",
    [
        ("image/png", "123ABC==", None),
        ("image/svg+xml", "PHN2ZyBoZWlnaHQ9IjIwMHB4IiB3aWR0aD0iMTAwcHgiAA==", None),  # null byte
        ("image/svg+xml", "PHN2ZyBoZWlnaHQ9IjIwMHB4IiB3aWR0aD0iMTAwcHgi", None),  # no end tag
        ("image/svg+xml", "PHN2ZyBoZWlnaHQ9IjIwMHB4IiB3aWR0aD0iMTAwcHgiPjwvc3ZnPg==", '<svg height="200px" width="100px" />'),
    ],
)
def test_get_svg_content(content_type, content_base64, expected_content):
    processor = SvgProcessor()
    svg = processor.get_svg(content_type, content_base64)
    if expected_content is None:
        assert svg is None
    else:
        assert processor.svg_to_string(svg) == expected_content


def test_to_base64():
    assert SvgProcessor().to_base64(b"00000") == "MDAwMDA="
    assert SvgProcessor().to_base64("abcde") == "YWJjZGU="


def test_convert_to_px():
    processor = SvgProcessor()
    assert processor.convert_to_px("10", "px") == 10
    assert processor.convert_to_px("1", "mm") == 4
    assert processor.convert_to_px(None, "px") is None
    assert processor.convert_to_px("abc", "px") is None
    assert processor.convert_to_px("100", "vh") is None


def test_px_conversion_ratio():
    processor = SvgProcessor()
    assert processor.get_px_conversion_ratio("px") == 1
    assert processor.get_px_conversion_ratio("pt") == 4 / 3
    assert processor.get_px_conversion_ratio("in") == 96
    assert processor.get_px_conversion_ratio(None) == 1


def test_calculate_dimension():
    processor = SvgProcessor()
    assert processor.calculate_dimension("100", "px", None) == 100
    assert processor.calculate_dimension("75", "pt", None) == 100
    assert processor.calculate_dimension("100", "vw", 800.0) == 800
    assert processor.calculate_dimension(None, "px", None) is None
    with pytest.raises(ValueError, match="vw units require a viewBox"):
        processor.calculate_dimension("100", "vw", None)


def test_calculate_special_unit():
    processor = SvgProcessor()
    assert processor.calculate_special_unit("50", "%", 1000) == 500
    assert processor.calculate_special_unit("75", "pt", 1000) == 100
    with pytest.raises(ValueError, match="could not convert string to float"):
        processor.calculate_special_unit("abc", "px", 1000)


def test_replace_svg_size_attributes():
    processor = SvgProcessor()
    svg = DET.fromstring('<svg width="100" height="100"></svg>')
    result = processor.svg_to_string(processor.replace_svg_size_attributes(svg, 200, 300))
    assert 'width="200px"' in result
    assert 'height="300px"' in result


@pytest.mark.parametrize(
    "svg_input",
    [
        "<svg width='10' height='10'></svg>",
        "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'></svg>",
    ],
)
def test_ensure_mandatory_attributes(svg_input):
    processor = SvgProcessor()
    svg = processor.svg_from_string(svg_input)
    updated = processor.ensure_mandatory_attributes(svg)
    assert updated is svg
    assert processor.svg_to_string(updated).count('xmlns="http://www.w3.org/2000/svg"') == 1


def test_apply_img_dimensions_from_svg():
    processor = SvgProcessor()
    soup = BeautifulSoup('<img style="width: 500px; height: 300px; color: red;">', "html.parser")
    node = soup.find("img")
    svg = DET.fromstring('<svg width="100" height="200"></svg>')
    processor._apply_img_dimensions_from_svg(node, svg)
    assert node.get("width") == "100px"
    style = node.get("style")
    assert "width: 100px" in style
    assert "height:" not in style.lower()
    assert "color: red" in style


def test_svg_from_string_invalid_returns_none():
    assert SvgProcessor().svg_from_string("<svg>") is None


# ---------------- process_svg integration (real Chromium) ----------------


@pytest.mark.asyncio
async def test_process_svg_converts_base64_svg_to_png():
    manager = ChromiumManager()
    await manager.start()
    try:
        processor = SvgProcessor(chromium_manager=manager)
        svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100" fill="blue"/></svg>'
        soup = BeautifulSoup(f'<img src="{_svg_data_url(svg)}">', "html.parser")

        result = await processor.process_svg(soup)

        assert result.find("img")["src"].startswith("data:image/png;base64,")
    finally:
        await manager.stop()


@pytest.mark.asyncio
async def test_process_svg_inline_svg_without_manager_becomes_base64_svg():
    processor = SvgProcessor()  # no chromium_manager -> falls back to base64 SVG
    soup = BeautifulSoup('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"></svg>', "html.parser")

    result = await processor.process_svg(soup)

    assert "data:image/svg+xml;base64," in str(result)


@pytest.mark.asyncio
async def test_process_svg_passes_through_non_svg():
    manager = ChromiumManager()
    await manager.start()
    try:
        processor = SvgProcessor(chromium_manager=manager)
        png_url = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        soup = BeautifulSoup(f'<img src="{png_url}">', "html.parser")

        result = await processor.process_svg(soup)

        # Non-SVG image must be passed through untouched.
        assert result.find("img")["src"] == png_url
    finally:
        await manager.stop()


# ---------------- PandocController.preprocess_html_svgs ----------------


def test_is_svg_conversion_enabled_default_and_disabled():
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}):
        assert PandocController.is_svg_conversion_enabled() is True
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "false"}):
        assert PandocController.is_svg_conversion_enabled() is False
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "0"}):
        assert PandocController.is_svg_conversion_enabled() is False


@pytest.mark.asyncio
async def test_preprocess_html_svgs_disabled_returns_unchanged():
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "false"}):
        source = "<html><body><img src='data:image/svg+xml;base64,Zm9v'></body></html>"
        assert await PandocController.preprocess_html_svgs(source) == source


@pytest.mark.asyncio
async def test_preprocess_html_svgs_no_svg_returns_unchanged():
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}):
        source = "<html><body><p>no svg here</p></body></html>"
        assert await PandocController.preprocess_html_svgs(source) == source


@pytest.mark.asyncio
async def test_preprocess_html_svgs_browser_not_running_returns_unchanged():
    manager = ChromiumManager()  # constructed but not started -> is_running() is False
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
        source = "<html><body><img src='data:image/svg+xml;base64,Zm9v'></body></html>"
        assert await PandocController.preprocess_html_svgs(source) == source


@pytest.mark.asyncio
async def test_preprocess_html_svgs_renders_svg_to_png_str_input():
    manager = ChromiumManager()
    await manager.start()
    try:
        svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100"/></svg>'
        source = f'<html><body><img src="{_svg_data_url(svg)}"></body></html>'
        with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
            result = await PandocController.preprocess_html_svgs(source)
        assert isinstance(result, str)
        assert "data:image/png;base64," in result
    finally:
        await manager.stop()


@pytest.mark.asyncio
async def test_preprocess_html_svgs_renders_svg_to_png_bytes_input():
    manager = ChromiumManager()
    await manager.start()
    try:
        svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100"/></svg>'
        source = f'<html><body><img src="{_svg_data_url(svg)}"></body></html>'.encode()
        with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
            result = await PandocController.preprocess_html_svgs(source)
        assert isinstance(result, bytes)
        assert b"data:image/png;base64," in result
    finally:
        await manager.stop()


@pytest.mark.asyncio
async def test_preprocess_html_svgs_with_scale_factor_renders_png():
    manager = ChromiumManager()
    await manager.start()
    try:
        svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100"/></svg>'
        source = f'<html><body><img src="{_svg_data_url(svg)}"></body></html>'
        with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
            result = await PandocController.preprocess_html_svgs(source, scale_factor=2.0)
        assert "data:image/png;base64," in result
    finally:
        await manager.stop()


# ---------------- controller lifespan helpers (manager-level mock) ----------------


def _mock_manager(running=True, start_error=None):
    manager = MagicMock()
    manager.start = AsyncMock(side_effect=start_error)
    manager.stop = AsyncMock()
    manager.is_running = MagicMock(return_value=running)
    manager.get_version = MagicMock(return_value="131.0.0.0")
    manager.health_check = MagicMock(return_value=running)
    return manager


@pytest.mark.asyncio
async def test_start_chromium_disabled_does_not_start():
    manager = _mock_manager()
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "false"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
        await PandocController._start_chromium()
    manager.start.assert_not_called()


@pytest.mark.asyncio
async def test_start_chromium_enabled_starts():
    manager = _mock_manager()
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
        await PandocController._start_chromium()
    manager.start.assert_awaited_once()


@pytest.mark.asyncio
async def test_start_chromium_swallows_start_error():
    manager = _mock_manager(start_error=RuntimeError("no browser"))
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
        # Must not raise: SVG rasterization is best effort.
        await PandocController._start_chromium()


@pytest.mark.asyncio
async def test_stop_chromium_stops_when_running():
    manager = _mock_manager(running=True)
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
        await PandocController._stop_chromium()
    manager.stop.assert_awaited_once()


@pytest.mark.asyncio
async def test_stop_chromium_noop_when_not_running():
    manager = _mock_manager(running=False)
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=manager):
        await PandocController._stop_chromium()
    manager.stop.assert_not_called()


def test_get_chromium_health_states():
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "false"}):
        assert PandocController.get_chromium_health() == "disabled"
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=_mock_manager(running=True)):
        assert PandocController.get_chromium_health() == "available"
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}), patch.object(PandocController, "get_chromium_manager", return_value=_mock_manager(running=False)):
        assert PandocController.get_chromium_health() == "stopped"
