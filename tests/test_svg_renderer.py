import hashlib
import importlib.metadata
import io
import os
import shutil
from unittest.mock import patch

import anyio
import pytest
from PIL import Image
from weasysvg import convert

from app import svg_renderer

DIAGRAM = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="red"/></svg>'


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, "chromium"),
        ("", "chromium"),
        ("weasysvg", "weasysvg"),
        (" WeasySVG ", "weasysvg"),
        ("chromium", "chromium"),
        ("rsvg", "chromium"),
    ],
)
def test_get_svg_renderer(value, expected):
    environment = {} if value is None else {"SVG_RENDERER": value}
    with patch.dict(os.environ, environment, clear=True):
        assert svg_renderer.get_svg_renderer() == expected


def test_configure_installs_weasysvg_only_when_selected():
    with patch.object(svg_renderer.weasysvg, "install") as install:
        with patch.dict(os.environ, {"SVG_RENDERER": "chromium"}):
            assert svg_renderer.configure_svg_renderer() == "chromium"
        install.assert_not_called()
        with patch.dict(os.environ, {"SVG_RENDERER": "weasysvg"}):
            assert svg_renderer.configure_svg_renderer() == "weasysvg"
        install.assert_called_once_with()


def test_get_weasysvg_version():
    assert svg_renderer.get_weasysvg_version() == importlib.metadata.version("weasysvg")
    with patch.object(importlib.metadata, "version", side_effect=importlib.metadata.PackageNotFoundError):
        assert svg_renderer.get_weasysvg_version() is None


@pytest.mark.parametrize(("target", "kept"), [("pdf", True), ("latex", True), ("html", True), ("epub", True), ("docx", False), ("pptx", False), ("odt", False)])
def test_keeps_svg(target, kept):
    assert svg_renderer.keeps_svg(target) is kept


def test_pandoc_environment_puts_rsvg_convert_first():
    environment = svg_renderer.pandoc_environment({"PATH": "/usr/bin", "OTHER": "1"})
    assert environment["PATH"] == f"{svg_renderer.RSVG_CONVERT_DIRECTORY}{os.pathsep}/usr/bin"
    assert environment["OTHER"] == "1"
    assert svg_renderer.pandoc_environment({})["PATH"] == svg_renderer.RSVG_CONVERT_DIRECTORY
    with patch.dict(os.environ, {"PATH": "/bin"}):
        assert svg_renderer.pandoc_environment(None)["PATH"].endswith("/bin")


def png_size(png):
    with Image.open(io.BytesIO(png)) as image:
        return image.size, image.getpixel((image.width // 2, image.height // 2))


def test_render_draws_a_png_of_the_requested_size():
    png, flat = svg_renderer.render(DIAGRAM.encode(), 40, 20, 2.0, layer=False)
    size, center = png_size(png)
    assert size == (80, 40)
    assert center[:3] == (255, 0, 0)
    assert flat is None


def test_render_resizes_to_the_size_of_the_image():
    png, _ = svg_renderer.render(DIAGRAM.encode(), 60, 30, 1.0, layer=False)
    assert png_size(png)[0] == (60, 30)


@pytest.mark.skipif(shutil.which("pdftocairo") is None, reason="needs pdftocairo of Poppler")
def test_render_flattens_the_layer():
    _, flat = svg_renderer.render(DIAGRAM.encode(), 40, 20, 1.0, layer=True)
    assert flat is not None
    assert b"<svg" in flat


def test_render_leaves_the_png_alone_when_flattening_fails():
    with patch.object(convert, "to_svg", side_effect=convert.ConvertError("no pdftocairo")):
        png, flat = svg_renderer.render(DIAGRAM.encode(), 40, 20, 1.0, layer=True)
    assert png.startswith(b"\x89PNG")
    assert flat is None


@pytest.mark.asyncio
async def test_rasterizer_keeps_the_flat_svg_by_the_hash_of_the_png():
    rasterizer = svg_renderer.for_target(anyio.CapacityLimiter(1), "docx")
    with patch.object(svg_renderer, "render", return_value=(b"png", b"<svg/>")) as render:
        assert await rasterizer.convert_svg_to_png(DIAGRAM, 40, 20) == b"png"
    render.assert_called_once_with(DIAGRAM.encode(), 40, 20, 1.0, layer=True)
    assert rasterizer.svg_layers == {hashlib.sha256(b"png").digest(): b"<svg/>"}


@pytest.mark.asyncio
async def test_rasterizer_of_another_target_keeps_no_layer():
    rasterizer = svg_renderer.for_target(anyio.CapacityLimiter(1), "pptx")
    with patch.object(svg_renderer, "render", return_value=(b"png", None)) as render:
        assert await rasterizer.convert_svg_to_png(DIAGRAM, 40, 20, 3.0) == b"png"
    render.assert_called_once_with(DIAGRAM.encode(), 40, 20, 3.0, layer=False)
    assert rasterizer.svg_layers == {}


@pytest.mark.asyncio
async def test_rasterizer_counts_a_failure():
    rasterizer = svg_renderer.for_target(anyio.CapacityLimiter(1), "docx")
    with patch.object(svg_renderer, "render", side_effect=ValueError("broken")), patch.object(svg_renderer, "increment_svg_conversion_failure") as failure, pytest.raises(ValueError, match="broken"):
        await rasterizer.convert_svg_to_png(DIAGRAM, 40, 20)
    failure.assert_called_once_with()


def test_render_refuses_a_bitmap_past_the_limits_of_weasysvg():
    with pytest.raises(convert.ConvertError, match="too large"):
        svg_renderer.render(DIAGRAM.encode(), 40, 20, convert.MAX_PNG_SIDE / 20, layer=False)
    # A tall SVG drawn into a wide box: the side that needs more sets the bitmap
    tall = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="2000"><rect width="10" height="2000"/></svg>'
    with pytest.raises(convert.ConvertError, match="too large"):
        svg_renderer.render(tall.encode(), 2000, 10, 1.0, layer=False)


def test_render_holds_the_pdfium_lock():
    seen = []
    original = svg_renderer.pdfium.PdfDocument

    def document(pdf):
        seen.append(svg_renderer._PDFIUM_LOCK.locked())
        return original(pdf)

    with patch.object(svg_renderer.pdfium, "PdfDocument", side_effect=document):
        svg_renderer.render(DIAGRAM.encode(), 40, 20, 1.0, layer=False)
    assert seen == [True]
    assert not svg_renderer._PDFIUM_LOCK.locked()
