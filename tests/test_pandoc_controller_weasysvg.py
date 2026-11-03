"""The controller draws SVG with weasysvg: the SVG kept as vector, or drawn as a PNG with its flat SVG."""

import base64
import hashlib
import os
from unittest.mock import MagicMock, patch

import pytest
from starlette.testclient import TestClient

from app import pandoc_controller, svg_renderer

SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="red"/></svg>'
SOURCE = f'<html><body><img src="data:image/svg+xml;base64,{base64.b64encode(SVG.encode()).decode()}"></body></html>'


@pytest.fixture
def weasysvg_mode():
    with patch.dict(os.environ, {"ENABLE_SVG_CONVERSION": "true"}):
        yield


@pytest.mark.asyncio
@pytest.mark.parametrize("target", ["pdf", "latex", "html", "epub"])
async def test_preprocess_keeps_the_svg_for_a_target_that_draws_it(weasysvg_mode, target):
    assert await pandoc_controller.preprocess_html_svgs(SOURCE, None, target) == SOURCE


@pytest.mark.asyncio
async def test_preprocess_draws_the_png_and_keeps_its_flat_svg_for_docx(weasysvg_mode):
    layers: dict[bytes, bytes] = {}
    png = b"\x89PNG drawn by weasysvg"
    with patch.object(svg_renderer, "render", return_value=(png, b"<svg flat/>")):
        result = await pandoc_controller.preprocess_html_svgs(SOURCE.encode(), None, "docx", layers)
    assert isinstance(result, bytes)
    assert base64.b64encode(png) in result
    assert layers == {hashlib.sha256(png).digest(): b"<svg flat/>"}


@pytest.mark.asyncio
async def test_preprocess_draws_the_png_alone_for_pptx(weasysvg_mode):
    with patch.object(svg_renderer, "render", return_value=(b"\x89PNG", None)) as render:
        result = await pandoc_controller.preprocess_html_svgs(SOURCE, 2.0, "pptx")
    assert "data:image/png;base64," in result
    assert render.call_args.kwargs == {"layer": False}


@pytest.mark.asyncio
async def test_preprocess_passes_the_html_through_when_drawing_fails(weasysvg_mode):
    layers: dict[bytes, bytes] = {}
    with patch.object(pandoc_controller, "SvgProcessor", side_effect=RuntimeError("broken")):
        assert await pandoc_controller.preprocess_html_svgs(SOURCE, None, "docx", layers) == SOURCE
    assert layers == {}


def test_version_and_health_name_the_renderer(weasysvg_mode):
    with patch("app.pandoc_controller.get_pandoc_version", return_value="3.1.9"):
        result = pandoc_controller.version()
        response = TestClient(pandoc_controller.app).get("/health").json()
    assert result.svgRenderer == "weasysvg"
    assert result.weasysvg == svg_renderer.get_weasysvg_version()
    assert response["svgRenderer"] == "weasysvg"
    assert "chromium" not in response


def test_pandoc_finds_rsvg_convert(weasysvg_mode):
    with patch.object(pandoc_controller, "_run_pandoc") as run, patch("pathlib.Path.open", MagicMock()), patch("pathlib.Path.exists", return_value=True), patch("pathlib.Path.unlink"):
        pandoc_controller.run_pandoc_conversion("# Title", "markdown", "pdf")
    environment = run.call_args.args[1]
    assert environment["PATH"].startswith(svg_renderer.RSVG_CONVERT_DIRECTORY)


def test_postprocess_attaches_the_svg_layers_to_a_docx():
    layers = {b"key": b"<svg/>"}
    with patch.object(pandoc_controller.docx_post_process, "process", return_value=b"processed"), patch.object(pandoc_controller.docx_svg_layers, "attach", return_value=b"layered") as attach:
        response = pandoc_controller.postprocess_and_build_response(b"docx", "docx", "a.docx", svg_layers=layers)
    attach.assert_called_once_with(b"processed", layers)
    assert response.body == b"layered"


INLINE = '<html><body><svg width="40" height="20" viewBox="0 0 40 20"><switch><foreignObject requiredFeatures="x"><div>Label</div></foreignObject></switch></svg></body></html>'


@pytest.mark.asyncio
async def test_preprocess_gives_pandoc_an_inline_svg_as_a_data_url(weasysvg_mode):
    with patch.object(svg_renderer, "render") as render:
        result = await pandoc_controller.preprocess_html_svgs(INLINE.encode(), None, "pdf")
    render.assert_not_called()
    src = result.decode().split('src="data:image/svg+xml;base64,', 1)[1].split('"', 1)[0]
    svg = base64.b64decode(src).decode()
    assert 'viewBox="0 0 40 20"' in svg
    assert "<foreignObject" in svg


@pytest.mark.asyncio
async def test_preprocess_passes_an_inline_svg_through_when_it_cannot_parse(weasysvg_mode):
    with patch.object(pandoc_controller.SvgProcessor, "replace_inline_svgs_with_img", side_effect=RuntimeError("broken")):
        assert await pandoc_controller.preprocess_html_svgs(INLINE, None, "pdf") == INLINE
