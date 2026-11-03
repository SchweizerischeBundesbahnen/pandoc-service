"""
Selection of the SVG renderer, and the renderer that draws SVG with weasysvg.

SVG_RENDERER picks who draws the SVG of an HTML source. ``chromium`` (default) rasterizes it to PNG
through the browser. ``weasysvg`` draws it with WeasyPrint and the weasysvg plugin, and the service
starts no browser:

- For a DOCX, the PNG goes into the document as before, and :mod:`app.docx_svg_layers` adds the SVG
  flattened by weasysvg on top of it. Word from 2016 shows the vector, an older Word the PNG.
- For a PDF or LaTeX target, the SVG stays as it is. pandoc converts it with ``rsvg-convert``, which
  is ``weasysvg-convert`` under that name: the PDF gets the SVG as vector.
- For HTML and EPUB, the SVG stays as it is: a browser draws it.
- For any other target, the PNG alone, as Chromium made it.
"""

from __future__ import annotations

import functools
import hashlib
import importlib.metadata
import io
import logging
import os
import threading
import time
from typing import TYPE_CHECKING

import pypdfium2 as pdfium
import weasysvg
from anyio import to_thread
from weasysvg import convert

from app.prometheus_metrics import increment_svg_conversion_failure, increment_svg_conversion_success

if TYPE_CHECKING:
    import anyio

CHROMIUM = "chromium"
WEASYSVG = "weasysvg"
SVG_RENDERERS = (CHROMIUM, WEASYSVG)
DEFAULT_SVG_RENDERER = CHROMIUM

# The directory that holds rsvg-convert, a link to weasysvg-convert. pandoc gets it in PATH in the weasysvg
# mode only: in the chromium mode pandoc finds no rsvg-convert, as before.
RSVG_CONVERT_DIRECTORY = "/opt/weasysvg/bin"

# The targets pandoc converts an SVG for itself, through rsvg-convert, or a browser draws
_SVG_KEPT_TARGETS = frozenset(("pdf", "latex", "html", "epub"))
# The targets that get the flat SVG on top of the PNG
_SVG_LAYER_TARGETS = frozenset(("docx",))

# PDFium is not thread-safe, not even across documents: one render at a time in the process
_PDFIUM_LOCK = threading.Lock()

logger = logging.getLogger(__name__)


def get_svg_renderer() -> str:
    """
    Read the SVG renderer from the SVG_RENDERER variable.

    Returns:
        ``chromium`` or ``weasysvg``. The value is case-insensitive and may carry whitespace. An unset or
        empty value is ``chromium``; an unknown one too, with a warning.
    """
    value = os.environ.get("SVG_RENDERER", "").strip().lower()
    if not value:
        return DEFAULT_SVG_RENDERER
    if value not in SVG_RENDERERS:
        logger.warning("Invalid SVG_RENDERER value '%s', using default: %s", value, DEFAULT_SVG_RENDERER)
        return DEFAULT_SVG_RENDERER
    return value


def configure_svg_renderer() -> str:
    """
    Select the SVG renderer and prepare it.

    weasysvg patches WeasyPrint for the whole process, so it is installed only when it is selected.

    Returns:
        The selected renderer.
    """
    renderer = get_svg_renderer()
    if renderer == WEASYSVG:
        weasysvg.install()
    return renderer


def get_weasysvg_version() -> str | None:
    """The installed weasysvg version, or None when it is missing."""
    try:
        return importlib.metadata.version("weasysvg")
    except importlib.metadata.PackageNotFoundError:
        return None


def keeps_svg(target_format: str) -> bool:
    """Whether the weasysvg mode leaves the SVG of a document as it is for this target."""
    return target_format in _SVG_KEPT_TARGETS


def pandoc_environment(environment: dict[str, str] | None) -> dict[str, str]:
    """The environment of pandoc in the weasysvg mode: rsvg-convert first in PATH."""
    base = dict(os.environ) if environment is None else dict(environment)
    base["PATH"] = os.pathsep.join(filter(None, (RSVG_CONVERT_DIRECTORY, base.get("PATH"))))
    return base


def render(svg: bytes, width: int, height: int, scale: float, *, layer: bool) -> tuple[bytes, bytes | None]:
    """
    Draw an SVG as a PNG of width x height px times scale, on a transparent background, as Chromium does.

    With layer, also flatten it into the SVG Word draws: see weasysvg-convert -f svg. A failing flattening
    leaves the PNG alone.

    Returns:
        The PNG, and the flat SVG or None.
    """
    target = (max(1, round(width * scale)), max(1, round(height * scale)))
    pdf = convert.to_pdf(svg)
    with _PDFIUM_LOCK:
        document = pdfium.PdfDocument(pdf)
        page = document[0]
        try:
            # Drawn at the target size along the side that needs more, then brought to the target
            zoom = max(target[0] / page.get_width(), target[1] / page.get_height())
            drawn = (round(page.get_width() * zoom), round(page.get_height() * zoom))
            # The limits of weasysvg-convert: a bitmap past them would take the memory of the service
            if max(drawn) > convert.MAX_PNG_SIDE or drawn[0] * drawn[1] > convert.MAX_PNG_PIXELS:
                raise convert.ConvertError(f"a PNG of {drawn[0]} x {drawn[1]} px is too large")
            bitmap = page.render(scale=zoom, fill_color=(255, 255, 255, 0))
            try:
                image = bitmap.to_pil()
            finally:
                bitmap.close()
        finally:
            page.close()
            document.close()
    if image.size != target:
        image = image.resize(target)
    out = io.BytesIO()
    image.save(out, format="PNG")

    flat = None
    if layer:
        try:
            flat = convert.to_svg(svg)
        except convert.ConvertError as e:
            logger.warning("Could not flatten an SVG, the PNG goes alone: %s", e)
    return out.getvalue(), flat


class WeasysvgRasterizer:
    """
    Draws SVG with weasysvg for SvgProcessor, in place of ChromiumManager.

    SvgProcessor calls convert_svg_to_png as it calls it on ChromiumManager. One instance serves one
    request: it keeps the flat SVG of each PNG it made, by the SHA-256 of the PNG, for
    :mod:`app.docx_svg_layers`.
    """

    def __init__(self, limiter: anyio.CapacityLimiter, *, layers: bool) -> None:
        self._limiter = limiter
        self._layers = layers
        self.svg_layers: dict[bytes, bytes] = {}

    async def convert_svg_to_png(self, svg_content: str, width: int, height: int, device_scale_factor: float | None = None) -> bytes:
        """The PNG of an SVG, drawn in a worker thread under the conversion limiter."""
        scale = device_scale_factor or 1.0
        start = time.time()
        try:
            png, flat = await to_thread.run_sync(
                functools.partial(render, svg_content.encode("utf-8"), width, height, scale, layer=self._layers),
                abandon_on_cancel=True,
                limiter=self._limiter,
            )
        except Exception:
            increment_svg_conversion_failure()
            raise
        increment_svg_conversion_success(time.time() - start)
        if flat is not None:
            self.svg_layers[hashlib.sha256(png).digest()] = flat
        return png


def for_target(limiter: anyio.CapacityLimiter, target_format: str) -> WeasysvgRasterizer:
    """The rasterizer of one request to the given target."""
    return WeasysvgRasterizer(limiter, layers=target_format in _SVG_LAYER_TARGETS)
