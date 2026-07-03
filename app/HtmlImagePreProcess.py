"""Give un-sized ``<img>`` elements an explicit pixel width/height at 96 dpi.

Pandoc's DOCX writer sizes an image with no explicit ``width``/``height`` from
the image's pixel dimensions divided by its embedded density (``pHYs`` for PNG,
JFIF for JPEG). Screenshots and Polarion attachments usually carry *no* density,
so pandoc falls back to **72 dpi** — while a browser lays the same image out at
the CSS reference of **96 dpi**. The result is a DOCX image ~96/72 = 1.33x larger
than it appears in the source (less when pandoc additionally clamps an
over-wide image to the page text width, which is why the discrepancy varies).

To make images match their browser size we set an explicit pixel ``width`` (and
``height``) on each ``<img>`` that has none: pandoc converts a ``px`` dimension
at a fixed 96 dpi, so the physical size becomes ``pixels / 96`` inches
regardless of the (missing/odd) embedded density. We also honour a CSS
``max-width`` / ``max-height`` by clamping — an image wider than its max-width is
scaled down just as the browser would, preserving aspect ratio.

Only images whose bytes we can read are touched: the exporter inlines images as
``data:`` URIs, so we decode the base64 payload and read the pixel size straight
from the PNG/JPEG/GIF/BMP header (no image library needed). Anything we can't
size — a non-``data:`` ``src`` (http/attachment reference), an SVG data URI
(handled earlier by :mod:`app.svg_processor`, which already sets a width), an
unknown format, or an image that already carries a ``width``/``height`` attribute
or CSS ``width``/``height`` (handled by ``filters/inline_styles.lua``) — is left
untouched.

The transformation is only applied for ``html -> docx`` conversions.
"""

from __future__ import annotations

import base64
import binascii
import logging
import struct

from lxml import etree, html  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)

# CSS px per inch (the browser/CSS reference resolution). Setting an explicit
# px dimension makes pandoc render the image at this dpi.
_CSS_DPI = 96

# See HtmlListsPreProcess for why these are named tuples rather than inline
# except-literals (ruff-format / PEP 758 interaction on Python 3.14).
_PARSE_FAILURES = (etree.ParseError, etree.ParserError, ValueError)
_DECODE_FAILURES = (binascii.Error, ValueError)

# CSS length units we can resolve to px for the max-width/max-height clamp.
# Absolute units use the same 96 px/in reference; a bare number is px. Relative
# units (%, em, vw, ...) have no fixed px value here, so those clamps are ignored.
_PX_PER_UNIT = {
    "": 1.0,
    "px": 1.0,
    "in": float(_CSS_DPI),
    "cm": _CSS_DPI / 2.54,
    "mm": _CSS_DPI / 25.4,
    "pt": _CSS_DPI / 72.0,
    "pc": _CSS_DPI / 6.0,
}


def preprocess(source: bytes) -> bytes:
    """Set explicit px ``width``/``height`` on un-sized, decodable ``<img>``.

    Returns the input unchanged when parsing fails or no image is rewritten.
    """
    try:
        # document_fromstring (not fragments_fromstring) so a full HTML document
        # keeps its <head> — the exporter sends <head><title>...</title>, which
        # pandoc renders with the "Title" style; fragments_fromstring drops the
        # head and the title would fall back to "First Paragraph". A bare
        # fragment is harmlessly wrapped in <html><body> (pandoc reads it the
        # same), and we only re-serialize at all when an image was actually sized.
        doc = html.document_fromstring(source)
    except _PARSE_FAILURES:
        logger.warning("HtmlImagePreProcess: HTML parse failed; passing input through")
        return source

    if not _size_images(doc):
        return source

    return html.tostring(doc, encoding="utf-8")


def _size_images(root: html.HtmlElement) -> bool:
    """Set width/height on every eligible ``<img>`` under ``root``."""
    rewrote = False
    for img in root.iter("img"):
        if _size_one_image(img):
            rewrote = True
    return rewrote


def _size_one_image(img: html.HtmlElement) -> bool:
    """Size a single ``<img>``; return True if it was modified."""
    # Already sized via an HTML attribute — leave it (an explicit <img width=..>
    # or the value app/svg_processor.py sets for rasterised SVGs).
    if img.get("width") or img.get("height"):
        return False

    style = _parse_style(img.get("style"))
    # An explicit CSS width/height is handled by filters/inline_styles.lua; only
    # the "no intrinsic size given" case is ours.
    if "width" in style or "height" in style:
        return False

    size = _decode_image_size(img.get("src"))
    if size is None:
        return False
    width_px, height_px = size

    scale = _clamp_scale(width_px, height_px, style)
    out_w = max(1, round(width_px * scale))
    out_h = max(1, round(height_px * scale))
    img.set("width", f"{out_w}px")
    img.set("height", f"{out_h}px")
    return True


def _clamp_scale(width_px: int, height_px: int, style: dict[str, str]) -> float:
    """Largest scale <= 1 that fits both max-width and max-height (px). 1.0 when
    neither applies or the image already fits."""
    scale = 1.0
    max_w = _length_to_px(style.get("max-width"))
    if max_w is not None and width_px > max_w:
        scale = min(scale, max_w / width_px)
    max_h = _length_to_px(style.get("max-height"))
    if max_h is not None and height_px > max_h:
        scale = min(scale, max_h / height_px)
    return scale


def _parse_style(style: str | None) -> dict[str, str]:
    """Split a CSS declaration list into a lowercased ``{prop: value}`` dict."""
    props: dict[str, str] = {}
    if not style:
        return props
    for decl in style.split(";"):
        key, sep, value = decl.partition(":")
        if sep:
            props[key.strip().lower()] = value.strip().lower()
    return props


def _length_to_px(value: str | None) -> float | None:
    """Convert a CSS length to px, or None when it isn't an absolute length
    (missing, ``%``, ``em``, ``vw``, unparseable)."""
    if not value:
        return None
    num = value.rstrip("abcdefghijklmnopqrstuvwxyz%").strip()
    unit = value[len(num) :].strip()
    try:
        magnitude = float(num)
    except ValueError:
        return None
    factor = _PX_PER_UNIT.get(unit)
    if factor is None or magnitude <= 0:
        return None
    return magnitude * factor


def _decode_image_size(src: str | None) -> tuple[int, int] | None:
    """Return (width_px, height_px) for a ``data:`` image URI we can read, else
    None (non-data URI, SVG, unknown/corrupt format)."""
    if not src or not src.startswith("data:"):
        return None
    header, _, payload = src.partition(",")
    if "svg" in header:  # SVGs are handled by app.svg_processor
        return None
    if "base64" not in header:
        return None
    try:
        raw = base64.b64decode(payload, validate=False)
    except _DECODE_FAILURES:
        return None
    return _read_raster_size(raw)


# Smallest header we bother inspecting (a GIF logical-screen descriptor); PNG
# and BMP need more and are length-checked in their own branch.
_MIN_HEADER_BYTES = 10
_PNG_HEADER_BYTES = 24  # 8-byte signature + IHDR through the height field
_BMP_HEADER_BYTES = 26  # 'BM' + file header + DIB header through height

# JPEG marker bytes. Every marker is prefixed with 0xFF. SOI/EOI and the restart
# markers (RSTn) are "standalone" — they carry no length field, so we step over
# them 2 bytes at a time; every other marker has a 2-byte segment length.
_JPEG_MARKER_PREFIX = 0xFF
_JPEG_STANDALONE = frozenset({0xD8, 0xD9})  # SOI, EOI
_JPEG_RST_FIRST, _JPEG_RST_LAST = 0xD0, 0xD7  # RST0..RST7
_JPEG_MIN_SEGMENT_LEN = 2  # a segment length includes its own 2 bytes
# SOF markers that carry frame dimensions (baseline/progressive/etc.), excluding
# the non-SOF markers in the 0xC0-0xCF range (DHT 0xC4, DAC 0xCC, RSTn).
_JPEG_SOF_MARKERS = frozenset({0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF})


def _read_raster_size(data: bytes) -> tuple[int, int] | None:
    """Read pixel dimensions from a PNG/GIF/BMP/JPEG byte header (each branch
    checks it has enough bytes before unpacking; returns None otherwise)."""
    if len(data) < _MIN_HEADER_BYTES:
        return None
    # PNG: 8-byte signature, then IHDR (width, height as big-endian uint32 at 16).
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= _PNG_HEADER_BYTES:
        width, height = struct.unpack(">II", data[16:24])
        return width, height
    # GIF: 'GIF87a'/'GIF89a', logical screen width/height little-endian uint16.
    if data[:6] in (b"GIF87a", b"GIF89a"):
        width, height = struct.unpack("<HH", data[6:10])
        return width, height
    # BMP: 'BM', DIB header width/height little-endian int32 at 18/22.
    if data[:2] == b"BM" and len(data) >= _BMP_HEADER_BYTES:
        width, height = struct.unpack("<ii", data[18:26])
        return abs(width), abs(height)
    # JPEG: scan for a start-of-frame marker carrying the dimensions.
    if data[:2] == b"\xff\xd8":
        return _read_jpeg_size(data)
    return None


def _read_jpeg_size(data: bytes) -> tuple[int, int] | None:
    """Walk JPEG segment markers to the start-of-frame and read its size."""
    i = 2  # skip the SOI (\xff\xd8)
    n = len(data)
    while i + 9 < n:
        if data[i] != _JPEG_MARKER_PREFIX:
            i += 1
            continue
        marker = data[i + 1]
        if marker in _JPEG_SOF_MARKERS:
            # SOF: FF, marker, len(2), precision(1), height(2), width(2)
            height, width = struct.unpack(">HH", data[i + 5 : i + 9])
            return width, height
        if marker in _JPEG_STANDALONE or _JPEG_RST_FIRST <= marker <= _JPEG_RST_LAST:
            i += 2  # standalone markers (SOI/EOI/RSTn) have no length field
            continue
        seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
        if seg_len < _JPEG_MIN_SEGMENT_LEN:
            return None
        i += 2 + seg_len
    return None
