"""Unit tests for ``app.HtmlImagePreProcess``.

The preprocessor gives un-sized ``<img>`` elements an explicit px width/height
read from their inlined image bytes, so pandoc renders them at the 96 dpi CSS
reference (not its 72 dpi no-density fallback), honouring any CSS max-width.
Each test feeds an HTML snippet in and asserts on the returned bytes.
"""

from __future__ import annotations

import base64
import struct
import zlib

from app import HtmlImagePreProcess

# --- helpers --------------------------------------------------------------


def _png(width: int, height: int, *, dpi: int | None = None) -> bytes:
    """A minimal valid PNG of the given pixel size (optionally with a pHYs
    density chunk), so the header parser has something real to read."""

    def chunk(typ: bytes, data: bytes) -> bytes:
        body = typ + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    chunks = [chunk(b"IHDR", ihdr)]
    if dpi is not None:
        ppm = round(dpi / 0.0254)
        chunks.append(chunk(b"pHYs", struct.pack(">IIB", ppm, ppm, 1)))
    raw = b"".join(b"\x00" + b"\xcc\xcc\xcc" * width for _ in range(height))
    chunks.append(chunk(b"IDAT", zlib.compress(raw)))
    chunks.append(chunk(b"IEND", b""))
    return b"\x89PNG\r\n\x1a\n" + b"".join(chunks)


def _data_uri(png: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(png).decode()


def _img(png: bytes, style: str = "", extra: str = "") -> bytes:
    style_attr = f' style="{style}"' if style else ""
    return f'<p><img src="{_data_uri(png)}"{style_attr}{extra}></p>'.encode()


# --- native-size sizing ---------------------------------------------------


def test_unsized_image_gets_native_px_dimensions():
    """No style at all: width/height set to the image's pixel size (rendered at
    96 dpi by pandoc)."""
    out = HtmlImagePreProcess.preprocess(_img(_png(300, 150)))
    assert b'width="300px"' in out
    assert b'height="150px"' in out


def test_density_metadata_is_ignored():
    """A pHYs density must NOT change the emitted px size — we normalise on the
    pixel count so the physical size is pixels/96in regardless of embedded dpi."""
    out = HtmlImagePreProcess.preprocess(_img(_png(300, 150, dpi=72)))
    assert b'width="300px"' in out and b'height="150px"' in out


# --- max-width / max-height clamp -----------------------------------------


def test_max_width_wider_than_image_does_not_clamp():
    """max-width above the image's native width leaves it at native size."""
    out = HtmlImagePreProcess.preprocess(_img(_png(512, 512), "max-width:650px;"))
    assert b'width="512px"' in out and b'height="512px"' in out


def test_max_width_narrower_than_image_clamps_keeping_ratio():
    """A 400x200 image with max-width:200px scales to 200x100 (ratio kept)."""
    out = HtmlImagePreProcess.preprocess(_img(_png(400, 200), "max-width:200px;"))
    assert b'width="200px"' in out and b'height="100px"' in out


def test_max_height_clamps():
    """max-height constrains too, keeping aspect ratio."""
    out = HtmlImagePreProcess.preprocess(_img(_png(400, 200), "max-height:50px;"))
    assert b'width="100px"' in out and b'height="50px"' in out


def test_relative_max_width_is_ignored():
    """A % / em max-width has no fixed px value here, so no clamp is applied —
    the image keeps its native size."""
    out = HtmlImagePreProcess.preprocess(_img(_png(300, 150), "max-width:80%;"))
    assert b'width="300px"' in out and b'height="150px"' in out


# --- pass-through cases ---------------------------------------------------


def test_existing_width_attribute_is_untouched():
    """An explicit width attribute wins; we do not add our own."""
    out = HtmlImagePreProcess.preprocess(_img(_png(300, 150), extra=' width="99"'))
    assert b'width="99"' in out
    assert b'width="300px"' not in out


def test_explicit_css_width_left_to_lua_filter():
    """A CSS width/height is handled by filters/inline_styles.lua; this
    preprocessor must not also set an attribute for it."""
    out = HtmlImagePreProcess.preprocess(_img(_png(300, 150), "width:120px;"))
    # style is preserved, and no width/height *attribute* was injected.
    assert b"width:120px" in out
    assert b'width="' not in out.split(b"<img")[1].split(b">")[0]


def test_non_data_uri_src_is_untouched():
    """We can't read a remote/attachment reference, so leave it alone."""
    src = b'<p><img src="http://example.com/x.png"></p>'
    assert HtmlImagePreProcess.preprocess(src) == src


def test_svg_data_uri_is_left_to_svg_processor():
    """SVG data URIs are handled earlier by app.svg_processor; skip them."""
    src = b'<p><img src="data:image/svg+xml;base64,PHN2Zy8+"></p>'
    assert HtmlImagePreProcess.preprocess(src) == src


def test_corrupt_base64_is_passed_through():
    src = b'<p><img src="data:image/png;base64,@@@not-base64@@@"></p>'
    assert HtmlImagePreProcess.preprocess(src) == src


def test_unparseable_input_passed_through():
    assert HtmlImagePreProcess.preprocess(b"\xff\xfe not html") == b"\xff\xfe not html"


def test_image_free_html_unchanged():
    src = b"<p>no images here</p>"
    assert HtmlImagePreProcess.preprocess(src) == src


def test_full_document_head_is_preserved_when_sizing():
    """Regression: when the input is a full HTML document, sizing an image must
    NOT drop the <head>. The exporter sends <head><title>..</title>, which pandoc
    renders with the "Title" style; losing the head made the title fall back to
    "First Paragraph"."""
    png = _png(300, 150)
    src = ("<html><head><title>Document Title</title></head><body><p>intro</p>" + _img(png).decode() + "</body></html>").encode()
    out = HtmlImagePreProcess.preprocess(src)
    assert b"<title>Document Title</title>" in out, f"head/title dropped: {out[:120]!r}"
    assert b'width="300px"' in out and b'height="150px"' in out  # image still sized


def test_idempotent():
    """Running twice yields the same output (the second pass sees width/height
    attributes already set and skips)."""
    once = HtmlImagePreProcess.preprocess(_img(_png(300, 150)))
    twice = HtmlImagePreProcess.preprocess(once)
    assert once == twice


# --- header parsing across formats ----------------------------------------


def test_reads_gif_dimensions():
    gif = b"GIF89a" + struct.pack("<HH", 20, 10) + b"\x00" * 7
    assert HtmlImagePreProcess._read_raster_size(gif) == (20, 10)


def test_reads_bmp_dimensions():
    bmp = b"BM" + b"\x00" * 16 + struct.pack("<ii", 25, 15) + b"\x00" * 4
    assert HtmlImagePreProcess._read_raster_size(bmp) == (25, 15)


def test_reads_jpeg_dimensions():
    jpeg = b"\xff\xd8" + b"\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00" + b"\xff\xc0\x00\x11\x08" + struct.pack(">HH", 30, 40) + b"\x03\x01\x22\x00\x02\x11\x01\x03\x11\x01"
    assert HtmlImagePreProcess._read_raster_size(jpeg) == (40, 30)


def test_unknown_format_returns_none():
    assert HtmlImagePreProcess._read_raster_size(b"not an image at all!!") is None
