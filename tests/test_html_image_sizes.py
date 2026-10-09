"""Unit tests for :mod:`app.html_image_sizes`."""

import base64
import struct
import zlib

from app.html_image_sizes import RequestedSize, digest, extract


def _png(width: int, height: int, shade: int = 0x80) -> bytes:
    def chunk(tag: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + tag + payload + struct.pack(">I", zlib.crc32(tag + payload))

    raw = b"".join(b"\x00" + bytes([shade]) * (width * 3) for _ in range(height))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def _src(image: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(image).decode()


def _sizes(img_attributes: str, image: bytes) -> list[RequestedSize]:
    return extract(f'<html><body><p><img src="{_src(image)}" {img_attributes}/></p></body></html>')[digest(image)]


def test_the_css_size_of_an_image_is_recorded():
    assert _sizes('style="width: 1300px;height: 732px;"', _png(650, 366)) == [RequestedSize("1300px", "732px")]


def test_an_attribute_wins_over_the_css_on_its_own_side():
    assert _sizes('width="200" style="width: 1300px; height: 50%"', _png(650, 366)) == [RequestedSize("200", "50%")]


def test_an_image_stating_no_size_records_its_own_less_its_max_width():
    """As app/html_image_pre_process.py gives it to pandoc: 650 px wide, brought to a max-width of 325."""
    assert _sizes('style="max-width: 325px;"', _png(650, 366)) == [RequestedSize("325px", "183px")]


def test_a_size_the_css_leaves_out_is_left_empty():
    assert _sizes('style="width: 50%"', _png(650, 366)) == [RequestedSize("50%", "")]


def test_images_with_the_same_bytes_keep_their_sizes_in_document_order():
    image = _png(20, 10)
    html = f'<html><body><img src="{_src(image)}" style="width: 40px"/><p><img src="{_src(image)}" style="width: 80px"/></p></body></html>'

    assert extract(html)[digest(image)] == [RequestedSize("40px", ""), RequestedSize("80px", "")]


def test_images_with_other_bytes_are_kept_apart():
    first, second = _png(20, 10, 0x10), _png(20, 10, 0x20)
    html = f'<html><body><img src="{_src(first)}" width="1"/><img src="{_src(second)}" width="2"/></body></html>'

    sizes = extract(html)

    assert sizes[digest(first)] == [RequestedSize("1", "")]
    assert sizes[digest(second)] == [RequestedSize("2", "")]


def test_an_image_which_is_not_a_base64_data_uri_is_not_recorded():
    html = '<html><body><img src="https://example.invalid/a.png" width="10"/><img src="data:image/svg+xml;utf8,<svg/>" width="10"/></body></html>'

    assert extract(html) == {}


def test_a_str_source_is_read_as_well_as_bytes():
    image = _png(4, 2)

    assert extract(f'<img src="{_src(image)}" width="8">'.encode()) == extract(f'<img src="{_src(image)}" width="8">')


def test_an_unreadable_source_records_nothing():
    assert extract(b"") == {}
