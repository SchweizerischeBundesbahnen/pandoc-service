"""Tests for the page comparison of tests/visual.py.

The comparison is the thing every visual test rests on, so what it can and cannot see is worth
stating here rather than leaving it to be discovered by a regression it failed to catch.
"""

from __future__ import annotations

from PIL import Image

from tests.visual import MAX_DIFFERING_SHARE, PIXEL_TOLERANCE, _differing_share, render_pages

# Two colours no reader would confuse, whose greys are nearer to each other than the tolerance
GREEN = (0, 128, 0)
DULL_RED = (166, 74, 74)
PURE_RED = (255, 0, 0)


def _page(color: tuple[int, int, int]) -> Image.Image:
    return Image.new("RGB", (100, 100), color)


def _pdf_of_a_red_square() -> bytes:
    """A one page PDF holding a red square, written out here so the test needs no PDF writer."""
    content = b"1 0 0 rg 10 10 80 80 re f\n"
    objects = [
        b"<</Type /Catalog /Pages 2 0 R>>",
        b"<</Type /Pages /Kids [3 0 R] /Count 1>>",
        b"<</Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R>>",
        b"<</Length " + str(len(content)).encode() + b">>\nstream\n" + content + b"endstream",
    ]
    pdf = bytearray(b"%PDF-1.4\n")
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf += str(number).encode() + b" 0 obj\n" + body + b"\nendobj\n"
    table = len(pdf)
    pdf += b"xref\n0 " + str(len(objects) + 1).encode() + b"\n0000000000 65535 f \n"
    for offset in offsets:
        pdf += f"{offset:010d} 00000 n \n".encode()
    pdf += b"trailer\n<</Size " + str(len(objects) + 1).encode() + b" /Root 1 0 R>>\nstartxref\n" + str(table).encode() + b"\n%%EOF\n"
    return bytes(pdf)


def test_a_page_which_changed_colour_alone_is_seen() -> None:
    """A page read in grey says little about colour, and this service is full of it.

    `docx_color_pre_process`, `docx_math_color_post_process` and the colour filters all decide
    colours, so a comparison blind to them is blind to what half of that code is for.
    """
    share, _ = _differing_share(_page(GREEN), _page(DULL_RED))

    assert share == 1.0, "every pixel changed colour, and every pixel should count"
    assert share > MAX_DIFFERING_SHARE


def test_the_two_colours_are_nearer_in_grey_than_the_tolerance() -> None:
    """What the test above rests on: in grey the difference is too small to be seen."""
    green_grey = _page(GREEN).convert("L").getpixel((0, 0))
    red_grey = _page(DULL_RED).convert("L").getpixel((0, 0))

    assert (green_grey, red_grey) == (75, 102)
    # Not the same grey, but nearer than a page comparison calls a difference, so grey passes them
    assert abs(green_grey - red_grey) < PIXEL_TOLERANCE


def test_a_page_which_did_not_change_differs_nowhere() -> None:
    share, _ = _differing_share(_page(GREEN), _page(GREEN))

    assert share == 0.0


def test_a_move_smaller_than_the_tolerance_is_let_through() -> None:
    """Antialiasing moves a channel by a little, and a little is not a difference."""
    nudged = (GREEN[0], GREEN[1] + PIXEL_TOLERANCE - 1, GREEN[2])

    share, _ = _differing_share(_page(GREEN), _page(nudged))

    assert share == 0.0


def test_a_move_in_one_channel_alone_still_counts() -> None:
    """The channels are brought together after the subtraction, so one of them moving is enough."""
    moved = (GREEN[0], GREEN[1], GREEN[2] + PIXEL_TOLERANCE + 1)

    share, _ = _differing_share(_page(GREEN), _page(moved))

    assert share == 1.0


def test_a_rendered_page_carries_the_colour_of_the_pdf() -> None:
    """What the tests above rest on: the pages themselves arrive in colour.

    They are given pages made here, so on their own they would still pass where `render_pages` went
    back to reading a PDF in grey, and every visual test would go colour blind with it.
    """
    pages = render_pages(_pdf_of_a_red_square())

    assert len(pages) == 1
    assert pages[0].mode == "RGB"
    # The mode alone would hold for a grey page carried in three equal channels, so read the red
    centre = pages[0].getpixel((pages[0].width // 2, pages[0].height // 2))
    assert centre == PURE_RED
