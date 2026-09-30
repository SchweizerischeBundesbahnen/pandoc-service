"""Tests for the page comparison of tests/visual.py.

The comparison is the thing every visual test rests on, so what it can and cannot see is worth
stating here rather than leaving it to be discovered by a regression it failed to catch.
"""

from __future__ import annotations

from PIL import Image

from tests.visual import MAX_DIFFERING_SHARE, PIXEL_TOLERANCE, _differing_share

# Two colours no reader would confuse, and the same brightness under the luma of `convert("L")`
GREEN = (0, 128, 0)
DULL_RED = (166, 74, 74)


def _page(color: tuple[int, int, int]) -> Image.Image:
    return Image.new("RGB", (100, 100), color)


def test_a_page_which_changed_colour_alone_is_seen() -> None:
    """A page read in grey says nothing about colour, and this service is full of it.

    `docx_color_pre_process`, `docx_math_color_post_process` and the colour filters all decide
    colours, so a comparison blind to them is blind to what half of that code is for.
    """
    share, _ = _differing_share(_page(GREEN), _page(DULL_RED))

    assert share == 1.0, "every pixel changed colour, and every pixel should count"
    assert share > MAX_DIFFERING_SHARE


def test_the_two_colours_of_that_page_are_the_same_grey() -> None:
    """What the test above rests on: in grey the difference is not there to be seen."""
    assert _page(GREEN).convert("L").getpixel((0, 0)) == 75
    assert _page(DULL_RED).convert("L").getpixel((0, 0)) == 102
    # Within the tolerance a page comparison allows, so a grey comparison passes them as equal
    assert abs(75 - 102) < PIXEL_TOLERANCE


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
