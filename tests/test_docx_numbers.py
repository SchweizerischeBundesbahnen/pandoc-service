"""Unit tests for :mod:`app.docx_numbers`."""

import pytest

from app.docx_numbers import whole_number


@pytest.mark.parametrize(("value", "expected"), [("0", 0), ("1440", 1440), ("4294967295", 4294967295), ("007", 7)])
def test_ascii_digits_are_a_whole_number(value, expected):
    assert whole_number(value) == expected


@pytest.mark.parametrize("value", [None, "", "²", "١٢", "12.5", "1e3", " 12", "-12", "+12", "12345678901", "9" * 5000])
def test_anything_else_is_none(value):
    """A superscript two passes str.isdigit and fails int; int refuses more than 4300 digits."""
    assert whole_number(value) is None


@pytest.mark.parametrize(("value", "expected"), [("-360", -360), ("360", 360), ("-", None), ("--1", None), ("-²", None)])
def test_a_signed_number_may_start_with_a_minus(value, expected):
    assert whole_number(value, signed=True) == expected
