"""Integration tests for ``filters/html_tables_to_latex.lua``.

Runs the real ``pandoc`` binary (html -> latex) with the filter and inspects
the generated LaTeX. The filter recovers table width and horizontal alignment
from the ``<table style>`` that pandoc's LaTeX writer would otherwise ignore
(every table renders content-width and centered).

Contract checked here:
* ``margin`` pair -> longtable ``\\LTleft``/``\\LTright`` glue (left/center/right).
* ``width: N%`` -> column widths summing to ~N% of the line (pandoc emits
  ``\\real{fraction}`` column widths).
* tables with no style are left untouched.
"""

from __future__ import annotations

import re
import shutil
import subprocess

import pytest

_PANDOC = shutil.which("pandoc")
pytestmark = pytest.mark.skipif(_PANDOC is None, reason="pandoc binary not available")

_FILTER = "filters/html_tables_to_latex.lua"


def _html_to_latex(body: str) -> str:
    html = f"<html><head><title>t</title></head><body>{body}</body></html>"
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "html", "-t", "latex", "--lua-filter", _FILTER],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    return completed.stdout.decode()


def _table(style: str, cols: str = "<td>a</td><td>b</td><td>c</td>") -> str:
    return f'<table style="{style}"><tr>{cols}</tr></table>'


def test_left_aligned_table_pins_ltleft_to_zero():
    latex = _html_to_latex(_table("width: 100%; margin-left: 0px; margin-right: auto;"))
    assert "\\setlength{\\LTleft}{0pt}" in latex
    assert "\\setlength{\\LTright}{\\fill}" in latex


def test_right_aligned_table_pins_ltright_to_zero():
    latex = _html_to_latex(_table("width: 25%; margin-left: auto; margin-right: 0px;"))
    assert "\\setlength{\\LTleft}{\\fill}" in latex
    assert "\\setlength{\\LTright}{0pt}" in latex


def test_centered_table_uses_fill_on_both_sides():
    latex = _html_to_latex(_table("width: 25%; margin-left: auto; margin-right: auto;"))
    # The alignment block sets both to \fill (also the reset value), so just
    # assert the table was wrapped (glue present) and not pinned to an edge.
    assert "\\setlength{\\LTleft}{\\fill}" in latex
    assert "\\setlength{\\LTright}{0pt}" not in latex
    assert "\\setlength{\\LTleft}{0pt}" not in latex


def test_percentage_width_sets_fractional_columns():
    latex = _html_to_latex(_table("width: 40%; margin-left: 0px; margin-right: auto;"))
    fractions = [float(x) for x in re.findall(r"\\real\{([0-9.]+)\}", latex)]
    assert fractions, "expected \\real{} column widths in the longtable preamble"
    # Three columns summing to ~0.40 of the line width.
    assert abs(sum(fractions) - 0.40) < 0.01


def test_absolute_width_is_a_small_fraction():
    latex = _html_to_latex(_table("width: 100px; margin-left: 0px; margin-right: auto;"))
    fractions = [float(x) for x in re.findall(r"\\real\{([0-9.]+)\}", latex)]
    assert fractions
    # 100px = 75pt against a 468pt reference text width ~= 0.16 of the line.
    assert abs(sum(fractions) - (75.0 / 468.0)) < 0.02


def test_table_without_style_is_untouched():
    latex = _html_to_latex("<table><tr><td>a</td><td>b</td></tr></table>")
    assert "\\LTleft" not in latex
    assert "\\real{" not in latex


def test_non_latex_target_is_untouched():
    """The filter is a no-op for non-LaTeX writers (defensive FORMAT gate)."""
    html = f"<html><body>{_table('width: 40%; margin-left: auto; margin-right: auto;')}</body></html>"
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "html", "-t", "html", "--lua-filter", _FILTER],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    assert "LTleft" not in completed.stdout.decode()
