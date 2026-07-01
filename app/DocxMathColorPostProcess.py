r"""Decode math color markers into real OMML color (companion to HtmlMathColorPreProcess).

``HtmlMathColorPreProcess`` rewrites ``\color``/``\textcolor`` inside math scripts
into ``\text{@@PMC:RRGGBB@@}...\text{@@PMCEND@@}`` markers before pandoc runs,
because ``texmath`` cannot carry color through to OMML. Pandoc emits each
``\text{}`` marker as its own ``<m:r>`` run, with the formerly-colored content as
separate runs between the start and end markers in document order.

:func:`apply_math_colors` walks the OMML runs of the document **in place** (it operates
on the lxml tree ``DocxPostProcess`` has already parsed, so there is no extra
serialize/parse round-trip), maintaining a stack of active colors:

* a run whose text is a start marker pushes its color and is deleted;
* a run whose text is the end marker pops and is deleted;
* every other run, while the stack is non-empty, gets ``<w:color w:val="RRGGBB"/>``
  (the innermost/top-of-stack color) added to its ``<w:rPr>``.

Word renders per-run color inside equations, so the formula stays a live, editable
equation *and* regains the color that ``texmath`` dropped.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from lxml import etree  # type: ignore[import-untyped]

from app.HtmlMathColorPreProcess import MARKER_END, MARKER_PREFIX, MARKER_SUFFIX

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR OOXML namespace identifier (ECMA-376), never dereferenced
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"  # NOSONAR OOXML math namespace identifier (ECMA-376), never dereferenced

_M_R = f"{{{M_NS}}}r"
_M_T = f"{{{M_NS}}}t"
_M_RPR = f"{{{M_NS}}}rPr"
_W_RPR = f"{{{W_NS}}}rPr"
_W_COLOR = f"{{{W_NS}}}color"
_W_VAL = f"{{{W_NS}}}val"

# A start-marker run's text is exactly "@@PMC:RRGGBB@@"; capture the 6-hex color.
_START_MARKER_RE = re.compile(rf"^{re.escape(MARKER_PREFIX)}([0-9A-Fa-f]{{6}}){re.escape(MARKER_SUFFIX)}$")


def apply_math_colors(doc: DocumentObject) -> None:
    """Colorize the math runs between marker pairs and remove the markers, mutating the
    document in place. A document with no markers is left untouched.

    Kept parallel to ``DocxReferencesPostProcess.add_table_of_contents_entries``: a public
    function that takes the ``Document`` and is called directly from ``DocxPostProcess.process``.
    """
    color_stack: list[str] = []
    marker_runs: list[etree._Element] = []

    # Walking the live tree is safe here: marker runs are removed in a second pass
    # (below), never during this walk, and _apply_color inserts only <w:rPr> (never an
    # <m:r>), so the set of runs this loop visits does not change under it.
    for run in doc.element.iter(_M_R):
        text = _run_text(run)
        if text == MARKER_END:
            if color_stack:
                color_stack.pop()
            marker_runs.append(run)
            continue
        start = _START_MARKER_RE.match(text) if text is not None else None
        if start is not None:
            color_stack.append(start.group(1).upper())
            marker_runs.append(run)
            continue
        if color_stack:
            _apply_color(run, color_stack[-1])

    for run in marker_runs:
        parent = run.getparent()
        if parent is not None:  # pragma: no branch - a marker run always has a parent
            parent.remove(run)


def _run_text(run: etree._Element) -> str | None:
    """The text of a math run's ``<m:t>`` child, or ``None`` if it has none."""
    text_element = run.find(_M_T)
    if text_element is None:
        return None
    return text_element.text


def _apply_color(run: etree._Element, hex_color: str) -> None:
    """Set ``<w:color w:val=hex_color>`` on a math run, creating its ``<w:rPr>`` if absent.
    ``<w:rPr>`` sits after the math run properties ``<m:rPr>`` (if present) and before
    ``<m:t>``, matching how Word writes colored equation runs.
    """
    w_rpr = run.find(_W_RPR)
    if w_rpr is None:
        w_rpr = run.makeelement(_W_RPR, {})
        m_rpr = run.find(_M_RPR)
        insert_at = list(run).index(m_rpr) + 1 if m_rpr is not None else 0
        run.insert(insert_at, w_rpr)
    color = w_rpr.find(_W_COLOR)
    if color is None:
        color = etree.SubElement(w_rpr, _W_COLOR)
    color.set(_W_VAL, hex_color)
