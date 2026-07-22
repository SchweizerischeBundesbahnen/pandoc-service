"""Extract table width and alignment from HTML source for DOCX post-processing.

Pandoc's HTML reader keeps the ``<table style="...">`` declaration in the
Table node's ``Attr`` key-value list, but the DOCX writer discards it: every
table comes out with ``<w:tblW w:type="auto"/>`` and no alignment. On top of
that, :mod:`app.DocxPostProcess` used to force every table to
``<w:tblW w:w="5000" w:type="pct"/>`` (100 %) with autofit layout, so a table
authored at ``width: 40%`` or ``margin-left: auto`` still rendered full-width
and left-aligned in Word.

This module recovers the layout intent *before* pandoc runs by parsing the
same HTML pandoc will convert. It returns one :class:`TableLayout` per
``<table>`` **in document order (depth-first, nested tables included)** — the
same order in which pandoc emits ``<w:tbl>`` elements and in which
:func:`app.DocxPostProcess._replace_table_properties` walks them, so the two
lists line up index-for-index. The DOCX post-processor consumes the list and
writes real ``<w:tblW>``/``<w:jc>``/``<w:tblInd>`` properties.

Only the properties that survive meaningfully into Word are extracted:

* **width** — ``width: N%`` becomes a percentage (OOXML ``pct`` type, where
  100 % == 5000 fiftieths of a percent); an absolute length (``width: 50px``,
  ``pt``/``cm``/…) becomes twips (OOXML ``dxa`` type). ``width: auto`` or a
  missing/unparseable width yields no width (the post-processor keeps its
  100 % default).
* **alignment** — derived from the ``margin-left``/``margin-right`` pair the
  way a browser centres/right-aligns a block: ``0``/``auto`` → left,
  ``auto``/``auto`` → center, ``auto``/``0`` → right.
* **indent** — a positive ``margin-left`` length on an otherwise left-aligned
  table becomes a left table indent (twips), mirroring the paragraph-indent
  handling in :mod:`app.HtmlParagraphPreProcess`.

Everything is best-effort: any parse failure returns an empty list so the
caller falls back to the previous behaviour rather than breaking a conversion.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass

from lxml import etree, html  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)

# OOXML: a table width of 100 % is expressed as 5000 fiftieths of a percent.
MAX_PCT = 5000

# CSS unit -> twips conversion factor. 1 twip = 1/1440 inch; CSS reference DPI
# is 96, so 1 px = 15 twips. Mirrors app/HtmlParagraphPreProcess.py so table
# and paragraph indents use one consistent conversion.
_UNIT_TO_TWIPS: dict[str, float] = {
    "px": 1440 / 96,
    "pt": 1440 / 72,
    "in": 1440.0,
    "cm": 1440 / 2.54,
    "mm": 1440 / 25.4,
    "pc": 1440 / 6,
    "em": 240.0,
    "rem": 240.0,
}

# A numeric value followed by an optional unit (letters or %). Keyword values
# such as "auto" have no leading digit and simply fail to match. The pattern is
# free of variable-width whitespace quantifiers to avoid SonarCloud's S5852
# "regex could backtrack" warning (see app/HtmlParagraphPreProcess.py).
_VALUE_RE = re.compile(r"^([+-]?\d+(?:\.\d+)?)([a-z%]*)$", re.IGNORECASE)

# Exceptions that mean "input isn't parseable HTML, give up quietly". Bound to
# a name rather than an inline tuple for the same reason as
# app/HtmlParagraphPreProcess.py (ruff/PEP 758 except-tuple rewrite).
_PARSE_FAILURES = (etree.ParseError, etree.ParserError, ValueError)


@dataclass(frozen=True)
class TableLayout:
    """Layout intent recovered from a single HTML ``<table>``.

    ``width_type`` is ``"pct"`` (``width_value`` in fiftieths of a percent),
    ``"dxa"`` (``width_value`` in twips) or ``None`` (no explicit width).
    ``jc`` is ``"left"``/``"center"``/``"right"`` or ``None``.
    ``indent_twips`` is a positive left table indent or ``None``.
    """

    width_type: str | None = None
    width_value: int | None = None
    jc: str | None = None
    indent_twips: int | None = None

    @property
    def is_empty(self) -> bool:
        """True when nothing worth applying was recovered."""
        return self.width_type is None and self.jc is None and self.indent_twips is None


def extract(source: bytes | str) -> list[TableLayout]:
    """Return one :class:`TableLayout` per ``<table>`` in document order.

    Depth-first, so a nested table appears immediately after its parent — the
    order pandoc's DOCX writer and :mod:`app.DocxPostProcess` both use. Returns
    an empty list when the input has no tables or cannot be parsed.
    """
    # Feed lxml bytes, never a decoded str: lxml rejects a Unicode string that
    # carries an XML/HTML encoding declaration ("<?xml ... encoding=...?>"),
    # which is exactly what the exporter emits. Encoding a str back to bytes
    # sidesteps that (same approach as app/HtmlParagraphPreProcess.py).
    data = source if isinstance(source, bytes) else source.encode("utf-8")
    try:
        doc = html.document_fromstring(data)
    except _PARSE_FAILURES:
        logger.warning("HtmlTableLayout: HTML parse failed; no table layouts extracted")
        return []

    # iter("table") yields elements in document order (depth-first), matching
    # both pandoc's <w:tbl> emission order and the post-processor's traversal.
    return [_parse_table_style(table.get("style") or "") for table in doc.iter("table")]


def _parse_table_style(style: str) -> TableLayout:
    """Build a :class:`TableLayout` from a table's inline ``style`` string."""
    declarations = _split_declarations(style)
    width_type, width_value = _parse_width(declarations.get("width"))
    margin_left = declarations.get("margin-left")
    margin_right = declarations.get("margin-right")
    jc = _resolve_alignment(margin_left, margin_right)
    # A positive left margin on a left-aligned table is a real indent; for
    # centered/right-aligned tables the margins are the alignment mechanism
    # (auto), not an indent, so we never treat them as one.
    indent_twips = _positive_length_twips(margin_left) if jc == "left" else None
    return TableLayout(width_type=width_type, width_value=width_value, jc=jc, indent_twips=indent_twips)


def _split_declarations(style: str) -> dict[str, str]:
    """Parse ``"k1: v1; k2: v2"`` into ``{k1: v1, k2: v2}`` (lowercased keys).

    Later declarations win, matching CSS cascade order. ``partition(":")``
    keeps ``max-width`` and ``width`` distinct — only an exact ``width`` key is
    ever read as the table width.
    """
    result: dict[str, str] = {}
    for declaration in style.split(";"):
        prop, sep, value = declaration.partition(":")
        if sep:
            result[prop.strip().lower()] = value.strip()
    return result


def _parse_width(value: str | None) -> tuple[str | None, int | None]:
    """Return ``(width_type, width_value)`` for a CSS ``width`` value.

    ``N%`` -> ``("pct", fiftieths)`` clamped to (0, 5000]; an absolute length
    -> ``("dxa", twips)``; ``auto``/missing/unparseable/zero -> ``(None, None)``.
    """
    if not value:
        return None, None
    match = _VALUE_RE.match(value)
    if not match:
        return None, None
    number = float(match.group(1))
    unit = match.group(2).lower()

    if unit == "%":
        pct = round(number * 50)
        pct = min(pct, MAX_PCT)
        return ("pct", pct) if pct > 0 else (None, None)

    # A bare number is an invalid CSS width; treat missing-unit as px to match
    # emitters that drop the unit (same convention as HtmlParagraphPreProcess).
    factor = _UNIT_TO_TWIPS.get(unit or "px")
    if factor is None:
        return None, None
    twips = round(number * factor)
    return ("dxa", twips) if twips > 0 else (None, None)


def _resolve_alignment(margin_left: str | None, margin_right: str | None) -> str | None:
    """Map the ``margin-left``/``margin-right`` pair to a table justification.

    Follows how a browser positions a fixed-width block: an ``auto`` margin
    absorbs the free space on that side. ``auto``/``auto`` centres,
    ``auto``/fixed pushes right, fixed/``auto`` keeps it left. Returns ``None``
    when neither margin is ``auto`` (no alignment intent to record).
    """
    left_auto = margin_left is not None and margin_left.strip().lower() == "auto"
    right_auto = margin_right is not None and margin_right.strip().lower() == "auto"

    if left_auto and right_auto:
        return "center"
    if left_auto:
        return "right"
    if right_auto:
        return "left"
    return None


def _positive_length_twips(value: str | None) -> int | None:
    """Return positive twips for a CSS length, or ``None`` for zero/auto/bad input."""
    if not value:
        return None
    match = _VALUE_RE.match(value)
    if not match:
        return None
    number = float(match.group(1))
    unit = match.group(2).lower()
    if unit == "%":
        return None
    factor = _UNIT_TO_TWIPS.get(unit or "px")
    if factor is None:
        return None
    twips = round(number * factor)
    return twips if twips > 0 else None
