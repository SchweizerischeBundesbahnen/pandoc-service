"""Rewrite table properties in a DOCX so tables survive into LaTeX/PDF.

Pandoc's DOCX reader has three table-related problems this preprocessor fixes:

1. **Missing grid-column widths**: ``<w:gridCol/>`` elements without a
   ``w:w`` attribute cause pandoc to produce an empty ``ColSpec`` list,
   which in turn drops every cell in the table.  This preprocessor adds
   equal-width ``w:w`` attributes when they are missing so pandoc can
   compute column proportions and preserve cell content.

2. **Dropped cell formatting**: pandoc strips cell-level properties
   (background shading, borders, vertical alignment) before building the
   AST.  This preprocessor prepends a PUA-delimited sentinel to the first
   paragraph of each coloured cell so the companion Lua filter
   (``filters/docx/docx_tables_to_latex.lua``) can inject ``\\cellcolor`` into
   the LaTeX output.

3. **Dropped table width & alignment**: pandoc normalises the column widths
   to sum to 1.0 (losing the table's ``<w:tblW>`` share of the page) and
   discards the table's ``<w:jc>`` alignment entirely, so every table renders
   full-width and centered in the PDF.  This preprocessor records the table's
   width fraction and alignment on the first cell's sentinel; the Lua filter
   scales the ``ColSpec`` widths back down and emits longtable ``\\LTleft``/
   ``\\LTright`` glue.  Every table with an explicit ``<w:tblW>`` is tagged,
   including 100% ones (pandoc may otherwise render them content-width and
   centered when the DOCX carries no usable column widths).

Sentinel format (PUA delimiters U+E010 / U+E011), a ``;``-separated key map::

    <U+E010>bg=RRGGBB;ha=justify;tw=0.4000;ta=left<U+E011>

``bg``, ``ha``, ``li`` and ``va`` are per cell; ``tw`` (0..1 line fraction) and ``ta``
(left/center/right) are table-level and live on the first cell (merged into
its sentinel if it also carries a per-cell key).

``ha`` carries justification only. Pandoc's AST has no justified alignment —
its ``Alignment`` is left/right/center/default — so a cell's ``w:jc="both"``
reaches the reader as ``AlignLeft`` and ``distribute`` as ``AlignDefault``,
and either way the text comes out flush left. The other three alignments
survive on the ``Cell`` and need no sentinel.

``li`` is the left indent of the cell's first paragraph, in points. pandoc reads
an indented paragraph as a BlockQuote and the filter unwraps it, so the indent
rides on the sentinel.

``va`` is the vertical alignment of a cell of the header row: ``t``, ``c`` or ``b``. pandoc sets
every cell of the header at the bottom of its row; Word takes the alignment of the cell, else that
of the first row or of the whole table in the table style, else the top.

This is the table companion to :mod:`app.docx_color_pre_process` /
:mod:`app.docx_paragraph_pre_process`; it runs on the same docx->latex path
and is independent of them (table cells carry ``<w:tcPr>``, not
run/paragraph properties).
"""

from __future__ import annotations

import logging
from types import MappingProxyType
from typing import TYPE_CHECKING, NamedTuple
from xml.etree import ElementTree as ET

from . import docx_page_geometry
from .docx_numbers import whole_number
from .docx_ooxml import W_NS, enumerate_body_parts, parse_xml, read_entries, repack, serialize_tree

if TYPE_CHECKING:
    from collections.abc import Mapping

logger = logging.getLogger(__name__)

# PUA delimiters — different from list-level (\uE000/\uE001) to avoid
# ambiguity when a list paragraph happens to sit inside a table cell.
# The style a Polarion caption keeps instead of Caption, so pandoc does not number it again.
PLAIN_CAPTION_STYLE = "PandocCaption"
SENTINEL_OPEN = "\ue010"
SENTINEL_CLOSE = "\ue011"

_TBL_TAG = f"{{{W_NS}}}tbl"
_TBLPR_TAG = f"{{{W_NS}}}tblPr"
_TBLGRID_TAG = f"{{{W_NS}}}tblGrid"
_GRIDCOL_TAG = f"{{{W_NS}}}gridCol"
_TBLW_TAG = f"{{{W_NS}}}tblW"
_JC_TAG = f"{{{W_NS}}}jc"
_IND_TAG = f"{{{W_NS}}}ind"
_TR_TAG = f"{{{W_NS}}}tr"
_TC_TAG = f"{{{W_NS}}}tc"
_TCPR_TAG = f"{{{W_NS}}}tcPr"
_GRIDSPAN_TAG = f"{{{W_NS}}}gridSpan"
_SHD_TAG = f"{{{W_NS}}}shd"
_P_TAG = f"{{{W_NS}}}p"
_PPR_TAG = f"{{{W_NS}}}pPr"
_PSTYLE_TAG = f"{{{W_NS}}}pStyle"
_R_TAG = f"{{{W_NS}}}r"
_T_TAG = f"{{{W_NS}}}t"
_FLDSIMPLE_TAG = f"{{{W_NS}}}fldSimple"
_INSTRTEXT_TAG = f"{{{W_NS}}}instrText"
_INSTR_ATTR = f"{{{W_NS}}}instr"
_W_ATTR = f"{{{W_NS}}}w"
_TYPE_ATTR = f"{{{W_NS}}}type"
_VAL_ATTR = f"{{{W_NS}}}val"
_FILL_ATTR = f"{{{W_NS}}}fill"
_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

# OOXML table width of 100% == 5000 fiftieths of a percent.
_MAX_PCT = 5000.0
# Reference text width in twips for turning an absolute (dxa) table width into a
# line fraction when no document says otherwise: Letter (8.5in) minus 1in margins
# each side = 6.5in = 9360 twips. Every part of a package takes the text width of
# the first section of its document, which LaTeX sets the PDF at.
_REFERENCE_WIDTH_TWIPS = 9360.0
# Map an OOXML <w:jc w:val> to a canonical alignment token.
_JC_TO_ALIGN = {"left": "left", "center": "center", "right": "right", "start": "left", "end": "right"}

# Paragraph style id that HTML->DOCX assigns to Polarion captions (see
# filters/html/html_captions.lua). Their text already carries the sequence number
# ("Table 1 ..."), so if pandoc's DOCX reader turns them into a LaTeX
# \caption it prepends its OWN "Table N:" counter, printing the number twice.
_CAPTION_STYLE_VAL = "Caption"
# The paragraph style pandoc's DOCX writer gives the paragraphs of a table cell.
_COMPACT_STYLE_VAL = "Compact"

# Default column width (twips) assigned when <w:gridCol> lacks a w:w
# attribute.  The exact value is unimportant — pandoc only uses the
# *ratios* between columns — so we pick a round number that gives a
# plausible-looking layout (~3.3 inches per column).
_DEFAULT_GRIDCOL_WIDTH = "4800"

HEX_COLOR_LENGTH = 6
_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX with table grid-column widths filled in and each
    styled table cell's background prepended as a sentinel run.

    Returns the input unchanged when there are no tables to fix or the
    package is not a recognisable DOCX.
    """
    entries = read_entries(docx_bytes)
    if entries is None:
        logger.warning("Input is not a valid zip / DOCX; skipping table preprocess")
        return docx_bytes

    changed = False
    width = docx_page_geometry.package_text_width_twips(entries)
    for part in enumerate_body_parts(entries.keys()):
        rewritten, part_changed = rewrite_part(entries[part], width)
        if part_changed:
            entries[part] = rewritten
            changed = True

    if not changed:
        return docx_bytes

    return repack(entries)


def _normalize_hex(value: str | None) -> str | None:
    """Uppercase a 6-digit hex string, or None for missing / unusable values.

    White (``FFFFFF``) is treated as "no background" because it is the page
    colour and wrapping every cell in ``\\cellcolor`` would be wasteful.
    """
    if not value:
        return None
    stripped = value.strip()
    if stripped.lower() == "auto":
        return None
    stripped = stripped.removeprefix("#")
    if len(stripped) == HEX_COLOR_LENGTH and all(c in _HEX_DIGITS for c in stripped):
        upper = stripped.upper()
        if upper == "FFFFFF":
            return None
        return upper
    return None


def _extract_cell_bg(tcpr: ET.Element) -> str | None:
    """Return the hex fill colour from ``<w:tcPr>/<w:shd>``, or None."""
    shd = tcpr.find(_SHD_TAG)
    if shd is None:
        return None
    return _normalize_hex(shd.get(_FILL_ATTR))


def _build_sentinel_text(kv: dict[str, str]) -> str:
    """Encode a property map as a ``<OPEN>k1=v1;k2=v2<CLOSE>`` sentinel string."""
    payload = ";".join(f"{key}={value}" for key, value in kv.items())
    return f"{SENTINEL_OPEN}{payload}{SENTINEL_CLOSE}"


def _parse_sentinel_text(text: str) -> tuple[dict[str, str], str]:
    """Split a leading sentinel into ``({k: v}, rest)``.

    Returns ``({}, text)`` when *text* does not start with a complete sentinel.
    """
    if not text.startswith(SENTINEL_OPEN):
        return {}, text
    close = text.find(SENTINEL_CLOSE)
    if close == -1:
        return {}, text
    payload = text[len(SENTINEL_OPEN) : close]
    rest = text[close + len(SENTINEL_CLOSE) :]
    kv: dict[str, str] = {}
    for segment in payload.split(";"):
        key, sep, value = segment.partition("=")
        if sep:
            kv[key] = value
    return kv, rest


def _first_run_text_element(para: ET.Element) -> ET.Element | None:
    """Return the ``<w:t>`` of the paragraph's first run (skipping ``<w:pPr>``),
    or None when the paragraph does not begin with a run.
    """
    for child in para:
        if child.tag == _R_TAG:
            return child.find(_T_TAG)
        if child.tag != _PPR_TAG:
            return None
    return None


def _make_sentinel_run(sentinel_text: str) -> ET.Element:
    """Build a ``<w:r><w:t>sentinel</w:t></w:r>`` element."""
    run = ET.Element(_R_TAG)
    text = ET.SubElement(run, _T_TAG, {_SPACE_ATTR: "preserve"})
    text.text = sentinel_text
    return run


def _is_positive_int(value: str | None) -> bool:
    """True when *value* is a positive integer string (e.g. a ``w:w`` width)."""
    if value is None:
        return False
    stripped = value.strip()
    return stripped.isdigit() and int(stripped) > 0


def _row_column_count(tr: ET.Element) -> int:
    """Number of grid columns a row spans (summing each cell's ``w:gridSpan``)."""
    total = 0
    for tc in tr.findall(_TC_TAG):
        span = 1
        tcpr = tc.find(_TCPR_TAG)
        if tcpr is not None:
            gridspan = tcpr.find(_GRIDSPAN_TAG)
            span_value = gridspan.get(_VAL_ATTR) if gridspan is not None else None
            if span_value is not None and _is_positive_int(span_value):
                span = int(span_value)
        total += span
    return total


def _table_column_count(tbl: ET.Element) -> int:
    """Widest row's column count — how many ``<w:gridCol>`` the table needs."""
    return max((_row_column_count(tr) for tr in tbl.findall(_TR_TAG)), default=0)


def _fix_grid_col_widths(tbl: ET.Element) -> bool:
    """Ensure the table has a ``<w:tblGrid>`` with one positive-width
    ``<w:gridCol>`` per column.

    Pandoc's DOCX reader derives ``ColSpec`` proportions from the grid columns.
    If the grid is missing, has too few columns, or carries missing/zero/invalid
    widths (some editors emit ``w:w="0"``), pandoc ends up with an empty or
    partial column list and the LaTeX writer renders the table content-width
    and centered — ignoring its ``<w:tblW>``. This normalises the grid so every
    column has a positive width (the ratios are all pandoc uses), which keeps
    the cells and lets the width/alignment recovery place the table correctly.

    Returns True when the grid was created or modified.
    """
    num_cols = _table_column_count(tbl)
    if num_cols == 0:
        return False

    changed = False
    tblgrid = tbl.find(_TBLGRID_TAG)
    if tblgrid is None:
        tblgrid = ET.Element(_TBLGRID_TAG)
        tblpr = tbl.find(_TBLPR_TAG)
        insert_at = list(tbl).index(tblpr) + 1 if tblpr is not None else 0
        tbl.insert(insert_at, tblgrid)
        changed = True

    grid_cols = tblgrid.findall(_GRIDCOL_TAG)
    for col in grid_cols:
        if not _is_positive_int(col.get(_W_ATTR)):
            col.set(_W_ATTR, _DEFAULT_GRIDCOL_WIDTH)
            changed = True

            # Add any missing columns so pandoc sees the full width, not a subset.
    for _ in range(len(grid_cols), num_cols):
        col = ET.SubElement(tblgrid, _GRIDCOL_TAG)
        col.set(_W_ATTR, _DEFAULT_GRIDCOL_WIDTH)
        changed = True

    return changed


def _own_first_para(tc: ET.Element) -> ET.Element | None:
    """Return the cell's own first ``<w:p>``, or None when it has none.

    Direct children only, and stopping at a nested ``<w:tbl>``, for two
    reasons. A recursive walk would reach into the inner table's cells, whose
    paragraphs belong to those cells and not to this one. And a cell holding a
    nested table has structure ``[tcPr, tbl, p]`` where the trailing ``<w:p/>``
    is the mandatory cell-mark paragraph: the sentinel cannot go there, because
    that puts it after the ``Table`` AST node.
    """
    for child in tc:
        if child.tag == _TBL_TAG:
            return None
        if child.tag == _P_TAG:
            return child
    return None


def _find_or_create_first_para(tc: ET.Element, tcpr: ET.Element | None) -> ET.Element:
    """Return the cell's own first ``<w:p>``, inserting an empty one if it has none."""
    para = _own_first_para(tc)
    if para is not None:
        return para
    para = ET.Element(_P_TAG)
    insert_pos = 1 if next(iter(tc), None) is tcpr else 0
    tc.insert(insert_pos, para)
    return para


def _ensure_sentinel(para: ET.Element, kv: dict[str, str]) -> None:
    """Merge *kv* into the paragraph's leading sentinel, creating one if absent.

    Merging (rather than prepending a second sentinel) keeps a single parseable
    ``<OPEN>...<CLOSE>`` at the very start, so table-level keys (``tw``/``ta``)
    and a cell background (``bg``) can share one sentinel on the first cell.
    Idempotent: re-applying the same keys leaves the text unchanged.
    """
    text_el = _first_run_text_element(para)
    if text_el is not None and text_el.text and text_el.text.startswith(SENTINEL_OPEN):
        existing, rest = _parse_sentinel_text(text_el.text)
        existing.update(kv)
        text_el.text = _build_sentinel_text(existing) + rest
        return

    ppr = para.find(_PPR_TAG)
    insert_at = 1 if ppr is not None and next(iter(para), None) is ppr else 0
    para.insert(insert_at, _make_sentinel_run(_build_sentinel_text(kv)))


# Word's justified alignments. "both" justifies every line but the last;
# "distribute" stretches the last line too. LaTeX has no separate mode for the
# latter, so both are carried as one value and render as ordinary justification.
_JUSTIFIED_JC_VALUES = frozenset({"both", "distribute"})


def _extract_cell_justification(tc: ET.Element) -> str | None:
    """Return "justify" when the cell's own first paragraph is justified.

    Reads only the cell's own paragraph: a cell that opens with a nested table
    has none, and taking the inner table's first paragraph would attribute an
    inner cell's alignment to the outer one - and then make the outer cell grow
    a paragraph purely to hold the sentinel.
    """
    first_para = _own_first_para(tc)
    if first_para is None:
        return None
    ppr = first_para.find(_PPR_TAG)
    if ppr is None:
        return None
    jc = ppr.find(_JC_TAG)
    if jc is None:
        return None
    return "justify" if jc.get(_VAL_ATTR) in _JUSTIFIED_JC_VALUES else None


def _paragraph_indent(para: ET.Element) -> int | None:
    """The positive left indent of a paragraph in twips, None for none.

    A list paragraph carries the indent of its level, which LaTeX sets for the list already.
    """
    ppr = para.find(_PPR_TAG)
    ind = ppr.find(_IND_TAG) if ppr is not None and ppr.find(f"{{{W_NS}}}numPr") is None else None
    if ind is None:
        return None
    twips = whole_number(ind.get(f"{{{W_NS}}}left") or ind.get(f"{{{W_NS}}}start"), signed=True)
    return twips if twips is not None and twips > 0 else None


def _extract_cell_indent(tc: ET.Element) -> str | None:
    """The left indent of the cell's own paragraphs in points, when those with text share a positive one.

    The indent applies to the whole cell, so a cell of paragraphs indented apart, or holding a table, gets none.
    """
    indents = set()
    for child in tc:
        if child.tag == _TBL_TAG:
            return None
        if child.tag == _P_TAG and any(text.text for text in child.iter(_T_TAG)):
            indents.add(_paragraph_indent(child))
    indent = indents.pop() if len(indents) == 1 else None
    return f"{indent / 20:g}" if indent is not None else None


def _cell_properties(tc: ET.Element, tcpr: ET.Element | None) -> dict[str, str]:
    """The per-cell sentinel keys this cell needs, empty when it needs none.

    Covers the background colour, justification and left indent; see the module docstring
    for why justification needs carrying and the other alignments do not.
    """
    props: dict[str, str] = {}

    if tcpr is not None:
        bg = _extract_cell_bg(tcpr)
        if bg:
            props["bg"] = bg

    justification = _extract_cell_justification(tc)
    if justification:
        props["ha"] = justification

    indent = _extract_cell_indent(tc)
    if indent:
        props["li"] = indent

    return props


def _already_tagged(para: ET.Element, props: dict[str, str]) -> bool:
    """True when the paragraph's leading sentinel already carries *props*.

    This is what keeps the pass idempotent: re-running it over its own output
    must not rewrite a sentinel that already says the same thing.
    """
    text_el = _first_run_text_element(para)
    if text_el is None or not text_el.text:
        return False
    existing, _ = _parse_sentinel_text(text_el.text)
    return all(existing.get(key) == value for key, value in props.items())


def _tag_cell_properties(tbl: ET.Element, alignments: Mapping[str, tuple[str | None, str | None]] | None = None) -> bool:
    """Prepend sentinels encoding the per-cell properties pandoc would drop.

    Returns True when any cell was modified.
    """
    changed = False
    # The lines of each cell of a grid whose cells draw lines of their own
    borders = _cell_borders(tbl) if _grid_lines(tbl) is not None else None
    head = _head_cell_alignments(tbl, alignments or {})

    for tc in tbl.iter(_TC_TAG):
        tcpr = tc.find(_TCPR_TAG)
        props = _cell_properties(tc, tcpr)
        if borders is not None and tc in borders:
            props["bd"] = borders[tc]
        if tc in head:
            props["va"] = head[tc]
        if not props:
            continue

        first_para = _find_or_create_first_para(tc, tcpr)
        if _already_tagged(first_para, props):
            continue

        _ensure_sentinel(first_para, props)
        changed = True

    return changed


def _extract_table_layout(tbl: ET.Element, text_width_twips: float) -> tuple[float | None, str | None, bool]:
    """Return ``(width_fraction, alignment, is_absolute)`` from ``<w:tblPr>``.

    ``width_fraction`` is the table's share of the line width: ``pct`` -> value
    / 5000; ``dxa`` -> twips / ``text_width_twips``; both clamped to 1.0. It is
    None when there is no usable ``<w:tblW>``. ``alignment`` is left/center/right
    from ``<w:jc>``, or None. ``is_absolute`` is True when the width came from a
    ``dxa`` (absolute) value — such tables are usually deliberately narrow, so
    the filter tightens their inter-column padding to keep them close to the
    requested size. These are the properties pandoc's DOCX reader discards (it
    keeps only column-relative widths and no table jc).
    """
    tblpr = tbl.find(_TBLPR_TAG)
    if tblpr is None:
        return None, None, False

    fraction: float | None = None
    is_absolute = False
    tblw = tblpr.find(_TBLW_TAG)
    if tblw is not None:
        raw = tblw.get(_W_ATTR)
        try:
            value = float(raw) if raw is not None else 0.0
        except ValueError:
            value = 0.0
        if value > 0:
            width_type = tblw.get(_TYPE_ATTR)
            if width_type == "pct":
                fraction = min(value / _MAX_PCT, 1.0)
            elif width_type == "dxa":
                fraction = min(value / text_width_twips, 1.0)
                is_absolute = True

    align: str | None = None
    jc = tblpr.find(_JC_TAG)
    if jc is not None:
        align = _JC_TO_ALIGN.get((jc.get(_VAL_ATTR) or "").lower())

    return fraction, align, is_absolute


def _first_own_cell(tbl: ET.Element) -> ET.Element | None:
    """Return this table's first cell (first ``<w:tc>`` of its first ``<w:tr>``),
    ignoring cells of nested tables.
    """
    for child in tbl:
        if child.tag == _TR_TAG:
            for cell in child:
                if cell.tag == _TC_TAG:
                    return cell
            return None
    return None


_BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
_NO_BORDER = frozenset({"nil", "none"})


def _draws_a_border(borders: ET.Element | None) -> bool:
    """Whether a <w:tblBorders> or <w:tcBorders> draws a line on any side."""
    if borders is None:
        return False
    return any((side := borders.find(f"{{{W_NS}}}{name}")) is not None and side.get(_VAL_ATTR) not in _NO_BORDER for name in _BORDER_SIDES)


def bordered_table_styles(styles: ET.Element | None) -> frozenset[str]:
    """The table styles whose own table properties draw borders, as a grid does."""
    if styles is None:
        return frozenset()
    found = set()
    for style in styles.iter(f"{{{W_NS}}}style"):
        if style.get(f"{{{W_NS}}}type") == "table" and _draws_a_border(style.find(f"{_TBLPR_TAG}/{{{W_NS}}}tblBorders")):
            found.add(style.get(f"{{{W_NS}}}styleId") or "")
    return frozenset(found)


# A side of the first cell and the border of the table it falls back to.
_GRID_SIDES = (("top", "top"), ("left", "left"), ("bottom", "insideH"), ("right", "insideV"))
# The width of a border is in eighths of a point; Word draws one of no width at a quarter point.
_EIGHTHS_PER_POINT = 8
_THINNEST_LINE = 0.25
_THICKEST_LINE = 6.0


def _grid_sides(tbl: ET.Element) -> list[ET.Element] | None:
    """The borders of every side of the first cell, its own else the table's; None when a side draws none.

    The first cell stands for the table.
    """
    tblpr = tbl.find(_TBLPR_TAG)
    table_borders = tblpr.find(f"{{{W_NS}}}tblBorders") if tblpr is not None else None
    first = _first_own_cell(tbl)
    tcpr = first.find(_TCPR_TAG) if first is not None else None
    cell_borders = tcpr.find(f"{{{W_NS}}}tcBorders") if tcpr is not None else None
    sides = []
    for cell_side, table_side in _GRID_SIDES:
        side = cell_borders.find(f"{{{W_NS}}}{cell_side}") if cell_borders is not None else None
        if side is None and table_borders is not None:
            side = table_borders.find(f"{{{W_NS}}}{table_side}")
        if side is None or side.get(_VAL_ATTR) in _NO_BORDER:
            return None
        sides.append(side)
    return sides


def _line_width(side: ET.Element) -> float:
    """The width of a border in points, within what a PDF line shows."""
    try:
        width = int(side.get(f"{{{W_NS}}}sz") or "4") / _EIGHTHS_PER_POINT
    except ValueError:
        width = 4 / _EIGHTHS_PER_POINT
    return min(max(width, _THINNEST_LINE), _THICKEST_LINE)


def _grid_lines(tbl: ET.Element) -> tuple[str, float] | None:
    """The colour and width of the lines of a table which draws every side of its cells, else None."""
    sides = _grid_sides(tbl)
    if sides is None:
        return None
    raw_color = (sides[0].get(f"{{{W_NS}}}color") or "").upper()
    if raw_color == "FFFFFF":
        # White lines on white paper: nothing to draw
        return None
    return _normalize_hex(raw_color) or "000000", _line_width(sides[0])


# The styles of a border which LaTeX draws as Word does; any other draws as a single line.
_BORDER_STYLES = frozenset({"single", "dashed", "dotted", "double"})
_VMERGE_TAG = f"{{{W_NS}}}vMerge"


def _side_value(side: ET.Element | None) -> str:
    """A border as the filter reads it: ``style.eighths.RRGGBB``, or ``-`` for none."""
    if side is None or side.get(_VAL_ATTR) in _NO_BORDER:
        return "-"
    style = side.get(_VAL_ATTR) if side.get(_VAL_ATTR) in _BORDER_STYLES else "single"
    try:
        eighths = int(side.get(f"{{{W_NS}}}sz") or "4")
    except ValueError:
        eighths = 4
    eighths = min(max(eighths, 2), 48)
    color = _normalize_hex(side.get(f"{{{W_NS}}}color")) or ("FFFFFF" if (side.get(f"{{{W_NS}}}color") or "").upper() == "FFFFFF" else "000000")
    return f"{style}.{eighths}.{color}"


def _span(tc: ET.Element) -> int:
    gridspan = tc.find(f"{_TCPR_TAG}/{_GRIDSPAN_TAG}")
    value = gridspan.get(_VAL_ATTR) if gridspan is not None else None
    return int(value) if value is not None and _is_positive_int(value) else 1


def _continues_merge(tc: ET.Element) -> bool:
    vmerge = tc.find(f"{_TCPR_TAG}/{_VMERGE_TAG}")
    return vmerge is not None and vmerge.get(_VAL_ATTR) in (None, "continue")


def _kept_cells(rows: list[ET.Element]) -> list[list]:
    """Each cell pandoc keeps, with its row and column, its span and the cell and row it ends with."""
    cells: list[list] = []
    open_merges: dict[int, list] = {}
    for index, tr in enumerate(rows):
        column = 0
        for tc in tr.findall(_TC_TAG):
            span = _span(tc)
            if _continues_merge(tc) and column in open_merges:
                open_merges[column][4:6] = [tc, index]
            else:
                entry = [tc, index, column, span, tc, index]
                cells.append(entry)
                open_merges[column] = entry
            column += span
    return cells


def _cell_borders(tbl: ET.Element) -> dict[ET.Element, str] | None:
    """The sides of each cell of a grid, top, right, bottom and left, when they are not all alike.

    A side is the cell's own border, else the table's: its outer border at the
    edge of the table, the inside one between two cells. A cell merged down to
    later rows takes its bottom from the last of them, the cell pandoc keeps.
    """
    tblpr = tbl.find(_TBLPR_TAG)
    table_borders = tblpr.find(f"{{{W_NS}}}tblBorders") if tblpr is not None else None
    rows = tbl.findall(_TR_TAG)
    columns = _table_column_count(tbl)
    cells = _kept_cells(rows)

    def side(tc: ET.Element, name: str, inner: str, edge: bool) -> str:
        own = tc.find(f"{_TCPR_TAG}/{{{W_NS}}}tcBorders/{{{W_NS}}}{name}")
        if own is None and table_borders is not None:
            own = table_borders.find(f"{{{W_NS}}}{name if edge else inner}")
        return _side_value(own)

    sides: dict[ET.Element, str] = {}
    for tc, row, column, span, last_tc, last_row in cells:
        sides[tc] = "/".join(
            (
                side(tc, "top", "insideH", row == 0),
                side(tc, "right", "insideV", column + span >= columns),
                side(last_tc, "bottom", "insideH", last_row == len(rows) - 1),
                side(tc, "left", "insideV", column == 0),
            ),
        )
    if len({border for value in sides.values() for border in value.split("/")}) <= 1:
        # One border on every side of every cell: the lines of the table say it all
        return None
    return sides


# The bit of the hexadecimal w:val of <w:tblLook> which turns the formatting of the first row on.
_FIRST_ROW_BIT = 0x0020


def _look_bits(value: str | None) -> int:
    try:
        return int(value or "0", 16)
    except ValueError:
        return 0


def _first_row_look(tblpr: ET.Element | None) -> bool:
    """Whether the table turns the formatting of its first row on, which also makes it the header for pandoc."""
    look = tblpr.find(f"{{{W_NS}}}tblLook") if tblpr is not None else None
    return look is not None and (look.get(f"{{{W_NS}}}firstRow") in ("1", "true", "on") or _look_bits(look.get(_VAL_ATTR)) & _FIRST_ROW_BIT != 0)


# The vertical alignment of a cell, as the key the filter reads.
_VERTICAL_ALIGNMENTS = {"top": "t", "center": "c", "bottom": "b"}


def _vertical_alignment(tcpr: ET.Element | None) -> str | None:
    v_align = tcpr.find(f"{{{W_NS}}}vAlign") if tcpr is not None else None
    return _VERTICAL_ALIGNMENTS.get(v_align.get(_VAL_ATTR) or "") if v_align is not None else None


def head_alignments(styles: ET.Element | None) -> dict[str, tuple[str | None, str | None]]:
    """The vertical alignment each table style gives the cells of the first row, and the cells of the table.

    A style without one of its own takes that of the style it is based on.
    """
    if styles is None:
        return {}
    by_id = {style.get(f"{{{W_NS}}}styleId") or "": style for style in styles.iter(f"{{{W_NS}}}style") if style.get(f"{{{W_NS}}}type") == "table"}

    def own(style: ET.Element) -> tuple[str | None, str | None]:
        first_row = None
        for conditional in style.findall(f"{{{W_NS}}}tblStylePr"):
            if conditional.get(f"{{{W_NS}}}type") == "firstRow":
                first_row = _vertical_alignment(conditional.find(_TCPR_TAG))
        return first_row, _vertical_alignment(style.find(_TCPR_TAG))

    def resolve(style_id: str, seen: frozenset[str]) -> tuple[str | None, str | None]:
        style = by_id.get(style_id)
        if style is None or style_id in seen:
            return None, None
        first_row, table = own(style)
        based_on = style.find(f"{{{W_NS}}}basedOn")
        if based_on is not None and (first_row is None or table is None):
            base_first_row, base_table = resolve(based_on.get(_VAL_ATTR) or "", seen | {style_id})
            first_row, table = first_row or base_first_row, table or base_table
        return first_row, table

    found = {style_id: resolve(style_id, frozenset()) for style_id in by_id}
    # A table without a style of its own takes the default table style, under the empty name
    default = next((style_id for style_id, style in by_id.items() if style.get(f"{{{W_NS}}}default") in ("1", "true", "on")), None)
    if default is not None:
        found[""] = found[default]
    return found


class TableStyles(NamedTuple):
    """What the table styles of a document give its tables: borders, and the vertical alignment of cells."""

    bordered: frozenset[str] = frozenset()
    alignments: Mapping[str, tuple[str | None, str | None]] = MappingProxyType({})


def table_styles(styles: ET.Element | None) -> TableStyles:
    return TableStyles(bordered_table_styles(styles), head_alignments(styles))


def _head_cell_alignments(tbl: ET.Element, alignments: Mapping[str, tuple[str | None, str | None]]) -> dict[ET.Element, str]:
    """The vertical alignment of each cell of the header rows: of the cell, else of its table style, else the top.

    The header rows are those pandoc makes the header of: the rows which repeat on each page, and the first
    row of a table which turns its formatting on. Only the first row takes the formatting of the first row.
    """
    tblpr = tbl.find(_TBLPR_TAG)
    first_row = _first_row_look(tblpr)
    rows = tbl.findall(_TR_TAG)
    count = 0
    while count < len(rows) and rows[count].find(f"{{{W_NS}}}trPr/{{{W_NS}}}tblHeader") is not None:
        count += 1
    count = max(count, 1 if first_row and rows else 0)
    style = tblpr.find(f"{{{W_NS}}}tblStyle") if tblpr is not None else None
    style_first_row, style_table = alignments.get(style.get(_VAL_ATTR) or "" if style is not None else "", (None, None))
    found = {}
    for index, row in enumerate(rows[:count]):
        for tc in row.findall(_TC_TAG):
            found[tc] = _vertical_alignment(tc.find(_TCPR_TAG)) or (style_first_row if index == 0 and first_row else None) or style_table or "t"
    return found


def _rules(tbl: ET.Element, bordered_styles: frozenset[str]) -> str:
    """The lines Word draws in a table: ``grid`` for borders of its own, of its cells or of its style,
    ``head`` for the line under the header row of pandoc's Table style, ``none`` without one.
    """
    tblpr = tbl.find(_TBLPR_TAG)
    style = tblpr.find(f"{{{W_NS}}}tblStyle") if tblpr is not None else None
    if (
        _draws_a_border(tblpr.find(f"{{{W_NS}}}tblBorders") if tblpr is not None else None)
        or (style is not None and style.get(_VAL_ATTR) in bordered_styles)
        or any(_draws_a_border(cell.find(f"{{{W_NS}}}tcBorders")) for cell in tbl.iter(_TCPR_TAG))
    ):
        return "grid"
    first_row = _first_row_look(tblpr)
    has_header = any(row.find(f"{{{W_NS}}}trPr/{{{W_NS}}}tblHeader") is not None for row in tbl.findall(f"{{{W_NS}}}tr"))
    return "head" if first_row and has_header else "none"


def _tag_table_layout(tbl: ET.Element, text_width_twips: float, bordered_styles: frozenset[str] = frozenset()) -> bool:
    """Tag the first cell with the table's width fraction and/or alignment.

    Every table with an explicit ``<w:tblW>`` is tagged, **including 100 %**:
    pandoc's DOCX reader may still emit a table with no usable column widths,
    which the LaTeX writer then renders content-width (and, with no alignment
    glue, centered) instead of filling the line. The companion filter forces
    the recovered fraction onto the columns and re-applies the alignment, so a
    100 % table fills the text width flush-left rather than floating centered.
    (This matches how the HTML path in filters/html/html_tables_to_latex.lua already
    handles full-width tables.) Returns True when a sentinel was written.
    """
    fraction, align, is_absolute = _extract_table_layout(tbl, text_width_twips)
    kv: dict[str, str] = {}
    if fraction is not None:
        kv["tw"] = f"{min(fraction, 1.0):.4f}"
        if is_absolute:
            # Absolute (px/pt) widths are usually deliberately narrow; flag them
            # so the filter tightens inter-column padding, which otherwise adds
            # a fixed minimum that leaves a small table wider than requested.
            kv["aw"] = "1"
    if align is not None:
        kv["ta"] = align
    # The lines of a styled table: LaTeX draws booktabs rules around every table, Word only the borders it
    # has. A table with borders keeps LaTeX's rules, and one without a style is left as it is.
    # A table which draws every side of its cells is set as a grid in their colour and width.
    tblpr = tbl.find(_TBLPR_TAG)
    grid = _grid_lines(tbl)
    if grid is not None:
        kv["rl"], kv["lc"], kv["lw"] = "grid", grid[0], f"{grid[1]:g}"
    elif tblpr is not None and tblpr.find(f"{{{W_NS}}}tblStyle") is not None:
        rules = _rules(tbl, bordered_styles)
        if rules != "grid":
            kv["rl"] = rules
    first_tc = _first_own_cell(tbl)
    if first_tc is None:
        return False
    # A cell paragraph of the Compact style pandoc writes: its rows take the space of Compact, others of Normal
    own_para = _own_first_para(first_tc)
    own_style = own_para.find(f"{_PPR_TAG}/{_PSTYLE_TAG}") if own_para is not None else None
    if own_style is not None and own_style.get(_VAL_ATTR) == _COMPACT_STYLE_VAL:
        kv["cc"] = "1"
    if not kv:
        return False

    first_para = _find_or_create_first_para(first_tc, first_tc.find(_TCPR_TAG))
    _ensure_sentinel(first_para, kv)
    return True


def _neutralize_caption_paragraphs(tree: ET.Element) -> bool:
    """Give every ``Caption`` paragraph the style ``PandocCaption`` so pandoc keeps it where it is.

    Left as ``Caption``, pandoc's DOCX reader attaches the paragraph to the
    adjacent table or picture: LaTeX numbers it again ("Table N: Table N ...")
    or docx_caption_labels_to_latex moves it above the table, though Word
    prints it below. A caption carries its number as text, literally for
    Polarion and as the cached result of a ``SEQ`` field for Word, so it needs
    no counter of LaTeX; and the lists of figures and tables come from the
    entries of their fields (see app/docx_toc_pre_process.py). A style of its
    own keeps the look of the Caption style, which
    filters/docx/docx_document_styles_to_latex.lua sets. pandoc's own
    ``TableCaption``/``ImageCaption`` styles are left untouched.

    Returns True when any paragraph was changed.
    """
    changed = False
    for para in tree.iter(_P_TAG):
        ppr = para.find(_PPR_TAG)
        if ppr is None:
            continue
        pstyle = ppr.find(_PSTYLE_TAG)
        if pstyle is not None and pstyle.get(_VAL_ATTR) == _CAPTION_STYLE_VAL:
            pstyle.set(_VAL_ATTR, PLAIN_CAPTION_STYLE)
            changed = True
    return changed


def rewrite_part(
    xml_bytes: bytes,
    text_width_twips: float | None = None,
    bordered_styles: frozenset[str] = frozenset(),
    alignments: Mapping[str, tuple[str | None, str | None]] | None = None,
) -> tuple[bytes, bool]:
    """Fix grid-column widths, tag styled table cells, carry table width/
    alignment, and neutralise caption styles in one body part.

    Returns ``(new_bytes, changed)``.
    """
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping table preprocess")
        return xml_bytes, False

    changed = False
    # An absolute table width is a share of the line LaTeX sets the table at: the text of its section.
    # Outside the body that is the text of the first section of the document, which the caller reads
    # once for every part, the footnotes among them. Called on its own, a part with a body reads it,
    # and any other falls back to Letter.
    if text_width_twips is not None:
        width = text_width_twips
    elif tree.find(f"{{{W_NS}}}body") is not None:
        width = docx_page_geometry.text_width_twips(tree)
    else:
        width = _REFERENCE_WIDTH_TWIPS
    section_widths = docx_page_geometry.table_text_widths_twips(tree)

    for tbl in tree.iter(_TBL_TAG):
        if _fix_grid_col_widths(tbl):
            changed = True
        if _tag_table_layout(tbl, section_widths.get(tbl, width), bordered_styles):
            changed = True
        if _tag_cell_properties(tbl, alignments):
            changed = True

    if _neutralize_caption_paragraphs(tree):
        changed = True

    if not changed:
        return xml_bytes, False

    return serialize_tree(tree), True
