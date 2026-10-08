"""Decide the column widths of a DOCX table before the images in its cells are fitted to them.

pandoc keeps the column widths of an HTML table only when a ``<colgroup>`` states them in percent.
Otherwise it splits its text width evenly across the columns. Word lays out such a table by what
its cells hold, so a column with a picture comes out much wider than its even share, and a picture
fitted to that share leaves most of its cell empty.

This module decides the widths the way a browser lays out an automatic table: from the widths the
HTML states for the columns, and from what each column holds where it states none. The post-processor
writes the result into the grid and into each cell, so Word lays the table out the same way.
"""

import sys
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from collections.abc import Callable, Sequence

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"  # NOSONAR
_W = f"{{{SCHEMA}}}"

EMU_PER_TWIP = 635
EMU_PER_POINT = 12700
# OOXML states a percentage in fiftieths of a percent: 100 % is 5000.
FULL_PCT = 5000
# The narrowest a picture is laid out when its column is short of space: one inch.
MIN_IMAGE_WIDTH_EMU = 914400
# An average character of a proportional font is about half its size wide.
CHARACTER_WIDTH_EM = 0.5
# Word's default tab stop: half an inch.
TAB_WIDTH_EMU = 457200
# Word's font size where neither the text nor its styles state one: 10 pt, in half-points.
DEFAULT_FONT_SIZE_HALF_POINTS = 20

# A stated column width: ("pct", fiftieths of a percent) or ("dxa", twips).
type ColumnWidth = tuple[str, int]


def decide(tbl: Any, table_width: int, stated: Sequence[ColumnWidth | None] | None, cell_margins: Callable[[Any], int], styles: Any = None) -> list[int] | None:
    """The width of each column of the table, in EMU, adding up to `table_width`.

    `stated` holds the width the HTML gives each column, or None for one it leaves to its content.
    `cell_margins` gives the left plus right margin of a cell, in EMU. None when the table has no
    column to lay out.
    """
    columns = _column_count(tbl)
    if columns == 0 or table_width <= 0:
        return None
    stated_emu = _stated_emu(stated, columns, table_width)
    if all(width is not None for width in stated_emu):
        return _scale([width or 0 for width in stated_emu], table_width)
    minimums, maximums = column_extents(tbl, columns, cell_margins, _FontSizes(styles))
    fixed = [width is not None for width in stated_emu]
    for index, width in enumerate(stated_emu):
        if width is not None:
            minimums[index] = maximums[index] = width
    return distribute(minimums, maximums, table_width, fixed)


def distribute(minimums: list[int], maximums: list[int], width: int, fixed: list[bool]) -> list[int]:
    """Share `width` between columns that need at least `minimums` and would take up to `maximums`.

    This is how a browser lays out an automatic table. With room for every column at its widest,
    the space left goes to the columns without a stated width, in proportion to their widest.
    Without that room, each column gets its narrowest plus a share of the rest in proportion to
    how much wider it would be. Without room even for the narrowest, all are scaled down alike.
    """
    if sum(maximums) <= width:
        flexible = [maximum if not stated else 0 for maximum, stated in zip(maximums, fixed, strict=True)]
        weights = flexible if sum(flexible) > 0 else maximums
        if sum(weights) <= 0:
            return _scale([1] * len(maximums), width)
        extra = width - sum(maximums)
        return [maximum + extra * weight // sum(weights) for maximum, weight in zip(maximums, weights, strict=True)]
    if sum(minimums) < width:
        growth = [maximum - minimum for minimum, maximum in zip(minimums, maximums, strict=True)]
        spare = width - sum(minimums)
        return [minimum + spare * grow // sum(growth) for minimum, grow in zip(minimums, growth, strict=True)]
    return _scale(minimums, width)


def column_extents(tbl: Any, columns: int, cell_margins: Callable[[Any], int], font_sizes: _FontSizes) -> tuple[list[int], list[int]]:
    """The narrowest and the widest each column's content can be laid out at, in EMU, margins included."""
    minimums = [0] * columns
    maximums = [0] * columns
    spanning: list[tuple[int, int, int, int]] = []
    for offset, span, tc in cells(tbl, columns):
        content_min, content_max = _content_extent(tc, font_sizes, cell_margins)
        margins = cell_margins(tc)
        cell_min, cell_max = content_min + margins, content_max + margins
        if span == 1:
            minimums[offset] = max(minimums[offset], cell_min)
            maximums[offset] = max(maximums[offset], cell_max)
        else:
            spanning.append((offset, span, cell_min, cell_max))
    # A cell spanning columns widens them only by what they lack for it, shared evenly.
    for offset, span, cell_min, cell_max in spanning:
        _widen(minimums, offset, span, cell_min)
        _widen(maximums, offset, span, cell_max)
    return minimums, [max(minimum, maximum) for minimum, maximum in zip(minimums, maximums, strict=True)]


def _widen(widths: list[int], offset: int, span: int, needed: int) -> None:
    lacking = needed - sum(widths[offset : offset + span])
    if lacking > 0:
        for index in range(offset, offset + span):
            widths[index] += lacking // span


def _column_count(tbl: Any) -> int:
    grid = tbl.findall(f"{_W}tblGrid/{_W}gridCol")
    if grid:
        return len(grid)
    return max((offset + span for offset, span, _ in cells(tbl, sys.maxsize)), default=0)


def cells(tbl: Any, columns: int) -> list[tuple[int, int, Any]]:
    """Each cell of the table's own rows with the grid column it starts at and the columns it spans."""
    cells = []
    for tr in tbl.findall(f"{_W}tr"):
        offset = _int_value(tr.find(f"{_W}trPr/{_W}gridBefore"), 0)
        for tc in tr.findall(f"{_W}tc"):
            span = max(_int_value(tc.find(f"{_W}tcPr/{_W}gridSpan"), 1), 1)
            if offset < columns:
                cells.append((offset, min(span, columns - offset), tc))
            offset += span
    return cells


def _stated_emu(stated: Sequence[ColumnWidth | None] | None, columns: int, table_width: int) -> list[int | None]:
    """The stated widths in EMU. A list that does not match the grid states nothing."""
    if stated is None or len(stated) != columns:
        return [None] * columns
    widths: list[int | None] = []
    for width in stated:
        if width is None:
            widths.append(None)
        elif width[0] == "pct":
            widths.append(table_width * width[1] // FULL_PCT)
        else:
            widths.append(width[1] * EMU_PER_TWIP)
    return widths


def _scale(widths: list[int], width: int) -> list[int]:
    total = sum(widths)
    if total <= 0:
        return [width // len(widths)] * len(widths)
    return [part * width // total for part in widths]


def _content_extent(container: Any, font_sizes: _FontSizes, cell_margins: Callable[[Any], int]) -> tuple[int, int]:
    """The narrowest and the widest the blocks of a cell can be laid out at, in EMU."""
    minimum = maximum = 0
    for child in container:
        if child.tag == f"{_W}p":
            block_min, block_max = _paragraph_extent(child, font_sizes)
        elif child.tag == f"{_W}tbl":
            # A nested table is as narrow and as wide as its columns together.
            nested_min, nested_max = column_extents(child, _column_count(child), cell_margins, font_sizes)
            block_min, block_max = sum(nested_min), sum(nested_max)
        elif child.tag in (f"{_W}sdt", f"{_W}sdtContent", f"{_W}customXml"):
            block_min, block_max = _content_extent(child.find(f"{_W}sdtContent") if child.tag == f"{_W}sdt" else child, font_sizes, cell_margins)
        else:
            continue
        minimum = max(minimum, block_min)
        maximum = max(maximum, block_max)
    return minimum, maximum


def _paragraph_extent(p: Any, font_sizes: _FontSizes) -> tuple[int, int]:
    """The longest word or picture of a paragraph, and its longest line, in EMU."""
    paragraph_size = font_sizes.paragraph(p)
    minimum = line = longest = 0
    for element in p.iter(f"{_W}t", f"{_W}br", f"{_W}cr", f"{_W}tab", f"{{{WP_NS}}}extent"):
        if element.tag == f"{{{WP_NS}}}extent":
            width = _int_attribute(element, "cx")
            minimum = max(minimum, min(width, MIN_IMAGE_WIDTH_EMU))
            line += width
        elif element.tag == f"{_W}t":
            character = font_sizes.run(element.getparent(), paragraph_size) * EMU_PER_POINT * CHARACTER_WIDTH_EM / 2
            text = element.text or ""
            minimum = max(minimum, *(round(len(word) * character) for word in text.split()), 0)
            line += round(len(text) * character)
        elif element.tag == f"{_W}tab":
            line += TAB_WIDTH_EMU
        else:
            longest = max(longest, line)
            line = 0
    return minimum, max(longest, line)


class _FontSizes:
    """The font size text is set in, in half-points: its run's, else its paragraph style's, else the document's."""

    def __init__(self, styles: Any) -> None:
        self._styles = {style.get(f"{_W}styleId"): style for style in styles.findall(f"{_W}style")} if styles is not None else {}
        default_size = styles.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr/{_W}sz") if styles is not None else None
        self._default = _int_value(default_size, DEFAULT_FONT_SIZE_HALF_POINTS)
        self._default_paragraph_style = next(
            (style_id for style_id, style in self._styles.items() if style.get(f"{_W}type") == "paragraph" and style.get(f"{_W}default") in ("1", "true")),
            None,
        )
        self._cache: dict[str | None, int] = {}

    def paragraph(self, p: Any) -> int:
        reference = p.find(f"{_W}pPr/{_W}pStyle")
        return self._style_size(reference.get(f"{_W}val") if reference is not None else self._default_paragraph_style)

    def run(self, run: Any, paragraph_size: int) -> int:
        size = run.find(f"{_W}rPr/{_W}sz") if run is not None else None
        return _int_value(size, paragraph_size)

    def _style_size(self, style_id: str | None) -> int:
        if style_id not in self._cache:
            self._cache[style_id] = self._resolve(style_id)
        return self._cache[style_id]

    def _resolve(self, style_id: str | None) -> int:
        seen: set[str] = set()
        while style_id is not None and style_id not in seen:
            seen.add(style_id)
            style = self._styles.get(style_id)
            if style is None:
                break
            size = style.find(f"{_W}rPr/{_W}sz")
            if size is not None:
                return _int_value(size, self._default)
            based_on = style.find(f"{_W}basedOn")
            style_id = based_on.get(f"{_W}val") if based_on is not None else None
        return self._default


def _int_value(element: Any, fallback: int) -> int:
    """The whole-number w:val of an element, or the fallback where it has none."""
    value = element.get(f"{_W}val") if element is not None else None
    return int(value) if value is not None and value.isdigit() else fallback


def _int_attribute(element: Any, name: str) -> int:
    value = element.get(name)
    return int(value) if value is not None and value.isdigit() else 0
