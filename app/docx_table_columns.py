"""Decide the column widths of a DOCX table before the images in its cells are fitted to them.

pandoc keeps the column widths of an HTML table only when a ``<colgroup>`` states them in percent.
Otherwise it splits its text width evenly across the columns. Word lays out such a table by what
its cells hold, so a column with a picture comes out much wider than its even share, and a picture
fitted to that share leaves most of its cell empty.

This module decides the widths the way a browser lays out an automatic table: from the widths the
HTML states for the columns, and from what each column holds where it states none. The post-processor
writes the result into the grid and into each cell, so Word lays the table out the same way.
"""

import re
import unicodedata
from typing import TYPE_CHECKING, Any, NamedTuple

from app.docx_numbers import whole_number

if TYPE_CHECKING:
    from collections.abc import Callable, Sequence

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"  # NOSONAR
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"  # NOSONAR
_W = f"{{{SCHEMA}}}"
_M = f"{{{MATH_NS}}}"

EMU_PER_TWIP = 635
EMU_PER_POINT = 12700
# OOXML states a percentage in fiftieths of a percent: 100 % is 5000.
FULL_PCT = 5000
# An average character of a proportional font is about half its size wide.
CHARACTER_WIDTH_EM = 0.5
# The narrowest a picture is lowered to when the table has no room for it at its width: an inch.
# A column of text keeps its longest word whole before that.
CAPPED_PICTURE_WIDTH_EMU = 914400
# Word's default tab stop: half an inch.
TAB_WIDTH_EMU = 457200
# Word's font size where neither the text nor its styles state one: 10 pt, in half-points.
DEFAULT_FONT_SIZE_HALF_POINTS = 20

# How wide the characters of an equation are, in ems of its size, about as Word sets them in Cambria
# Math: a letter or a digit, a bracket or a punctuation mark, a space, and a sign between two terms
# with the space Word puts around it.
MATH_LETTER_EM = 0.55
MATH_BRACKET_EM = 0.4
MATH_SPACE_EM = 0.25
MATH_SIGN_EM = 1.2
# The sign of a root, and the sign of a sum or an integral.
MATH_RADICAL_EM = 0.8
MATH_NARY_EM = 1.3
# An index, a limit and the like are set at about 0.7 of the size.
MATH_SCRIPT_SCALE = 0.7
# The columns of a matrix stand an em apart.
MATH_MATRIX_GAP_EM = 1.0
_MATH_BRACKETS = frozenset("()[]{}|,.;:!'\"\u2032\u2033\u2034")
# Characters which take no room: format characters such as a zero-width space, and combining marks.
_NO_WIDTH_CATEGORIES = frozenset({"Cf", "Mn", "Me"})

# What ends a word in the text of a run: whitespace a line may break at, not a no-break space
_WHITESPACE = re.compile(r"[^\S\u00a0\u2007\u202f]+")

# The most columns a table without a grid is laid out with: the most a colspan covers in HTML.
MAX_COLUMNS = 1000


class Extents(NamedTuple):
    """How narrow and how wide the columns of a table can be laid out, in EMU, margins included.

    ``minimums`` and ``maximums`` count a cell spanning columns in, its need shared evenly by them.
    ``own_minimums`` and ``floors`` hold the cells of one column only: their narrowest, and their
    narrowest with the pictures lowered to an inch, below which their words would break. Each of
    ``spans`` is a cell spanning columns: its first column, the count, its narrowest and its floor.
    """

    minimums: list[int]
    maximums: list[int]
    own_minimums: list[int]
    floors: list[int]
    spans: list[tuple[int, int, int, int]]


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
    extents = column_extents(tbl, columns, cell_margins, _FontSizes(styles))
    minimums, maximums = extents.minimums, extents.maximums
    fixed = [width is not None for width in stated_emu]
    for index, width in enumerate(stated_emu):
        if width is not None:
            minimums[index] = maximums[index] = width
    # The columns without a stated width share what the stated ones leave
    free = [index for index in range(columns) if not fixed[index]]
    room = max(table_width - sum(width or 0 for width in stated_emu), 0)
    position = {index: at for at, index in enumerate(free)}
    free_spans: list[tuple[list[int], int, int]] = []
    for offset, span, needed, floor in extents.spans:
        covered = range(offset, offset + span)
        # What a stated column under a spanning cell gives it, the free ones need not
        given = sum(stated_emu[index] or 0 for index in covered if fixed[index])
        positions = [position[index] for index in covered if not fixed[index]]
        if positions and needed > given:
            free_spans.append((positions, needed - given, max(floor - given, 0)))
    capped = _cap(
        [minimums[index] for index in free],
        room,
        [extents.own_minimums[index] for index in free],
        [extents.floors[index] for index in free],
        free_spans,
    )
    for index, minimum in zip(free, capped, strict=True):
        minimums[index] = minimum
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


def column_extents(tbl: Any, columns: int, cell_margins: Callable[[Any], int], font_sizes: _FontSizes) -> Extents:
    """How narrow and how wide each column's content can be laid out, in EMU, margins included. See Extents."""
    minimums = [0] * columns
    maximums = [0] * columns
    floors = [0] * columns
    spanning: list[tuple[int, int, int, int, int]] = []
    for offset, span, tc in cells(tbl, columns):
        content_min, content_max, content_floor = _content_extent(tc, font_sizes, cell_margins)
        margins = cell_margins(tc)
        cell_min, cell_max, cell_floor = content_min + margins, content_max + margins, content_floor + margins
        if span == 1:
            minimums[offset] = max(minimums[offset], cell_min)
            maximums[offset] = max(maximums[offset], cell_max)
            floors[offset] = max(floors[offset], cell_floor)
        else:
            spanning.append((offset, span, cell_min, cell_max, cell_floor))
    # A cell spanning columns widens them only by what they lack for it, shared evenly.
    own_minimums = list(minimums)
    for offset, span, cell_min, cell_max, _cell_floor in spanning:
        _widen(minimums, offset, span, cell_min)
        _widen(maximums, offset, span, cell_max)
    widest = [max(minimum, maximum) for minimum, maximum in zip(minimums, maximums, strict=True)]
    spans = [(offset, span, cell_min, cell_floor) for offset, span, cell_min, _cell_max, cell_floor in spanning]
    return Extents(minimums, widest, own_minimums, floors, spans)


def _widen(widths: list[int], offset: int, span: int, needed: int) -> None:
    lacking = needed - sum(widths[offset : offset + span])
    if lacking > 0:
        for index in range(offset, offset + span):
            widths[index] += lacking // span


def _column_count(tbl: Any) -> int:
    """The columns of the grid; without one, those the cells span, at most MAX_COLUMNS.

    A table without a grid takes its count from gridSpan and gridBefore values, and a short element
    stating a huge one must not make the post-processor build a list that long.
    """
    grid = tbl.findall(f"{_W}tblGrid/{_W}gridCol")
    if grid:
        return len(grid)
    return max((offset + span for offset, span, _ in cells(tbl, MAX_COLUMNS)), default=0)


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


def _cap(minimums: list[int], width: int, own_minimums: list[int], floors: list[int], spans: list[tuple[list[int], int, int]]) -> list[int]:
    """Lower the widest columns to one level until they fit `width`, the narrow ones left whole.

    `minimums` count a cell spanning columns in, shared evenly: where they fit, they are kept. Else
    each column is held to a level, between its floor and its own narrowest: a picture wider than
    the table gives up what the table lacks, and the text beside it keeps its words whole. A cell
    spanning columns, given as the positions it covers, its narrowest and its floor, is held to the
    same level, and what it then lacks goes on top. The level is the highest at which all of that
    fits. Where even the floors do not fit, as beside stated widths filling the table, the columns
    keep them and are scaled down with the others, as a picture always was.
    """
    if sum(minimums) <= width:
        return minimums

    def fitted(level: int) -> list[int]:
        held = [max(floor, min(minimum, level)) for floor, minimum in zip(floors, own_minimums, strict=True)]
        return _raise_for_spans(held, [(positions, max(floor, min(needed, level))) for positions, needed, floor in spans])

    low, high = 0, max([*own_minimums, *(needed for _positions, needed, _floor in spans)])
    if sum(fitted(low)) >= width:
        return fitted(low)
    while low < high:
        level = (low + high + 1) // 2
        if sum(fitted(level)) <= width:
            low = level
        else:
            high = level - 1
    return fitted(low)


def _raise_for_spans(widths: list[int], spans: list[tuple[list[int], int]]) -> list[int]:
    """Raise `widths` by the least in all until each spanning cell has the width it needs.

    The cells are taken by their last column, and what one lacks goes to that column, which the
    cells after it are the likeliest to share: for spans of adjacent columns this adds the least.
    """
    raised = list(widths)
    for positions, needed in sorted(spans, key=lambda cell: max(cell[0])):
        lacking = needed - sum(raised[position] for position in positions)
        if lacking > 0:
            raised[max(positions)] += lacking
    return raised


def _scale(widths: list[int], width: int) -> list[int]:
    total = sum(widths)
    if total <= 0:
        return [width // len(widths)] * len(widths)
    return [part * width // total for part in widths]


def _content_extent(container: Any, font_sizes: _FontSizes, cell_margins: Callable[[Any], int]) -> tuple[int, int, int]:
    """The narrowest, the widest and the floor the blocks of a cell can be laid out at, in EMU."""
    minimum = maximum = floor = 0
    for child in container:
        if child.tag == f"{_W}p":
            block_min, block_max, block_floor = _paragraph_extent(child, font_sizes)
        elif child.tag == f"{_W}tbl":
            # A nested table is as narrow and as wide as its columns together.
            nested = column_extents(child, _column_count(child), cell_margins, font_sizes)
            spans_of_columns = [(list(range(offset, offset + span)), floor) for offset, span, _needed, floor in nested.spans]
            block_min, block_max = sum(nested.minimums), sum(nested.maximums)
            block_floor = sum(_raise_for_spans(nested.floors, spans_of_columns))
        elif child.tag in (f"{_W}sdt", f"{_W}sdtContent", f"{_W}customXml"):
            block_min, block_max, block_floor = _content_extent(child.find(f"{_W}sdtContent") if child.tag == f"{_W}sdt" else child, font_sizes, cell_margins)
        else:
            continue
        minimum = max(minimum, block_min)
        maximum = max(maximum, block_max)
        floor = max(floor, block_floor)
    return minimum, maximum, floor


def _paragraph_extent(p: Any, font_sizes: _FontSizes) -> tuple[int, int, int]:
    """The longest word, picture or unbreakable part of an equation of a paragraph, its longest line, and its floor, in EMU.

    The floor is the longest word or part of an equation, or picture lowered to an inch.
    """
    paragraph_size = font_sizes.paragraph(p)
    minimum = line = longest = floor = 0
    # The width of the word being read: formatting can split one word into several runs
    word = 0
    for element in p.iter(f"{_W}t", f"{_W}br", f"{_W}cr", f"{_W}tab", f"{{{WP_NS}}}extent", f"{_M}oMath"):
        if element.tag == f"{_M}oMath":
            # An equation holds no text of the paragraph: counted for nothing, a column of equations
            # got the width of its margins, and Word cut the equations off or broke them a term a line
            unbreakable, whole = _equation_extent(element, paragraph_size * EMU_PER_POINT / 2)
            minimum = max(minimum, unbreakable)
            floor = max(floor, unbreakable)
            word = 0
            line += whole
        elif element.tag == f"{{{WP_NS}}}extent":
            width = _int_attribute(element, "cx")
            # A browser keeps a picture of its size and wraps the text beside it
            minimum = max(minimum, width)
            floor = max(floor, min(width, CAPPED_PICTURE_WIDTH_EMU))
            word = 0
            line += width
        elif element.tag == f"{_W}t":
            character = font_sizes.run(element.getparent(), paragraph_size) * EMU_PER_POINT * CHARACTER_WIDTH_EM / 2
            text = element.text or ""
            for index, piece in enumerate(_WHITESPACE.split(text)):
                # A piece after whitespace starts a word; the first one goes on with the word of the run before
                word = (word if index == 0 else 0) + round(len(piece) * character)
                minimum = max(minimum, word)
                floor = max(floor, word)
            line += round(len(text) * character)
        elif element.tag == f"{_W}tab":
            line += TAB_WIDTH_EMU
            word = 0
        else:
            longest = max(longest, line)
            line = word = 0
    return minimum, max(longest, line), floor


def _equation_extent(equation: Any, em: float) -> tuple[int, int]:
    """The widest part of an equation no line breaks, and the equation on one line, in EMU.

    `em` is the size of the equation in EMU. Word breaks an equation before a sign between two terms
    of its own level, and nowhere inside a fraction, a root, brackets or an index: the narrowest a
    column can hold it is its widest part between two such signs. The widths are estimates from the
    characters, as those of the text are: see the MATH constants.
    """
    widest = part = whole = 0.0
    for child in equation:
        if child.tag == f"{_M}r":
            for character in _run_text(child):
                width = _character_em(character)
                if width == MATH_SIGN_EM:
                    # A line may break before the sign: it starts the next part
                    widest, part = max(widest, part), 0.0
                part += width
                whole += width
        else:
            width = _math_em(child)
            part += width
            whole += width
    return round(max(widest, part) * em), round(whole * em)


def _math_em(element: Any) -> float:
    """The width of a part of an equation set on one line, in ems of the size of the equation."""
    if not isinstance(element.tag, str) or not element.tag.startswith(_M) or element.tag.endswith("Pr"):
        return 0.0
    # What holds a part of an equation, or decorates one, is as wide as what it holds
    return _MATH_ELEMENTS.get(element.tag[len(_M) :], _children_em)(element)


def _children_em(element: Any) -> float:
    return sum(_math_em(child) for child in element)


def _run_em(run: Any) -> float:
    return sum(_character_em(character) for character in _run_text(run))


def _fraction_em(fraction: Any) -> float:
    numerator, denominator = _part_em(fraction, "num"), _part_em(fraction, "den")
    kind = fraction.find(f"{_M}fPr/{_M}type")
    if kind is not None and kind.get(f"{_M}val") in ("lin", "skw"):
        # A fraction on one line sets its parts side by side, with the slash between them
        return numerator + denominator + MATH_BRACKET_EM
    return max(numerator, denominator)


def _scripts_em(element: Any) -> float:
    """The indexes of an element, a limit among them, at the size they are set in."""
    return MATH_SCRIPT_SCALE * max(_part_em(element, script) for script in ("sub", "sup", "lim"))


def _indexed_em(element: Any) -> float:
    # An index stands beside what it belongs to
    return _part_em(element, "e") + _scripts_em(element)


def _limit_em(element: Any) -> float:
    # A limit stands above or below what it belongs to
    return max(_part_em(element, "e"), _scripts_em(element))


def _root_em(root: Any) -> float:
    return MATH_RADICAL_EM + _part_em(root, "e")


def _nary_em(nary: Any) -> float:
    location = nary.find(f"{_M}naryPr/{_M}limLoc")
    above = location is not None and location.get(f"{_M}val") == "undOvr"
    sign = max(MATH_NARY_EM, _scripts_em(nary)) if above else MATH_NARY_EM + _scripts_em(nary)
    return sign + _part_em(nary, "e")


def _brackets_em(brackets: Any) -> float:
    parts = [_math_em(part) for part in brackets.findall(f"{_M}e")]
    # The brackets around, and a separator between each two of the parts
    return sum(parts) + MATH_BRACKET_EM * (len(parts) + 1)


def _matrix_em(matrix: Any) -> float:
    rows = [[_math_em(cell) for cell in row.findall(f"{_M}e")] for row in matrix.findall(f"{_M}mr")]
    columns = max((len(row) for row in rows), default=0)
    widths = [max((row[index] for row in rows if index < len(row)), default=0.0) for index in range(columns)]
    return sum(widths) + MATH_MATRIX_GAP_EM * max(columns - 1, 0)


def _array_em(array: Any) -> float:
    return max((_math_em(row) for row in array.findall(f"{_M}e")), default=0.0)


def _function_em(function: Any) -> float:
    return _part_em(function, "fName") + MATH_SPACE_EM + _part_em(function, "e")


# How the width of each kind of element is counted; any other is as wide as what it holds.
_MATH_ELEMENTS: dict[str, Callable[[Any], float]] = {
    "r": _run_em,
    "f": _fraction_em,
    "sSup": _indexed_em,
    "sSub": _indexed_em,
    "sSubSup": _indexed_em,
    "sPre": _indexed_em,
    "limLow": _limit_em,
    "limUpp": _limit_em,
    "rad": _root_em,
    "nary": _nary_em,
    "d": _brackets_em,
    "m": _matrix_em,
    "eqArr": _array_em,
    "func": _function_em,
}


def _part_em(element: Any, name: str) -> float:
    part = element.find(f"{_M}{name}")
    return _math_em(part) if part is not None else 0.0


def _run_text(run: Any) -> str:
    return "".join(text.text or "" for text in run.findall(f"{_M}t"))


def _character_em(character: str) -> float:
    if character.isalnum():
        return MATH_LETTER_EM
    if character.isspace():
        return MATH_SPACE_EM
    if unicodedata.category(character) in _NO_WIDTH_CATEGORIES:
        return 0.0
    if character in _MATH_BRACKETS:
        return MATH_BRACKET_EM
    return MATH_SIGN_EM


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
    number = whole_number(element.get(f"{_W}val") if element is not None else None)
    return fallback if number is None else number


def _int_attribute(element: Any, name: str) -> int:
    number = whole_number(element.get(name))
    return 0 if number is None else number
