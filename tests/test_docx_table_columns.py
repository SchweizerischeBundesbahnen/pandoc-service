"""Unit tests for :mod:`app.docx_table_columns`."""

import io
import struct
import zlib

import pytest
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches

from app.docx_table_columns import MIN_IMAGE_WIDTH_EMU, decide, distribute

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
INCH = 914400
MARGINS = 2 * 108 * 635


def _margins(_tc) -> int:
    return MARGINS


def _png() -> bytes:
    def chunk(tag: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + tag + payload + struct.pack(">I", zlib.crc32(tag + payload))

    raw = b"".join(b"\x00" + b"\xff\xff\xff" * 4 for _ in range(2))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 2, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def _table(columns: int, rows: int = 1):
    doc = Document()
    return doc, doc.add_table(rows=rows, cols=columns)


# ----------------------- distribute -----------------------


def test_with_room_for_every_column_the_rest_goes_to_those_without_a_width():
    assert distribute([10, 10], [100, 300], 600, [True, False]) == [100, 500]


def test_with_room_for_every_column_and_all_stated_the_rest_is_shared_by_their_widest():
    assert distribute([10, 10], [100, 300], 800, [True, True]) == [200, 600]


def test_short_of_room_each_column_grows_from_its_narrowest_by_what_it_lacks():
    # 200 to share; the first would grow by 100, the second by 300
    assert distribute([100, 100], [200, 400], 400, [False, False]) == [150, 250]


def test_without_room_even_for_the_narrowest_all_are_scaled_down_alike():
    assert distribute([300, 100], [600, 200], 200, [False, False]) == [150, 50]


def test_empty_columns_share_the_width_evenly():
    assert distribute([0, 0], [0, 0], 100, [False, False]) == [50, 50]


# ----------------------- decide -----------------------


def test_widths_stated_for_every_column_are_taken_as_they_are():
    _, table = _table(2)

    assert decide(table._tbl, 10000, [("pct", 1000), ("pct", 4000)], _margins) == [2000, 8000]


def test_absolute_stated_widths_are_scaled_to_the_table():
    _, table = _table(2)

    assert decide(table._tbl, 6000, [("dxa", 1000), ("dxa", 2000)], _margins) == [2000, 4000]


def test_a_column_holding_an_image_takes_the_room_its_neighbour_does_not_need():
    """The even share of pandoc's grid would leave the image half the table."""
    _, table = _table(2)
    table.cell(0, 0).paragraphs[0].add_run("Short")
    table.cell(0, 1).paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(9))

    widths = decide(table._tbl, int(6.5 * INCH), None, _margins)

    assert widths is not None
    assert sum(widths) == pytest.approx(6.5 * INCH, abs=2)
    assert widths[1] > 0.8 * 6.5 * INCH
    assert widths[0] >= MARGINS


def test_a_stated_column_keeps_its_width_and_the_others_share_the_rest():
    _, table = _table(2)
    table.cell(0, 1).paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(9))

    widths = decide(table._tbl, 10 * INCH, [("pct", 1000), None], _margins)

    assert widths is not None
    assert widths[0] == 2 * INCH


def test_stated_widths_which_do_not_match_the_grid_are_ignored():
    _, table = _table(2)
    table.cell(0, 0).paragraphs[0].add_run("Short")
    table.cell(0, 1).paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(9))

    widths = decide(table._tbl, int(6.5 * INCH), [("pct", 5000)], _margins)

    assert widths is not None
    assert widths[1] > widths[0]


def test_an_image_in_a_crowded_table_keeps_at_least_an_inch():
    _, table = _table(3)
    for index in range(2):
        table.cell(0, index).paragraphs[0].add_run(" ".join(["word"] * 400))
    table.cell(0, 2).paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(5))

    widths = decide(table._tbl, int(6.5 * INCH), None, _margins)

    assert widths is not None
    assert widths[2] >= MIN_IMAGE_WIDTH_EMU + MARGINS


def test_a_cell_spanning_columns_widens_them_by_what_they_lack():
    _, table = _table(2, rows=2)
    merged = table.cell(0, 0).merge(table.cell(0, 1))
    merged.paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(4))

    widths = decide(table._tbl, 10 * INCH, None, _margins)

    assert widths is not None
    assert sum(widths) == pytest.approx(10 * INCH, abs=2)
    assert widths[0] == pytest.approx(widths[1], abs=2)


def test_a_nested_table_needs_the_room_of_its_columns():
    _, table = _table(2)
    table.cell(0, 0).paragraphs[0].add_run("Short")
    inner = table.cell(0, 1).add_table(rows=1, cols=2)
    for index in range(2):
        inner.cell(0, index).paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(2))

    widths = decide(table._tbl, int(6.5 * INCH), None, _margins)

    assert widths is not None
    assert widths[1] >= 4 * INCH + 3 * MARGINS


def test_text_in_a_larger_font_needs_a_wider_column():
    doc, table = _table(2)
    table.cell(0, 0).paragraphs[0].add_run(" ".join(["word"] * 10))
    table.cell(0, 1).paragraphs[0].add_run().add_picture(io.BytesIO(_png()), width=Inches(9))
    styles = doc.styles.element
    small = decide(table._tbl, int(6.5 * INCH), None, _margins, styles)
    defaults = styles.find(f"{{{_W}}}docDefaults")
    if defaults is None:
        defaults = parse_xml(f"<w:docDefaults {nsdecls('w')}/>")
        styles.insert(0, defaults)
    for child in list(defaults):
        defaults.remove(child)
    defaults.append(parse_xml(f'<w:rPrDefault {nsdecls("w")}><w:rPr><w:sz w:val="48"/></w:rPr></w:rPrDefault>'))
    for style in styles.findall(f"{{{_W}}}style"):
        size = style.find(f"{{{_W}}}rPr/{{{_W}}}sz")
        if size is not None:
            size.getparent().remove(size)

    large = decide(table._tbl, int(6.5 * INCH), None, _margins, styles)

    assert small is not None
    assert large is not None
    assert large[0] > small[0]


def test_a_table_without_columns_has_no_widths():
    table = parse_xml(f"<w:tbl {nsdecls('w')}><w:tblPr/></w:tbl>")

    assert decide(table, INCH, None, _margins) is None


def test_a_table_without_a_grid_is_laid_out_with_no_more_columns_than_html_allows():
    from app.docx_table_columns import MAX_COLUMNS

    table = parse_xml(f'<w:tbl {nsdecls("w")}><w:tblPr/><w:tr><w:tc><w:tcPr><w:gridSpan w:val="100000000"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>')

    widths = decide(table, MAX_COLUMNS * 10, None, _margins)

    assert widths is not None
    assert len(widths) == MAX_COLUMNS
