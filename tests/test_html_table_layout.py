"""Unit tests for :mod:`app.html_table_layout`.

Each test feeds a small HTML fragment to :func:`html_table_layout.extract` and
checks the recovered :class:`TableLayout` list, plus a couple of tests that run
the extracted layouts through :func:`app.docx_post_process.process` end-to-end to
confirm the width/alignment lands in the resulting ``<w:tblPr>``.
"""

from __future__ import annotations

import io
import re
import shutil
import subprocess
import zipfile

import pytest

from app import docx_post_process, html_table_layout
from app.html_table_layout import MAX_PCT, TableLayout, extract


def _table(style: str) -> str:
    return f'<html><body><table style="{style}"><tr><td>x</td></tr></table></body></html>'


# ----------------------- width extraction -----------------------


@pytest.mark.parametrize(
    ("style", "expected_type", "expected_value"),
    [
        ("width: 100%;", "pct", MAX_PCT),
        ("width: 40%;", "pct", 2000),
        ("width: 80%;", "pct", 4000),
        ("width: 25%;", "pct", 1250),
        ("width: 50px;", "dxa", 750),
        ("width: 100px;", "dxa", 1500),
        ("width: 1in;", "dxa", 1440),
        ("width: auto;", None, None),
        ("width: 0%;", None, None),
        ("width: 0px;", None, None),
        ("border: 1px solid #ccc;", None, None),  # no width declared
    ],
)
def test_width_parsing(style, expected_type, expected_value):
    layout = extract(_table(style))[0]
    assert layout.width_type == expected_type
    assert layout.width_value == expected_value


def test_percentage_over_100_clamped_to_max():
    layout = extract(_table("width: 150%;"))[0]
    assert layout.width_type == "pct"
    assert layout.width_value == MAX_PCT


def test_max_width_is_not_confused_with_width():
    """``max-width`` must not be read as the table width."""
    layout = extract(_table("max-width: 512px;"))[0]
    assert layout.width_type is None
    assert layout.width_value is None


# ----------------------- alignment extraction -----------------------


@pytest.mark.parametrize(
    ("style", "expected_jc"),
    [
        ("margin-left: 0px; margin-right: auto;", "left"),
        ("margin-left: auto; margin-right: auto;", "center"),
        ("margin-left: auto; margin-right: 0px;", "right"),
        ("margin-left: 0px; margin-right: 0px;", None),  # no auto -> no intent
        ("border: 1px solid #ccc;", None),  # no margins at all
    ],
)
def test_alignment_parsing(style, expected_jc):
    assert extract(_table(style))[0].jc == expected_jc


def test_left_margin_becomes_indent_when_left_aligned():
    layout = extract(_table("margin-left: 48px; margin-right: auto;"))[0]
    assert layout.jc == "left"
    assert layout.indent_twips == 720  # 48px * 15 twips


def test_auto_margin_is_never_an_indent():
    """Centered/right-aligned tables use auto margins as the alignment
    mechanism, so no indent should be recorded."""
    centered = extract(_table("margin-left: auto; margin-right: auto;"))[0]
    assert centered.indent_twips is None


# ----------------------- document-order & robustness -----------------------


def test_extract_returns_one_layout_per_table_in_document_order():
    html = "<html><body>" + _inner_tables() + "</body></html>"
    layouts = extract(html)
    assert [layout.width_value for layout in layouts] == [2000, 4000]


def _inner_tables() -> str:
    return '<table style="width: 40%;"><tr><td>a</td></tr></table><table style="width: 80%;"><tr><td>b</td></tr></table>'


def test_nested_table_follows_its_parent_depth_first():
    html = '<html><body><table style="width: 100%;"><tr><td><table style="width: 25%;"><tr><td>n</td></tr></table></td></tr></table></body></html>'
    layouts = extract(html)
    assert [layout.width_value for layout in layouts] == [MAX_PCT, 1250]


def test_accepts_bytes_with_xml_encoding_declaration():
    """The exporter sends an ``<?xml ... encoding=...?>`` prologue; lxml rejects
    that on a decoded str, so extract must feed it bytes."""
    source = b"<?xml version='1.0' encoding='UTF-8'?><html><body>" + _table("width: 40%;").encode() + b"</body></html>"
    layouts = extract(source)
    assert any(layout.width_value == 2000 for layout in layouts)


def test_no_tables_returns_empty_list():
    assert extract("<html><body><p>no tables here</p></body></html>") == []


def test_unparseable_input_returns_empty_list():
    assert extract(b"\xff\xfe not html at all") == []


def test_table_layout_is_empty_helper():
    assert TableLayout().is_empty is True
    assert TableLayout(jc="center").is_empty is False


# ----------------------- end-to-end through docx_post_process -----------------------

# Resolve pandoc to an absolute path (satisfies ruff S607) and skip the
# end-to-end tests when the binary isn't installed; the extraction tests above
# need no external tools.
_PANDOC = shutil.which("pandoc")
requires_pandoc = pytest.mark.skipif(_PANDOC is None, reason="pandoc binary not available")


def _pandoc_html_to_docx(html: str) -> bytes:
    completed = subprocess.run(
        [_PANDOC, "-f", "html", "-t", "docx", "-o", "-"],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    return completed.stdout


def _tbl_props_xml(html_body: str) -> str:
    html = f"<html><head><title>t</title></head><body>{html_body}</body></html>"
    layouts = html_table_layout.extract(html)
    processed = docx_post_process.process(_pandoc_html_to_docx(html), None, None, layouts)
    return zipfile.ZipFile(io.BytesIO(processed)).read("word/document.xml").decode()


@requires_pandoc
def test_end_to_end_percentage_width_applied():
    body = '<table style="width: 40%; margin-left: 0px; margin-right: auto;"><tr><td>x</td></tr></table>'
    document_xml = _tbl_props_xml(body)
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", document_xml, re.DOTALL).group(0)
    assert '<w:tblW w:w="2000" w:type="pct"/>' in props
    assert '<w:jc w:val="left"/>' in props
    assert '<w:tblLayout w:type="autofit"/>' in props


@requires_pandoc
def test_end_to_end_centered_table():
    body = '<table style="width: 25%; margin-left: auto; margin-right: auto;"><tr><td>x</td></tr></table>'
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", _tbl_props_xml(body), re.DOTALL).group(0)
    assert '<w:tblW w:w="1250" w:type="pct"/>' in props
    assert '<w:jc w:val="center"/>' in props


@requires_pandoc
def test_end_to_end_absolute_width_uses_fixed_layout_and_rescaled_grid():
    body = '<table style="width: 100px; margin-left: 0px; margin-right: auto;"><tr><td>a</td><td>b</td></tr></table>'
    document_xml = _tbl_props_xml(body)
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", document_xml, re.DOTALL).group(0)
    assert '<w:tblW w:w="1500" w:type="dxa"/>' in props
    assert '<w:tblLayout w:type="fixed"/>' in props
    grid = re.search(r"<w:tblGrid>.*?</w:tblGrid>", document_xml, re.DOTALL).group(0)
    col_widths = [int(w) for w in re.findall(r'w:w="(\d+)"', grid)]
    assert sum(col_widths) == 1500  # 100px * 15 twips, distributed across columns


@requires_pandoc
def test_end_to_end_default_when_no_layouts_keeps_full_width():
    """A table with no style still fills the column (backwards-compatible)."""
    body = "<table><tr><td>x</td></tr></table>"
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", _tbl_props_xml(body), re.DOTALL).group(0)
    assert '<w:tblW w:w="5000" w:type="pct"/>' in props
    assert '<w:tblLayout w:type="autofit"/>' in props


@requires_pandoc
def test_count_mismatch_falls_back_to_defaults():
    """When the number of layouts doesn't match the number of tables, no
    width/alignment is applied and every table keeps the 100% default."""
    docx = _pandoc_html_to_docx("<html><head><title>t</title></head><body><table><tr><td>x</td></tr></table></body></html>")
    # Deliberately pass too many layouts (2 for 1 table).
    bogus = [TableLayout(width_type="pct", width_value=2000), TableLayout(width_type="pct", width_value=1000)]
    processed = docx_post_process.process(docx, None, None, bogus)
    document_xml = zipfile.ZipFile(io.BytesIO(processed)).read("word/document.xml").decode()
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", document_xml, re.DOTALL).group(0)
    assert '<w:tblW w:w="5000" w:type="pct"/>' in props
    assert "w:jc" not in props


def test_tables_after_a_huge_image_are_extracted():
    """A data: URI over libxml2's 10 MB limit used to end the parsed tree, so later tables were missed."""
    from tests.test_html_document import HUGE_SRC

    table = '<table style="width: {}"><tr><td>x</td></tr></table>'
    src = f'<html><body>{table.format("40%")}<img src="{HUGE_SRC}">{table.format("60%")}</body></html>'

    assert len(extract(src)) == 2


# ----------------------- column widths -----------------------


def test_column_widths_come_from_the_cells_of_the_first_row():
    """Polarion's work item attribute tables state 20 % and 80 % on their cells, which pandoc drops."""
    html = '<table style="width: 100%"><tr><td style="width:20%; border: 1px solid">a</td><td style="width:80%">b</td></tr><tr><td style="width:50%">c</td><td>d</td></tr></table>'
    assert extract(html)[0].column_widths == (("pct", 1000), ("pct", 4000))


def test_column_widths_come_from_the_first_row_of_a_head():
    html = "<table><thead><tr><th style='width: 30%'>a</th><th>b</th><th width='120'>c</th></tr></thead><tbody><tr><td>1</td><td>2</td><td>3</td></tr></tbody></table>"
    assert extract(html)[0].column_widths == (("pct", 1500), None, ("dxa", 1800))


def test_column_widths_of_a_colgroup_win_over_the_cells():
    html = "<table><colgroup><col style='width: 25%'/><col span='2' width='10%'/></colgroup><tr><td style='width: 90%'>a</td><td>b</td><td>c</td></tr></table>"
    assert extract(html)[0].column_widths == (("pct", 1250), ("pct", 500), ("pct", 500))


def test_a_cell_spanning_columns_states_no_width_for_any_of_them():
    html = "<table><tr><td colspan='2' style='width: 60%'>a</td><td style='width: 40%'>b</td></tr></table>"
    assert extract(html)[0].column_widths == (None, None, ("pct", 2000))


def test_a_table_without_column_widths_states_none():
    html = "<table><tr><td>a</td><td style='width: auto'>b</td></tr></table>"
    layout = extract(html)[0]
    assert layout.column_widths is None
    assert layout.is_empty


def test_column_widths_of_a_nested_table_stay_its_own():
    html = "<table><tr><td><table><tr><td style='width: 70%'>x</td><td style='width: 30%'>y</td></tr></table></td><td>b</td></tr></table>"
    outer, inner = extract(html)
    assert outer.column_widths is None
    assert inner.column_widths == (("pct", 3500), ("pct", 1500))


@pytest.mark.parametrize(
    "html",
    [
        "<table><tr><td colspan='100000000'>a</td><td style='width: 40%'>b</td></tr></table>",
        "<table><colgroup><col span='100000000'/><col style='width: 40%'/></colgroup><tr><td>a</td></tr></table>",
    ],
)
def test_a_huge_span_states_no_column_widths(html):
    """A one-line tag must not make the service build a list of a hundred million columns."""
    assert extract(html)[0].column_widths is None


@pytest.mark.parametrize(
    "html",
    [
        "<table><tr><td colspan='999'>a</td><td style='width: 40%'>b</td></tr></table>",
        "<table><colgroup><col span='999'/><col style='width: 40%'/></colgroup><tr><td>a</td></tr></table>",
    ],
)
def test_a_table_as_wide_as_a_browser_allows_keeps_its_column_widths(html):
    assert extract(html)[0].column_widths == (*(None,) * 999, ("pct", 2000))


def test_a_table_wider_than_a_browser_allows_states_no_column_widths():
    cols = "<col span='1000' style='width: 1%'/>" * 3
    assert extract(f"<table><colgroup>{cols}</colgroup><tr><td>a</td></tr></table>")[0].column_widths is None


@pytest.mark.parametrize(
    ("colspan", "covered"),
    # As entities, which parse to the character whatever encoding the parser takes the bytes in
    [("&#178;", 1), ("&#1633;&#1634;", 1), ("0003", 3), ("0" * 5000 + "2", 2), ("9" * 5000, html_table_layout.MAX_COLUMNS)],
)
def test_a_colspan_is_read_from_ascii_digits_without_failing_the_conversion(colspan, covered):
    """str.isdigit accepts a superscript two, which int refuses; int refuses more than 4300 digits."""
    html = f"<table><tr><td colspan='{colspan}'>a</td><td style='width: 40%'>b</td></tr></table>"

    widths = extract(html)[0].column_widths

    if covered < html_table_layout.MAX_COLUMNS:
        assert widths == (*(None,) * covered, ("pct", 2000))
    else:
        assert widths is None


def test_a_col_span_is_read_from_ascii_digits_without_failing_the_conversion():
    html = "<table><colgroup><col span='&#178;' style='width: 30%'/><col style='width: 70%'/></colgroup><tr><td>a</td></tr></table>"

    assert extract(html)[0].column_widths == (("pct", 1500), ("pct", 3500))


def test_a_colgroup_without_widths_leaves_the_columns_to_the_cells():
    html = "<table><colgroup><col/><col/></colgroup><tr><td style='width: 20%'>a</td><td style='width: 80%'>b</td></tr></table>"
    assert extract(html)[0].column_widths == (("pct", 1000), ("pct", 4000))


def test_a_col_without_a_width_leaves_its_column_to_the_cell():
    html = "<table><colgroup><col style='width: 30%'/><col/></colgroup><tr><td style='width: 90%'>a</td><td style='width: 70%'>b</td></tr></table>"
    assert extract(html)[0].column_widths == (("pct", 1500), ("pct", 3500))


def test_cells_beyond_the_colgroup_keep_their_widths():
    html = "<table><colgroup><col style='width: 30%'/></colgroup><tr><td>a</td><td style='width: 70%'>b</td></tr></table>"
    assert extract(html)[0].column_widths == (("pct", 1500), ("pct", 3500))
