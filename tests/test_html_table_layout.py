"""Unit tests for :mod:`app.HtmlTableLayout`.

Each test feeds a small HTML fragment to :func:`HtmlTableLayout.extract` and
checks the recovered :class:`TableLayout` list, plus a couple of tests that run
the extracted layouts through :func:`app.DocxPostProcess.process` end-to-end to
confirm the width/alignment lands in the resulting ``<w:tblPr>``.
"""

from __future__ import annotations

import io
import re
import shutil
import subprocess
import zipfile

import pytest

from app import DocxPostProcess, HtmlTableLayout
from app.HtmlTableLayout import MAX_PCT, TableLayout, extract


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


# ----------------------- end-to-end through DocxPostProcess -----------------------

# Resolve pandoc to an absolute path (satisfies ruff S607) and skip the
# end-to-end tests when the binary isn't installed; the extraction tests above
# need no external tools.
_PANDOC = shutil.which("pandoc")
requires_pandoc = pytest.mark.skipif(_PANDOC is None, reason="pandoc binary not available")


def _pandoc_html_to_docx(html: str) -> bytes:
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "html", "-t", "docx", "-o", "-"],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    return completed.stdout


def _tbl_props_xml(html_body: str) -> str:
    html = f"<html><head><title>t</title></head><body>{html_body}</body></html>"
    layouts = HtmlTableLayout.extract(html)
    processed = DocxPostProcess.process(_pandoc_html_to_docx(html), None, None, layouts)
    return zipfile.ZipFile(io.BytesIO(processed)).read("word/document.xml").decode()


@requires_pandoc
def test_end_to_end_percentage_width_applied():
    body = '<table style="width: 40%; margin-left: 0px; margin-right: auto;"><tr><td>x</td></tr></table>'
    document_xml = _tbl_props_xml(body)
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", document_xml, re.S).group(0)
    assert '<w:tblW w:w="2000" w:type="pct"/>' in props
    assert '<w:jc w:val="left"/>' in props
    assert '<w:tblLayout w:type="autofit"/>' in props


@requires_pandoc
def test_end_to_end_centered_table():
    body = '<table style="width: 25%; margin-left: auto; margin-right: auto;"><tr><td>x</td></tr></table>'
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", _tbl_props_xml(body), re.S).group(0)
    assert '<w:tblW w:w="1250" w:type="pct"/>' in props
    assert '<w:jc w:val="center"/>' in props


@requires_pandoc
def test_end_to_end_absolute_width_uses_fixed_layout_and_rescaled_grid():
    body = '<table style="width: 100px; margin-left: 0px; margin-right: auto;"><tr><td>a</td><td>b</td></tr></table>'
    document_xml = _tbl_props_xml(body)
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", document_xml, re.S).group(0)
    assert '<w:tblW w:w="1500" w:type="dxa"/>' in props
    assert '<w:tblLayout w:type="fixed"/>' in props
    grid = re.search(r"<w:tblGrid>.*?</w:tblGrid>", document_xml, re.S).group(0)
    col_widths = [int(w) for w in re.findall(r'w:w="(\d+)"', grid)]
    assert sum(col_widths) == 1500  # 100px * 15 twips, distributed across columns


@requires_pandoc
def test_end_to_end_default_when_no_layouts_keeps_full_width():
    """A table with no style still fills the column (backwards-compatible)."""
    body = "<table><tr><td>x</td></tr></table>"
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", _tbl_props_xml(body), re.S).group(0)
    assert '<w:tblW w:w="5000" w:type="pct"/>' in props
    assert '<w:tblLayout w:type="autofit"/>' in props


@requires_pandoc
def test_count_mismatch_falls_back_to_defaults():
    """When the number of layouts doesn't match the number of tables, no
    width/alignment is applied and every table keeps the 100% default."""
    docx = _pandoc_html_to_docx("<html><head><title>t</title></head><body><table><tr><td>x</td></tr></table></body></html>")
    # Deliberately pass too many layouts (2 for 1 table).
    bogus = [TableLayout(width_type="pct", width_value=2000), TableLayout(width_type="pct", width_value=1000)]
    processed = DocxPostProcess.process(docx, None, None, bogus)
    document_xml = zipfile.ZipFile(io.BytesIO(processed)).read("word/document.xml").decode()
    props = re.search(r"<w:tblPr>.*?</w:tblPr>", document_xml, re.S).group(0)
    assert '<w:tblW w:w="5000" w:type="pct"/>' in props
    assert "w:jc" not in props
