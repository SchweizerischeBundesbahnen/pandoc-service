"""End-to-end integration tests for the table-cell-styling section of
filters/inline_styles.lua.

These tests convert HTML → DOCX through the pandoc-service container (which
applies the Lua filter automatically for HTML→DOCX conversions), then inspect
the resulting DOCX XML for cell-level properties that the default DOCX writer
would otherwise drop (background-color, borders, vertical-align).
"""

import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from tests.test_container import TestParameters

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _convert_html_to_docx(test_parameters: TestParameters, html: str, *, preserve_table_styles: bool = True) -> bytes:
    """Convert HTML to DOCX via the pandoc-service container API.

    Returns the raw DOCX bytes. Raises AssertionError on non-2xx.
    """
    url = f"{test_parameters.base_url}/convert/html/to/docx"
    params = {}
    if preserve_table_styles:
        params["preserve_table_styles"] = "true"
    response = test_parameters.request_session.post(url, data=html, params=params)
    if response.status_code // 100 != 2:
        raise AssertionError(f"pandoc-service returned {response.status_code}:\n{response.text}")
    return response.content


def _parse_document_xml(docx_bytes: bytes) -> ET.Element:
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def _find_all(root: ET.Element, xpath: str) -> list[ET.Element]:
    return root.findall(xpath, {"w": W_NS})


# ---- Tests ----


def test_cell_background_color_preserved(test_parameters: TestParameters):
    """<td style="background-color:#D9EAF7"> should produce <w:shd w:fill="D9EAF7">."""
    html = """<table>
      <tr><td style="background-color:#D9EAF7;">Colored cell</td><td>Plain</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    shd_els = _find_all(root, ".//w:tc/w:tcPr/w:shd")
    fills = [el.get(f"{{{W_NS}}}fill") for el in shd_els]
    assert "D9EAF7" in fills, f"expected D9EAF7 in shd fills, got {fills}"


def test_cell_border_solid(test_parameters: TestParameters):
    """A solid border should produce <w:bottom w:val="single">."""
    html = """<table>
      <tr><td style="border-bottom:1.5pt solid black;">Bordered</td><td>Plain</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    borders = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders")
    assert len(borders) >= 1, "expected at least one <w:tcBorders>"
    bottom = borders[0].find(f"{{{W_NS}}}bottom")
    assert bottom is not None, "expected <w:bottom> inside tcBorders"
    assert bottom.get(f"{{{W_NS}}}val") == "single"
    assert bottom.get(f"{{{W_NS}}}color") == "000000"
    assert int(bottom.get(f"{{{W_NS}}}sz", "0")) == 12  # 1.5pt = 12 eighths


def test_cell_border_dashed(test_parameters: TestParameters):
    """A dashed border should produce w:val="dashed"."""
    html = """<table>
      <tr><td style="border-bottom:1pt dashed #6AA84F;">Dashed</td><td>X</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    bottoms = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:bottom")
    assert len(bottoms) >= 1
    assert bottoms[0].get(f"{{{W_NS}}}val") == "dashed"
    assert bottoms[0].get(f"{{{W_NS}}}color") == "6AA84F"


def test_cell_border_rgb_color_with_spaces(test_parameters: TestParameters):
    """border: 1px solid rgb(255, 0, 0) should parse the color correctly."""
    html = """<table>
      <tr><td style="border-bottom:1px solid rgb(255, 0, 0);">RGB</td><td>X</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    bottoms = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:bottom")
    assert len(bottoms) >= 1
    assert bottoms[0].get(f"{{{W_NS}}}color") == "FF0000", f"expected FF0000 for rgb(255,0,0), got {bottoms[0].get(f'{{{W_NS}}}color')}"


def test_cell_border_dotted(test_parameters: TestParameters):
    """A dotted border should produce w:val="dotted"."""
    html = """<table>
      <tr><td style="border-right:1.5pt dotted #3C78D8;">Dotted</td><td>X</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    rights = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:right")
    assert len(rights) >= 1
    assert rights[0].get(f"{{{W_NS}}}val") == "dotted"


def test_cell_border_double(test_parameters: TestParameters):
    """A double border should produce w:val="double"."""
    html = """<table>
      <tr><td style="border-bottom:1.5pt double #CC0000;">Double</td><td>X</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    bottoms = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:bottom")
    assert len(bottoms) >= 1
    assert bottoms[0].get(f"{{{W_NS}}}val") == "double"


def test_colspan_preserved(test_parameters: TestParameters):
    """colspan=2 should produce <w:gridSpan w:val="2">."""
    html = """<table>
      <tr><td colspan="2" style="background-color:#F4CCCC;">Merged</td></tr>
      <tr><td>A</td><td>B</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    spans = _find_all(root, ".//w:tc/w:tcPr/w:gridSpan")
    vals = [el.get(f"{{{W_NS}}}val") for el in spans]
    assert "2" in vals, f"expected gridSpan val=2, got {vals}"


def test_rowspan_preserved(test_parameters: TestParameters):
    """rowspan=2 should produce vMerge restart + vMerge continue."""
    html = """<table>
      <tr><td rowspan="2" style="background-color:#FFF2CC;">Spanning</td><td>B1</td></tr>
      <tr><td>B2</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    vmerges = _find_all(root, ".//w:tc/w:tcPr/w:vMerge")
    vals = [el.get(f"{{{W_NS}}}val") for el in vmerges]
    assert "restart" in vals, f"expected vMerge restart, got {vals}"
    # Continuation cell has vMerge with no val attribute (or val="continue")
    assert any(v is None or v == "continue" for v in vals), f"expected vMerge continuation, got {vals}"


def test_inline_styles_inside_styled_cell(test_parameters: TestParameters):
    """Inline styles (<span style="...">) inside a styled cell should be preserved."""
    html = """<table>
      <tr>
        <td style="background-color:#D9EAD3;">
          <span style="font-weight:bold;color:#274E13;">Bold green text</span>
        </td>
        <td>Plain</td>
      </tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    # Cell background
    shd_fills = [el.get(f"{{{W_NS}}}fill") for el in _find_all(root, ".//w:tc/w:tcPr/w:shd")]
    assert "D9EAD3" in shd_fills, f"cell background missing, fills: {shd_fills}"

    # Run-level bold + color
    runs = _find_all(root, ".//w:tc//w:r")
    found_bold_green = False
    for r in runs:
        rpr = r.find(f"{{{W_NS}}}rPr")
        if rpr is None:
            continue
        has_bold = rpr.find(f"{{{W_NS}}}b") is not None
        color_el = rpr.find(f"{{{W_NS}}}color")
        has_green = color_el is not None and color_el.get(f"{{{W_NS}}}val") == "274E13"
        if has_bold and has_green:
            found_bold_green = True
            break
    assert found_bold_green, "expected bold + green run inside styled cell"


def test_vertical_align_preserved(test_parameters: TestParameters):
    """vertical-align:top on a cell should produce <w:vAlign w:val="top"/>."""
    html = """<table>
      <tr><td style="vertical-align:top;">Top-aligned</td><td>X</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    valigns = _find_all(root, ".//w:tc/w:tcPr/w:vAlign")
    vals = [el.get(f"{{{W_NS}}}val") for el in valigns]
    assert "top" in vals, f"expected vAlign top, got {vals}"


def test_unstyled_table_passes_through(test_parameters: TestParameters):
    """A table with no styled cells should NOT be rewritten to raw OOXML —
    the default DOCX writer handles it instead."""
    html = """<table><tr><td>A</td><td>B</td></tr></table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    # Should still have a table
    tables = _find_all(root, ".//w:tbl")
    assert len(tables) >= 1, "table should still exist"
    # Content should be present
    texts = [el.text for el in _find_all(root, ".//w:t") if el.text]
    assert "A" in texts and "B" in texts


def test_disabled_by_default_without_metadata_flag(test_parameters: TestParameters):
    """Without -M preserve_table_styles=true, styled tables should NOT be
    rewritten — the feature is opt-in."""
    html = """<table>
      <tr><td style="background-color:#D9EAF7;">Styled</td><td>Plain</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html, preserve_table_styles=False))

    # Cell background should NOT be present (default pandoc drops it)
    shd_els = _find_all(root, ".//w:tc/w:tcPr/w:shd")
    assert len(shd_els) == 0, f"expected no cell shading without opt-in, got {len(shd_els)}"


def test_text_align_center_preserved(test_parameters: TestParameters):
    """text-align:center on a cell should produce <w:jc w:val="center"/> on paragraphs."""
    html = """<table>
      <tr><td style="text-align:center;background-color:#F2F2F2;">Centered</td><td>X</td></tr>
    </table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    jcs = _find_all(root, ".//w:tc//w:p/w:pPr/w:jc")
    vals = [el.get(f"{{{W_NS}}}val") for el in jcs]
    assert "center" in vals, f"expected jc center, got {vals}"


def test_full_test_html_file(test_parameters: TestParameters):
    """Smoke test: convert the real test.html file end-to-end."""
    test_html = Path(__file__).resolve().parents[1] / "tests" / "data" / "test" / "test.html"
    if not test_html.exists():
        pytest.skip("tests/data/test/test.html not found")
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, test_html.read_text(encoding="utf-8")))

    # Should have tables with cell styling
    shd_els = _find_all(root, ".//w:tc/w:tcPr/w:shd")
    assert len(shd_els) > 0, "expected styled cells in the output"

    # Check that D9EAF7 (the header row background) is present
    fills = [el.get(f"{{{W_NS}}}fill") for el in shd_els]
    assert "D9EAF7" in fills, f"expected D9EAF7 header background, got {fills}"

    # Check borders exist
    borders = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders")
    assert len(borders) > 0, "expected borders in styled tables"


# ---- Adjacent table separation ----

# Word ignores these when it decides whether two tables touch; see
# _TABLE_RANGE_MARKERS in app/docx_post_process.py.
_RANGE_MARKERS = {"bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "proofErr"}

_STYLED_TABLE = '<table><tbody><tr><td style="background-color:#EEEEEE;">A</td><td>B</td></tr></tbody></table>'
_PLAIN_TABLE = "<table><tbody><tr><td>X</td><td>Y</td></tr></tbody></table>"

# Two tables at different AST depths: the first wrapped in a <div>, the second
# its sibling. This is the shape Polarion emits around work item fields, and
# pandoc's own table separation does not see the two as adjacent.
_NESTED_DIV_TABLES_HTML = f"""<div id="workitem">
  <div id="steps">{_STYLED_TABLE}</div>
  <table><tbody><tr><td style="width:20%;">Severity</td><td style="width:80%;">Basic</td></tr></tbody></table>
</div>"""


def _body_blocks(root: ET.Element) -> list[str]:
    """Local names of the body's block children, minus markers Word skips over."""
    body = root.find(f"{{{W_NS}}}body")
    names = [child.tag.split("}")[-1] for child in body]
    return [name for name in names if name not in _RANGE_MARKERS and name != "sectPr"]


@pytest.mark.parametrize(
    "html",
    [
        _NESTED_DIV_TABLES_HTML,
        _STYLED_TABLE + _STYLED_TABLE,
        _STYLED_TABLE + _PLAIN_TABLE,
        _PLAIN_TABLE + _STYLED_TABLE,
    ],
    ids=["nested-div", "styled-styled", "styled-plain", "plain-styled"],
)
@pytest.mark.parametrize("preserve_table_styles", [True, False], ids=["opt-in", "default"])
def test_adjacent_tables_are_not_merged(test_parameters: TestParameters, html: str, preserve_table_styles: bool):
    """Consecutive tables must stay separate tables, separated by one paragraph."""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html, preserve_table_styles=preserve_table_styles))

    assert _body_blocks(root) == ["tbl", "p", "tbl"]


@pytest.mark.parametrize("preserve_table_styles", [True, False], ids=["opt-in", "default"])
def test_lone_styled_table_gains_no_extra_paragraph(test_parameters: TestParameters, preserve_table_styles: bool):
    """A table whose neighbour is ordinary content is left as it was."""
    html = f"<p>before</p>{_STYLED_TABLE}<p>after</p>"
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html, preserve_table_styles=preserve_table_styles))

    assert _body_blocks(root) == ["p", "tbl", "p"]


@pytest.mark.parametrize("preserve_table_styles", [True, False], ids=["opt-in", "default"])
def test_trailing_styled_table_gains_no_extra_paragraph(test_parameters: TestParameters, preserve_table_styles: bool):
    """A table that ends the document is left as it was."""
    html = f"<p>before</p>{_STYLED_TABLE}"
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html, preserve_table_styles=preserve_table_styles))

    assert _body_blocks(root) == ["p", "tbl"]


# ---- Cell content wrapped in a styled <div> ----------------------------
#
# Polarion wraps a work item cell's content in <div style="text-align:...">.
# filter.Table replaces the Table before pandoc descends into its cells, so
# filter.Div never runs on that div and block_to_ooxml has to unwrap it. It
# used to fall through to the stringify fallback, which cost the cell its
# alignment and every run property inside it.


def _cell_with_text(root: ET.Element, needle: str) -> ET.Element:
    for tc in root.iter(f"{{{W_NS}}}tc"):
        text = "".join(t.text or "" for t in tc.iter(f"{{{W_NS}}}t"))
        if needle in text:
            return tc
    raise AssertionError(f"no <w:tc> contained {needle!r}")


def _cell_jc(tc: ET.Element) -> str | None:
    jc = tc.find(f"{{{W_NS}}}p/{{{W_NS}}}pPr/{{{W_NS}}}jc")
    return None if jc is None else jc.get(f"{{{W_NS}}}val")


def _cell_border(tc: ET.Element, side: str) -> tuple[str | None, str | None] | None:
    el = tc.find(f"{{{W_NS}}}tcPr/{{{W_NS}}}tcBorders/{{{W_NS}}}{side}")
    if el is None:
        return None
    return el.get(f"{{{W_NS}}}val"), el.get(f"{{{W_NS}}}color")


def _styled_cell(css: str, inner: str) -> str:
    return f'<table><tbody><tr><td style="background-color:#D9EAD3;{css}"><div style="text-align: center;">{inner}</div></td></tr></tbody></table>'


def test_div_wrapped_cell_keeps_its_alignment(test_parameters: TestParameters):
    html = _styled_cell("", '<span style="font-weight: bold;">Group A</span>')
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    assert _cell_jc(_cell_with_text(root, "Group A")) == "center", "the div's text-align did not reach the cell paragraph"


@pytest.mark.parametrize(
    ("css", "tag"),
    [
        ("font-weight: bold;", "b"),
        ("font-style: italic;", "i"),
        ("text-decoration: underline;", "u"),
    ],
    ids=["bold", "italic", "underline"],
)
def test_div_wrapped_cell_keeps_run_formatting(test_parameters: TestParameters, css: str, tag: str):
    html = _styled_cell("", f'<span style="{css}">Styled</span>')
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    rpr = _cell_with_text(root, "Styled").find(f".//{{{W_NS}}}rPr")
    assert rpr is not None and rpr.find(f"{{{W_NS}}}{tag}") is not None, f"<w:{tag}> missing from a div-wrapped cell"


def test_div_wrapped_cell_keeps_run_color(test_parameters: TestParameters):
    html = _styled_cell("", '<span style="color: #274E13;">Green</span>')
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    color = _cell_with_text(root, "Green").find(f".//{{{W_NS}}}rPr/{{{W_NS}}}color")
    assert color is not None and color.get(f"{{{W_NS}}}val") == "274E13"


# ---- Border propagation across inner edges -----------------------------
#
# CSS border-collapse names the line between two cells once; OOXML wants it on
# both cells, and the side nobody names falls back to the table-level
# insideH/insideV (solid black). Renderers resolve that disagreement in favour
# of the table default, so a dashed cell border came out solid black on every
# inner edge while the same border on an outer edge rendered correctly.

_DASHED = "1.5pt dashed #6AA84F"
_BORDER_GRID_HTML = f"""<table><tbody>
<tr><td style="border-bottom:{_DASHED};border-right:{_DASHED};">A</td><td>B</td></tr>
<tr><td>C</td><td>D</td></tr>
</tbody></table>"""


def test_cell_border_reaches_the_neighbour_below(test_parameters: TestParameters):
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, _BORDER_GRID_HTML))

    assert _cell_border(_cell_with_text(root, "A"), "bottom") == ("dashed", "6AA84F")
    assert _cell_border(_cell_with_text(root, "C"), "top") == ("dashed", "6AA84F"), "the cell below did not receive the shared edge"


def test_cell_border_reaches_the_neighbour_to_the_right(test_parameters: TestParameters):
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, _BORDER_GRID_HTML))

    assert _cell_border(_cell_with_text(root, "A"), "right") == ("dashed", "6AA84F")
    assert _cell_border(_cell_with_text(root, "B"), "left") == ("dashed", "6AA84F"), "the cell to the right did not receive the shared edge"


def test_a_border_the_neighbour_names_itself_is_kept(test_parameters: TestParameters):
    """Both cells naming the edge is a genuine CSS conflict; leave both alone."""
    html = f"""<table><tbody>
    <tr><td style="border-bottom:{_DASHED};">A</td></tr>
    <tr><td style="border-top:1pt dotted #CC0000;">C</td></tr>
    </tbody></table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    assert _cell_border(_cell_with_text(root, "A"), "bottom") == ("dashed", "6AA84F")
    assert _cell_border(_cell_with_text(root, "C"), "top") == ("dotted", "CC0000")


def test_rowspan_origin_does_not_take_a_continuation_border(test_parameters: TestParameters):
    """A vMerge continuation shares its origin's css table.

    Writing a propagated border into it in place would give the origin a
    border belonging to a row it does not touch, so share_border copies first.
    """
    html = f"""<table><tbody>
    <tr><td rowspan="2">Merged</td><td>X</td></tr>
    <tr><td style="border-left:{_DASHED};">Y</td></tr>
    </tbody></table>"""
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    merged = _cell_with_text(root, "Merged")
    assert _cell_border(merged, "right") is None, "the origin row took a border that belongs to the continuation row"


# ---- Formatting declared on the cell itself ----------------------------
#
# <th style="font-weight:bold"> styles the cell, not a span inside it, so
# nothing in the cell content carries the formatting. Polarion emits its table
# headers exactly that way and the text came out plain.


@pytest.mark.parametrize(
    ("css", "tag"),
    [
        ("font-weight: bold;", "b"),
        ("font-style: italic;", "i"),
        ("text-decoration: underline;", "u"),
    ],
    ids=["bold", "italic", "underline"],
)
def test_cell_level_formatting_reaches_the_runs(test_parameters: TestParameters, css: str, tag: str):
    html = f'<table><tbody><tr><td style="background-color:#F2F2F2;{css}">Header</td></tr></tbody></table>'
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    rpr = _cell_with_text(root, "Header").find(f".//{{{W_NS}}}rPr")
    assert rpr is not None and rpr.find(f"{{{W_NS}}}{tag}") is not None, f"<w:{tag}> from the cell's own CSS never reached the runs"


def test_cell_level_formatting_reaches_a_div_wrapped_cell(test_parameters: TestParameters):
    """The Polarion shape: bold on the <th>, content wrapped in a <div>."""
    html = '<table><tbody><tr><th style="font-weight:bold;background-color:#F2F2F2;"><div style="text-align: center;">Nom</div></th></tr></tbody></table>'
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    cell = _cell_with_text(root, "Nom")
    assert cell.find(f".//{{{W_NS}}}rPr/{{{W_NS}}}b") is not None, "the cell's bold was lost"
    assert _cell_jc(cell) == "center", "the div's alignment was lost"


def test_a_span_overrides_the_cell_formatting(test_parameters: TestParameters):
    """Cell CSS seeds the walk; nested spans cascade over it as usual."""
    html = '<table><tbody><tr><td style="font-weight:bold;color:#111111;"><span style="color:#CC0000;">Red</span></td></tr></tbody></table>'
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    rpr = _cell_with_text(root, "Red").find(f".//{{{W_NS}}}rPr")
    assert rpr.find(f"{{{W_NS}}}b") is not None, "the cell's bold was dropped by the span"
    assert rpr.find(f"{{{W_NS}}}color").get(f"{{{W_NS}}}val") == "CC0000", "the span's colour did not override the cell's"


def test_cell_background_is_not_repeated_as_run_shading(test_parameters: TestParameters):
    """background-color is the cell fill, already emitted as <w:shd> in <w:tcPr>.

    Seeding it into the run properties too would paint it behind the text a
    second time.
    """
    html = '<table><tbody><tr><td style="background-color:#F2F2F2;font-weight:bold;">Shaded</td></tr></tbody></table>'
    root = _parse_document_xml(_convert_html_to_docx(test_parameters, html))

    cell = _cell_with_text(root, "Shaded")
    assert cell.find(f"{{{W_NS}}}tcPr/{{{W_NS}}}shd") is not None, "the cell lost its fill"
    assert cell.find(f".//{{{W_NS}}}rPr/{{{W_NS}}}shd") is None, "the cell fill was repeated as run shading"
