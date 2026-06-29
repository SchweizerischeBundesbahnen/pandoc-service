"""End-to-end integration tests for the table-cell-styling section of
filters/inline_styles.lua.

These tests invoke a real ``pandoc`` binary with the Lua filter, then
inspect the resulting DOCX XML for cell-level properties that the
default DOCX writer would otherwise drop (background-color, borders,
vertical-align).
"""

import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

PANDOC = shutil.which("pandoc")
FILTER_PATH = Path(__file__).resolve().parents[1] / "filters" / "inline_styles.lua"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

pytestmark = pytest.mark.skipif(
    PANDOC is None or not FILTER_PATH.exists(),
    reason="pandoc binary or filters/inline_styles.lua not available",
)


def _convert_html_to_docx(html: str, output_path: Path, *, preserve_table_styles: bool = True) -> None:
    src_path = output_path.with_suffix(".html")
    src_path.write_text(html, encoding="utf-8")
    cmd = [PANDOC, "-f", "html", "-t", "docx", f"--lua-filter={FILTER_PATH}", "-o", str(output_path), str(src_path)]
    if preserve_table_styles:
        cmd.extend(["-M", "preserve_table_styles=true"])
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise AssertionError(f"pandoc failed (exit {result.returncode}):\n{result.stderr}")


def _parse_document_xml(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def _find_all(root: ET.Element, xpath: str) -> list[ET.Element]:
    return root.findall(xpath, {"w": W_NS})


# ---- Tests ----


def test_cell_background_color_preserved():
    """<td style="background-color:#D9EAF7"> should produce <w:shd w:fill="D9EAF7">."""
    html = """<table>
      <tr><td style="background-color:#D9EAF7;">Colored cell</td><td>Plain</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        shd_els = _find_all(root, ".//w:tc/w:tcPr/w:shd")
        fills = [el.get(f"{{{W_NS}}}fill") for el in shd_els]
        assert "D9EAF7" in fills, f"expected D9EAF7 in shd fills, got {fills}"


def test_cell_border_solid():
    """A solid border should produce <w:bottom w:val="single">."""
    html = """<table>
      <tr><td style="border-bottom:1.5pt solid black;">Bordered</td><td>Plain</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        borders = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders")
        assert len(borders) >= 1, "expected at least one <w:tcBorders>"
        bottom = borders[0].find(f"{{{W_NS}}}bottom")
        assert bottom is not None, "expected <w:bottom> inside tcBorders"
        assert bottom.get(f"{{{W_NS}}}val") == "single"
        assert bottom.get(f"{{{W_NS}}}color") == "000000"
        assert int(bottom.get(f"{{{W_NS}}}sz", "0")) == 12  # 1.5pt = 12 eighths


def test_cell_border_dashed():
    """A dashed border should produce w:val="dashed"."""
    html = """<table>
      <tr><td style="border-bottom:1pt dashed #6AA84F;">Dashed</td><td>X</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        bottoms = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:bottom")
        assert len(bottoms) >= 1
        assert bottoms[0].get(f"{{{W_NS}}}val") == "dashed"
        assert bottoms[0].get(f"{{{W_NS}}}color") == "6AA84F"


def test_cell_border_rgb_color_with_spaces():
    """border: 1px solid rgb(255, 0, 0) should parse the color correctly."""
    html = """<table>
      <tr><td style="border-bottom:1px solid rgb(255, 0, 0);">RGB</td><td>X</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        bottoms = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:bottom")
        assert len(bottoms) >= 1
        assert bottoms[0].get(f"{{{W_NS}}}color") == "FF0000", (
            f"expected FF0000 for rgb(255,0,0), got {bottoms[0].get(f'{{{W_NS}}}color')}"
        )


def test_cell_border_dotted():
    """A dotted border should produce w:val="dotted"."""
    html = """<table>
      <tr><td style="border-right:1.5pt dotted #3C78D8;">Dotted</td><td>X</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        rights = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:right")
        assert len(rights) >= 1
        assert rights[0].get(f"{{{W_NS}}}val") == "dotted"


def test_cell_border_double():
    """A double border should produce w:val="double"."""
    html = """<table>
      <tr><td style="border-bottom:1.5pt double #CC0000;">Double</td><td>X</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        bottoms = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders/w:bottom")
        assert len(bottoms) >= 1
        assert bottoms[0].get(f"{{{W_NS}}}val") == "double"


def test_colspan_preserved():
    """colspan=2 should produce <w:gridSpan w:val="2">."""
    html = """<table>
      <tr><td colspan="2" style="background-color:#F4CCCC;">Merged</td></tr>
      <tr><td>A</td><td>B</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        spans = _find_all(root, ".//w:tc/w:tcPr/w:gridSpan")
        vals = [el.get(f"{{{W_NS}}}val") for el in spans]
        assert "2" in vals, f"expected gridSpan val=2, got {vals}"


def test_rowspan_preserved():
    """rowspan=2 should produce vMerge restart + vMerge continue."""
    html = """<table>
      <tr><td rowspan="2" style="background-color:#FFF2CC;">Spanning</td><td>B1</td></tr>
      <tr><td>B2</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        vmerges = _find_all(root, ".//w:tc/w:tcPr/w:vMerge")
        vals = [el.get(f"{{{W_NS}}}val") for el in vmerges]
        assert "restart" in vals, f"expected vMerge restart, got {vals}"
        # Continuation cell has vMerge with no val attribute (or val="continue")
        assert any(v is None or v == "continue" for v in vals), f"expected vMerge continuation, got {vals}"


def test_inline_styles_inside_styled_cell():
    """Inline styles (<span style="...">) inside a styled cell should be preserved."""
    html = """<table>
      <tr>
        <td style="background-color:#D9EAD3;">
          <span style="font-weight:bold;color:#274E13;">Bold green text</span>
        </td>
        <td>Plain</td>
      </tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

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


def test_vertical_align_preserved():
    """vertical-align:top on a cell should produce <w:vAlign w:val="top"/>."""
    html = """<table>
      <tr><td style="vertical-align:top;">Top-aligned</td><td>X</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        valigns = _find_all(root, ".//w:tc/w:tcPr/w:vAlign")
        vals = [el.get(f"{{{W_NS}}}val") for el in valigns]
        assert "top" in vals, f"expected vAlign top, got {vals}"


def test_unstyled_table_passes_through():
    """A table with no styled cells should NOT be rewritten to raw OOXML —
    the default DOCX writer handles it instead."""
    html = """<table><tr><td>A</td><td>B</td></tr></table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        # Should still have a table
        tables = _find_all(root, ".//w:tbl")
        assert len(tables) >= 1, "table should still exist"
        # Content should be present
        texts = [el.text for el in _find_all(root, ".//w:t") if el.text]
        assert "A" in texts and "B" in texts


def test_disabled_by_default_without_metadata_flag():
    """Without -M preserve_table_styles=true, styled tables should NOT be
    rewritten — the feature is opt-in."""
    html = """<table>
      <tr><td style="background-color:#D9EAF7;">Styled</td><td>Plain</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out, preserve_table_styles=False)
        root = _parse_document_xml(out)

        # Cell background should NOT be present (default pandoc drops it)
        shd_els = _find_all(root, ".//w:tc/w:tcPr/w:shd")
        assert len(shd_els) == 0, f"expected no cell shading without opt-in, got {len(shd_els)}"


def test_text_align_center_preserved():
    """text-align:center on a cell should produce <w:jc w:val="center"/> on paragraphs."""
    html = """<table>
      <tr><td style="text-align:center;background-color:#F2F2F2;">Centered</td><td>X</td></tr>
    </table>"""
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(html, out)
        root = _parse_document_xml(out)

        jcs = _find_all(root, ".//w:tc//w:p/w:pPr/w:jc")
        vals = [el.get(f"{{{W_NS}}}val") for el in jcs]
        assert "center" in vals, f"expected jc center, got {vals}"


def test_full_test_html_file():
    """Smoke test: convert the real test.html file end-to-end."""
    test_html = Path(__file__).resolve().parents[1] / "tests" / "data" / "test" / "test.html"
    if not test_html.exists():
        pytest.skip("tests/data/test/test.html not found")
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.docx"
        _convert_html_to_docx(test_html.read_text(encoding="utf-8"), out)
        root = _parse_document_xml(out)

        # Should have tables with cell styling
        shd_els = _find_all(root, ".//w:tc/w:tcPr/w:shd")
        assert len(shd_els) > 0, "expected styled cells in the output"

        # Check that D9EAF7 (the header row background) is present
        fills = [el.get(f"{{{W_NS}}}fill") for el in shd_els]
        assert "D9EAF7" in fills, f"expected D9EAF7 header background, got {fills}"

        # Check borders exist
        borders = _find_all(root, ".//w:tc/w:tcPr/w:tcBorders")
        assert len(borders) > 0, "expected borders in styled tables"