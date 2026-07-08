"""Unit tests for ``app.DocxTablePreProcess``.

Each test builds a minimal DOCX zip in memory containing a single
``word/document.xml`` and inspects the rewritten XML — same synthetic-fixture
approach as the other preprocessor tests.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET  # noqa: S405

import pytest

from app import DocxTablePreProcess
from app.DocxTablePreProcess import SENTINEL_CLOSE, SENTINEL_OPEN

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)


def _pack(parts: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _doc(*fragments: str) -> bytes:
    body = "".join(fragments)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + "</w:body></w:document>").encode("utf-8")


def _table_cell(fill: str | None, text: str = "cell") -> str:
    """Build a ``<w:tc>`` with optional ``<w:shd w:fill="..."/>``."""
    tcpr = ""
    if fill is not None:
        tcpr = f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr>'
    return f"<w:tc>{tcpr}<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def _table(*cells_xml: str, grid_widths: list[str] | None = None) -> str:
    """Wrap cells in a single-row table with an optional tblGrid."""
    cells = "".join(cells_xml)
    grid = ""
    if grid_widths is not None:
        cols = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid_widths)
        grid = f"<w:tblGrid>{cols}</w:tblGrid>"
    return f"<w:tbl>{grid}<w:tr>{cells}</w:tr></w:tbl>"


def _table_with_bare_grid(*cells_xml: str, num_cols: int) -> str:
    """Table with <w:gridCol/> lacking w:w attributes (pandoc-breaking)."""
    cells = "".join(cells_xml)
    cols = "<w:gridCol/>" * num_cols
    return f"<w:tbl><w:tblGrid>{cols}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>"


def _plain_para(text: str = "para") -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def _body(blob: bytes) -> ET.Element:
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        return ET.fromstring(zf.read("word/document.xml"))  # noqa: S314


def _first_run_text(para: ET.Element) -> str | None:
    run = para.find(f"{{{W_NS}}}r")
    if run is None:
        return None
    t = run.find(f"{{{W_NS}}}t")
    return t.text if t is not None else None


def _sentinel(bg: str) -> str:
    return f"{SENTINEL_OPEN}bg={bg}{SENTINEL_CLOSE}"


# --- grid-column width fix --------------------------------------------------


def test_gridcol_without_width_gets_default_width():
    """<w:gridCol/> without w:w gets a default width so pandoc reads cells."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table_with_bare_grid(_table_cell(None, "a"), _table_cell(None, "b"), num_cols=2),
            )
        }
    )
    result = DocxTablePreProcess.preprocess(blob)
    root = _body(result)
    grid_cols = root.findall(f".//{{{W_NS}}}gridCol")
    assert len(grid_cols) == 2
    for col in grid_cols:
        w_val = col.get(f"{{{W_NS}}}w")
        assert w_val is not None, "gridCol should have w:w after preprocessing"
        assert int(w_val) > 0


def test_gridcol_with_existing_width_is_not_changed():
    """<w:gridCol w:w="4500"/> keeps its original width."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell(None, "a"), _table_cell(None, "b"), grid_widths=["4500", "4500"]),
            )
        }
    )
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_mixed_gridcol_widths_only_fills_missing():
    """Only gridCols without w:w get the default; existing widths are kept."""
    tbl_xml = '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    blob = _pack({"word/document.xml": _doc(tbl_xml)})
    result = DocxTablePreProcess.preprocess(blob)
    root = _body(result)
    grid_cols = root.findall(f".//{{{W_NS}}}gridCol")
    assert grid_cols[0].get(f"{{{W_NS}}}w") == "3000"  # unchanged
    assert grid_cols[1].get(f"{{{W_NS}}}w") is not None  # was filled in


def test_table_without_tblgrid_gets_one_created():
    """A table with no <w:tblGrid> at all gets a grid with one positive-width
    column per cell, so pandoc keeps the cells and computes column widths
    (without a grid it drops the cells entirely and the table renders empty)."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell(None, "a"), _table_cell(None, "b")),
            )
        }
    )
    root = _body(DocxTablePreProcess.preprocess(blob))
    grid_cols = root.findall(f".//{{{W_NS}}}gridCol")
    assert len(grid_cols) == 2
    assert all(int(c.get(f"{{{W_NS}}}w")) > 0 for c in grid_cols)


def test_grid_is_padded_to_column_count():
    """A grid with fewer <w:gridCol> than the row's cells is padded so pandoc
    sees every column (a short grid otherwise yields a partial, narrow table)."""
    tbl = '<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    root = _body(DocxTablePreProcess.preprocess(_pack({"word/document.xml": _doc(tbl)})))
    grid_cols = root.findall(f".//{{{W_NS}}}gridCol")
    assert len(grid_cols) == 3
    assert all(int(c.get(f"{{{W_NS}}}w")) > 0 for c in grid_cols)


# --- cell background tagging ------------------------------------------------


def test_cell_with_background_gets_sentinel():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell("D9EAF7", "Hello"), grid_widths=["4800"]),
            )
        }
    )
    result = DocxTablePreProcess.preprocess(blob)
    para = _body(result).find(f".//{{{W_NS}}}p")
    assert _first_run_text(para) == _sentinel("D9EAF7")
    texts = [t.text for t in para.iter(f"{{{W_NS}}}t")]
    assert texts == [_sentinel("D9EAF7"), "Hello"]


def test_lowercase_hex_is_uppercased():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell("d9eaf7", "x"), grid_widths=["4800"]),
            )
        }
    )
    para = _body(DocxTablePreProcess.preprocess(blob)).find(f".//{{{W_NS}}}p")
    assert _first_run_text(para) == _sentinel("D9EAF7")


def test_multiple_cells_tagged_independently():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(
                    _table_cell("F4CCCC", "a"),
                    _table_cell("D9EAD3", "b"),
                    grid_widths=["4800", "4800"],
                )
            )
        }
    )
    paras = _body(DocxTablePreProcess.preprocess(blob)).findall(f".//{{{W_NS}}}p")
    assert _first_run_text(paras[0]) == _sentinel("F4CCCC")
    assert _first_run_text(paras[1]) == _sentinel("D9EAD3")


def test_sentinel_inserted_after_ppr():
    """When the cell's first paragraph has <w:pPr>, sentinel goes after it."""
    cell_xml = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="AABBCC"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell_xml}</w:tr></w:tbl>')})
    para = _body(DocxTablePreProcess.preprocess(blob)).find(f".//{{{W_NS}}}p")
    children = list(para)
    assert children[0].tag == f"{{{W_NS}}}pPr"
    assert children[1].tag == f"{{{W_NS}}}r"  # sentinel run right after pPr
    assert _first_run_text(para) == _sentinel("AABBCC")


# --- pass-through -----------------------------------------------------------


def test_cell_without_tcpr_is_untouched():
    """Table with grid widths but no cell formatting — no change needed."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell(None, "plain"), grid_widths=["4800"]),
            )
        }
    )
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_white_background_is_skipped():
    """FFFFFF is the page colour — treat as no background."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell("FFFFFF", "x"), grid_widths=["4800"]),
            )
        }
    )
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_auto_fill_is_skipped():
    cell = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>')})
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_cell_without_shd_is_untouched():
    cell = '<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>')})
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_document_without_tables_returned_unchanged():
    blob = _pack({"word/document.xml": _doc(_plain_para("hello"))})
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_mixed_doc_only_tags_styled_cells():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _plain_para("intro"),
                _table(_table_cell("F4CCCC", "colored"), _table_cell(None, "plain"), grid_widths=["4800", "4800"]),
            )
        }
    )
    result = DocxTablePreProcess.preprocess(blob)
    root = _body(result)
    paras = root.findall(f".//{{{W_NS}}}p")
    assert _first_run_text(paras[0]) == "intro"
    assert _first_run_text(paras[1]) == _sentinel("F4CCCC")
    assert _first_run_text(paras[2]) == "plain"


# --- robustness -------------------------------------------------------------


def test_preprocess_is_idempotent():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell("AABBCC", "x"), grid_widths=["4800"]),
            )
        }
    )
    once = DocxTablePreProcess.preprocess(blob)
    assert DocxTablePreProcess.preprocess(once) == once
    para = _body(once).find(f".//{{{W_NS}}}p")
    assert [t.text for t in para.iter(f"{{{W_NS}}}t")] == [_sentinel("AABBCC"), "x"]


def test_gridcol_fix_is_idempotent():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table_with_bare_grid(_table_cell(None, "x"), num_cols=1),
            )
        }
    )
    once = DocxTablePreProcess.preprocess(blob)
    assert DocxTablePreProcess.preprocess(once) == once


def test_non_docx_input_is_returned_unchanged():
    assert DocxTablePreProcess.preprocess(b"not a zip") == b"not a zip"


def test_zip_without_body_parts_returned_unchanged():
    blob = _pack({"word/styles.xml": b"<x/>", "docProps/core.xml": b"<x/>"})
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_malformed_document_xml_returned_unchanged():
    blob = _pack({"word/document.xml": b"<w:document><unclosed>"})
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_invalid_hex_fill_is_ignored():
    """Non-hex fill values (theme refs, named colours) are not encoded."""
    cell = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="accent1"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>')})
    assert DocxTablePreProcess.preprocess(blob) == blob


def test_both_gridcol_fix_and_sentinel_applied():
    """Bare gridCols + coloured cell: both fixes applied in one pass."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table_with_bare_grid(_table_cell("AABBCC", "x"), _table_cell(None, "y"), num_cols=2),
            )
        }
    )
    result = DocxTablePreProcess.preprocess(blob)
    root = _body(result)
    # Grid widths filled in
    grid_cols = root.findall(f".//{{{W_NS}}}gridCol")
    assert all(col.get(f"{{{W_NS}}}w") is not None for col in grid_cols)
    # Sentinel injected
    paras = root.findall(f".//{{{W_NS}}}p")
    assert _first_run_text(paras[0]) == _sentinel("AABBCC")
    assert _first_run_text(paras[1]) == "y"


def test_real_docx_with_table_styles():
    """End-to-end smoke test against the real test fixture."""
    fixture = Path("tests/data/test/test_convert_live_doc_with_table_with_inline_style.docx")
    if not fixture.exists():
        pytest.skip("Fixture not available: test_convert_live_doc_with_table_with_inline_style.docx")

    blob = fixture.read_bytes()
    result = DocxTablePreProcess.preprocess(blob)
    assert result != blob

    root = _body(result)
    # Grid columns now have widths
    for col in root.iter(f"{{{W_NS}}}gridCol"):
        assert col.get(f"{{{W_NS}}}w") is not None

    # At least one sentinel was injected
    sentinel_found = any(t.text and SENTINEL_OPEN in t.text for t in root.iter(f"{{{W_NS}}}t"))
    assert sentinel_found, "Expected at least one table-cell sentinel in the preprocessed DOCX"


# --- table width & alignment tagging (DOCX -> PDF) --------------------------


def _table_with_pr(tblpr_inner: str, *cells_xml: str, grid_widths: list[str] | None = None) -> str:
    """Single-row table with a ``<w:tblPr>`` (tblW/jc live here)."""
    cells = "".join(cells_xml)
    grid = ""
    if grid_widths is not None:
        cols = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid_widths)
        grid = f"<w:tblGrid>{cols}</w:tblGrid>"
    return f"<w:tbl><w:tblPr>{tblpr_inner}</w:tblPr>{grid}<w:tr>{cells}</w:tr></w:tbl>"


def _first_cell_sentinel_kv(root: ET.Element) -> dict[str, str]:
    """Parse the leading sentinel of the table's first cell into a key map."""
    tc = root.find(f".//{{{W_NS}}}tc")
    assert tc is not None
    para = tc.find(f"{{{W_NS}}}p")
    text = _first_run_text(para)
    if not text or not text.startswith(SENTINEL_OPEN):
        return {}
    payload = text[len(SENTINEL_OPEN) : text.find(SENTINEL_CLOSE)]
    return dict(seg.split("=", 1) for seg in payload.split(";") if "=" in seg)


def test_percentage_width_tagged_on_first_cell():
    tblpr = '<w:tblW w:w="2000" w:type="pct"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob)))
    assert kv.get("tw") == "0.4000"
    assert kv.get("ta") == "left"


def test_full_width_table_is_tagged():
    """A 100% table IS tagged (tw=1.0): pandoc can still render it content-width
    and centered when the DOCX has no usable column widths, so the filter must
    force it full-width and flush-left."""
    tblpr = '<w:tblW w:w="5000" w:type="pct"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob)))
    assert kv.get("tw") == "1.0000"
    assert kv.get("ta") == "left"


def test_center_and_right_alignment_carried():
    for val, expected in (("center", "center"), ("right", "right")):
        tblpr = f'<w:tblW w:w="1250" w:type="pct"/><w:jc w:val="{val}"/>'
        blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
        kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob)))
        assert kv.get("tw") == "0.2500"
        assert kv.get("ta") == expected


def test_absolute_dxa_width_becomes_fraction():
    """50px == 750 twips against the 9360-twip reference ~= 0.08 of the line."""
    tblpr = '<w:tblW w:w="750" w:type="dxa"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob)))
    assert kv.get("tw") == "0.0801"


def test_table_width_merges_with_cell_background():
    """A narrow table whose first cell is also shaded gets one merged sentinel."""
    tblpr = '<w:tblW w:w="2000" w:type="pct"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell("F0F0F0", "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob)))
    assert kv.get("tw") == "0.4000"
    assert kv.get("ta") == "left"
    assert kv.get("bg") == "F0F0F0"


def test_table_without_tblw_carries_alignment_only():
    """No <w:tblW> -> no width fraction, but the alignment is still carried so a
    left/right table is not left to pandoc's centered default."""
    blob = _pack({"word/document.xml": _doc(_table_with_pr('<w:jc w:val="left"/>', _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob)))
    assert "tw" not in kv
    assert kv.get("ta") == "left"


def test_table_with_no_layout_at_all_is_not_tagged():
    """A table with neither <w:tblW> nor <w:jc> gets no layout sentinel."""
    blob = _pack({"word/document.xml": _doc(_table_with_pr("", _table_cell(None, "a"), grid_widths=["4800"]))})
    assert _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(blob))) == {}


def test_zero_width_gridcols_are_normalised():
    """<w:gridCol w:w="0"/> is treated as missing and given a positive width, so
    pandoc keeps column widths (and the table renders at its tblW width) instead
    of collapsing to a content-width, centered table."""
    tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="left"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="0"/><w:gridCol w:w="0"/></w:tblGrid>'
        "<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    root = _body(DocxTablePreProcess.preprocess(_pack({"word/document.xml": _doc(tbl)})))
    widths = [int(c.get(f"{{{W_NS}}}w")) for c in root.iter(f"{{{W_NS}}}gridCol")]
    assert all(w > 0 for w in widths)


# --- caption style neutralisation (avoid double numbering in PDF) -----------


def test_caption_paragraph_style_is_stripped():
    """A ``Caption``-styled paragraph loses its style so pandoc won't turn it
    into an auto-numbered LaTeX \\caption (its text already has the number)."""
    para = '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Table 1 My caption</w:t></w:r></w:p>'
    blob = _pack({"word/document.xml": _doc(para)})
    root = _body(DocxTablePreProcess.preprocess(blob))
    styles = [s.get(f"{{{W_NS}}}val") for s in root.iter(f"{{{W_NS}}}pStyle")]
    assert "Caption" not in styles
    # The text is preserved.
    assert any(t.text == "Table 1 My caption" for t in root.iter(f"{{{W_NS}}}t"))


def test_non_caption_paragraph_styles_are_left_alone():
    """Pandoc's own caption styles (with no embedded number) keep their style."""
    for style in ("TableCaption", "ImageCaption", "BodyText"):
        para = f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
        blob = _pack({"word/document.xml": _doc(para)})
        root = _body(DocxTablePreProcess.preprocess(blob))
        styles = [s.get(f"{{{W_NS}}}val") for s in root.iter(f"{{{W_NS}}}pStyle")]
        assert style in styles


def test_absolute_width_table_flagged_for_tight_padding():
    """A dxa (absolute px/pt) width carries aw=1 so the filter tightens its
    inter-column padding; a percentage width does not."""
    dxa = '<w:tblW w:w="750" w:type="dxa"/><w:jc w:val="left"/>'
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(_pack({"word/document.xml": _doc(_table_with_pr(dxa, _table_cell(None, "a"), grid_widths=["4800"]))}))))
    assert kv.get("aw") == "1"

    pct = '<w:tblW w:w="2000" w:type="pct"/><w:jc w:val="left"/>'
    kv = _first_cell_sentinel_kv(_body(DocxTablePreProcess.preprocess(_pack({"word/document.xml": _doc(_table_with_pr(pct, _table_cell(None, "a"), grid_widths=["4800"]))}))))
    assert "aw" not in kv
