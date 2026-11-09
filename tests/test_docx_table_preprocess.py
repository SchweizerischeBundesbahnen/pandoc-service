"""Unit tests for ``app.docx_table_pre_process``.

Each test builds a minimal DOCX zip in memory containing a single
``word/document.xml`` and inspects the rewritten XML — same synthetic-fixture
approach as the other preprocessor tests.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from app import docx_table_pre_process
from app.docx_table_pre_process import SENTINEL_CLOSE, SENTINEL_OPEN

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
        return ET.fromstring(zf.read("word/document.xml"))


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
    result = docx_table_pre_process.preprocess(blob)
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
    assert docx_table_pre_process.preprocess(blob) == blob


def test_mixed_gridcol_widths_only_fills_missing():
    """Only gridCols without w:w get the default; existing widths are kept."""
    tbl_xml = '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    blob = _pack({"word/document.xml": _doc(tbl_xml)})
    result = docx_table_pre_process.preprocess(blob)
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
    root = _body(docx_table_pre_process.preprocess(blob))
    grid_cols = root.findall(f".//{{{W_NS}}}gridCol")
    assert len(grid_cols) == 2
    assert all(int(c.get(f"{{{W_NS}}}w")) > 0 for c in grid_cols)


def test_grid_is_padded_to_column_count():
    """A grid with fewer <w:gridCol> than the row's cells is padded so pandoc
    sees every column (a short grid otherwise yields a partial, narrow table)."""
    tbl = '<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    root = _body(docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(tbl)})))
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
    result = docx_table_pre_process.preprocess(blob)
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
    para = _body(docx_table_pre_process.preprocess(blob)).find(f".//{{{W_NS}}}p")
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
    paras = _body(docx_table_pre_process.preprocess(blob)).findall(f".//{{{W_NS}}}p")
    assert _first_run_text(paras[0]) == _sentinel("F4CCCC")
    assert _first_run_text(paras[1]) == _sentinel("D9EAD3")


def test_sentinel_inserted_after_ppr():
    """When the cell's first paragraph has <w:pPr>, sentinel goes after it."""
    cell_xml = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="AABBCC"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell_xml}</w:tr></w:tbl>')})
    para = _body(docx_table_pre_process.preprocess(blob)).find(f".//{{{W_NS}}}p")
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
    assert docx_table_pre_process.preprocess(blob) == blob


def test_white_background_is_skipped():
    """FFFFFF is the page colour — treat as no background."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table(_table_cell("FFFFFF", "x"), grid_widths=["4800"]),
            )
        }
    )
    assert docx_table_pre_process.preprocess(blob) == blob


def test_auto_fill_is_skipped():
    cell = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>')})
    assert docx_table_pre_process.preprocess(blob) == blob


def test_cell_without_shd_is_untouched():
    cell = '<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>')})
    assert docx_table_pre_process.preprocess(blob) == blob


def test_document_without_tables_returned_unchanged():
    blob = _pack({"word/document.xml": _doc(_plain_para("hello"))})
    assert docx_table_pre_process.preprocess(blob) == blob


def test_mixed_doc_only_tags_styled_cells():
    blob = _pack(
        {
            "word/document.xml": _doc(
                _plain_para("intro"),
                _table(_table_cell("F4CCCC", "colored"), _table_cell(None, "plain"), grid_widths=["4800", "4800"]),
            )
        }
    )
    result = docx_table_pre_process.preprocess(blob)
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
    once = docx_table_pre_process.preprocess(blob)
    assert docx_table_pre_process.preprocess(once) == once
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
    once = docx_table_pre_process.preprocess(blob)
    assert docx_table_pre_process.preprocess(once) == once


def test_non_docx_input_is_returned_unchanged():
    assert docx_table_pre_process.preprocess(b"not a zip") == b"not a zip"


def test_zip_without_body_parts_returned_unchanged():
    blob = _pack({"word/styles.xml": b"<x/>", "docProps/core.xml": b"<x/>"})
    assert docx_table_pre_process.preprocess(blob) == blob


def test_malformed_document_xml_returned_unchanged():
    blob = _pack({"word/document.xml": b"<w:document><unclosed>"})
    assert docx_table_pre_process.preprocess(blob) == blob


def test_invalid_hex_fill_is_ignored():
    """Non-hex fill values (theme refs, named colours) are not encoded."""
    cell = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="accent1"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>'
    blob = _pack({"word/document.xml": _doc(f'<w:tbl><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>')})
    assert docx_table_pre_process.preprocess(blob) == blob


def test_both_gridcol_fix_and_sentinel_applied():
    """Bare gridCols + coloured cell: both fixes applied in one pass."""
    blob = _pack(
        {
            "word/document.xml": _doc(
                _table_with_bare_grid(_table_cell("AABBCC", "x"), _table_cell(None, "y"), num_cols=2),
            )
        }
    )
    result = docx_table_pre_process.preprocess(blob)
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
    result = docx_table_pre_process.preprocess(blob)
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
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob)))
    assert kv.get("tw") == "0.4000"
    assert kv.get("ta") == "left"


def test_full_width_table_is_tagged():
    """A 100% table IS tagged (tw=1.0): pandoc can still render it content-width
    and centered when the DOCX has no usable column widths, so the filter must
    force it full-width and flush-left."""
    tblpr = '<w:tblW w:w="5000" w:type="pct"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob)))
    assert kv.get("tw") == "1.0000"
    assert kv.get("ta") == "left"


def test_center_and_right_alignment_carried():
    for val, expected in (("center", "center"), ("right", "right")):
        tblpr = f'<w:tblW w:w="1250" w:type="pct"/><w:jc w:val="{val}"/>'
        blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
        kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob)))
        assert kv.get("tw") == "0.2500"
        assert kv.get("ta") == expected


def test_absolute_dxa_width_becomes_fraction():
    """50px == 750 twips against the 9360-twip reference ~= 0.08 of the line."""
    tblpr = '<w:tblW w:w="750" w:type="dxa"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob)))
    assert kv.get("tw") == "0.0801"


def test_table_width_merges_with_cell_background():
    """A narrow table whose first cell is also shaded gets one merged sentinel."""
    tblpr = '<w:tblW w:w="2000" w:type="pct"/><w:jc w:val="left"/>'
    blob = _pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell("F0F0F0", "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob)))
    assert kv.get("tw") == "0.4000"
    assert kv.get("ta") == "left"
    assert kv.get("bg") == "F0F0F0"


def test_table_without_tblw_carries_alignment_only():
    """No <w:tblW> -> no width fraction, but the alignment is still carried so a
    left/right table is not left to pandoc's centered default."""
    blob = _pack({"word/document.xml": _doc(_table_with_pr('<w:jc w:val="left"/>', _table_cell(None, "a"), grid_widths=["4800"]))})
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob)))
    assert "tw" not in kv
    assert kv.get("ta") == "left"


def test_table_with_no_layout_at_all_is_not_tagged():
    """A table with neither <w:tblW> nor <w:jc> gets no layout sentinel."""
    blob = _pack({"word/document.xml": _doc(_table_with_pr("", _table_cell(None, "a"), grid_widths=["4800"]))})
    assert _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(blob))) == {}


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
    root = _body(docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(tbl)})))
    widths = [int(c.get(f"{{{W_NS}}}w")) for c in root.iter(f"{{{W_NS}}}gridCol")]
    assert all(w > 0 for w in widths)


# --- caption style neutralisation (avoid double numbering in PDF) -----------


def test_caption_paragraph_style_is_stripped():
    """A ``Caption``-styled paragraph loses its style so pandoc won't turn it
    into an auto-numbered LaTeX \\caption (its text already has the number)."""
    para = '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Table 1 My caption</w:t></w:r></w:p>'
    blob = _pack({"word/document.xml": _doc(para)})
    root = _body(docx_table_pre_process.preprocess(blob))
    styles = [s.get(f"{{{W_NS}}}val") for s in root.iter(f"{{{W_NS}}}pStyle")]
    assert "Caption" not in styles
    # The text is preserved.
    assert any(t.text == "Table 1 My caption" for t in root.iter(f"{{{W_NS}}}t"))


def test_non_caption_paragraph_styles_are_left_alone():
    """Pandoc's own caption styles (with no embedded number) keep their style."""
    for style in ("TableCaption", "ImageCaption", "BodyText"):
        para = f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
        blob = _pack({"word/document.xml": _doc(para)})
        root = _body(docx_table_pre_process.preprocess(blob))
        styles = [s.get(f"{{{W_NS}}}val") for s in root.iter(f"{{{W_NS}}}pStyle")]
        assert style in styles


def test_absolute_width_table_flagged_for_tight_padding():
    """A dxa (absolute px/pt) width carries aw=1 so the filter tightens its
    inter-column padding; a percentage width does not."""
    dxa = '<w:tblW w:w="750" w:type="dxa"/><w:jc w:val="left"/>'
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table_with_pr(dxa, _table_cell(None, "a"), grid_widths=["4800"]))}))))
    assert kv.get("aw") == "1"

    pct = '<w:tblW w:w="2000" w:type="pct"/><w:jc w:val="left"/>'
    kv = _first_cell_sentinel_kv(_body(docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table_with_pr(pct, _table_cell(None, "a"), grid_widths=["4800"]))}))))
    assert "aw" not in kv


def test_every_caption_keeps_a_style_of_its_own():
    """A Polarion caption (literal number) and a Word caption (SEQ field) both become PandocCaption,
    so pandoc neither numbers them again nor moves them away from where Word prints them."""
    polarion = '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Table 1 My caption</w:t></w:r></w:p>'
    word = (
        '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>'
        "<w:r><w:t>Table </w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> SEQ Table \\* ARABIC </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "<w:r><w:t>: My caption</w:t></w:r></w:p>"
    )
    for caption in (polarion, word):
        root = _body(docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(caption)})))
        assert [s.get(f"{{{W_NS}}}val") for s in root.iter(f"{{{W_NS}}}pStyle")] == ["PandocCaption"]


def test_pandoc_captions_are_left_alone():
    caption = '<w:p><w:pPr><w:pStyle w:val="ImageCaption"/></w:pPr><w:r><w:t>A picture</w:t></w:r></w:p>'
    root = _body(docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(caption)})))
    assert [s.get(f"{{{W_NS}}}val") for s in root.iter(f"{{{W_NS}}}pStyle")] == ["ImageCaption"]


# --- cell justification -----------------------------------------------------
#
# Pandoc's AST has no justified alignment, so a cell's <w:jc w:val="both">
# reaches the reader as AlignLeft and "distribute" as AlignDefault. Either way
# filters/docx/docx_tables_to_latex.lua would flush the cell left, so the sentinel
# carries it. The other three alignments survive on the Cell and are not tagged.


def _justified_cell(jc: str | None, fill: str | None = None, text: str = "cell") -> str:
    tcpr = f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr>' if fill else ""
    ppr = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ""
    return f"<w:tc>{tcpr}<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def _cell_sentinel_text(blob: bytes, index: int = 0) -> str:
    cells = _body(blob).iter(f"{{{W_NS}}}tc")
    para = next(iter(list(cells)[index].iter(f"{{{W_NS}}}p")))
    return _first_run_text(para) or ""


@pytest.mark.parametrize("jc", ["both", "distribute"])
def test_justified_cell_is_tagged(jc):
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(_justified_cell(jc)))}))

    assert "ha=justify" in _cell_sentinel_text(out)


@pytest.mark.parametrize("jc", ["left", "center", "right", None])
def test_other_alignments_are_not_tagged(jc):
    """They survive on the Cell, so tagging them would be dead weight."""
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(_justified_cell(jc)))}))

    assert "ha=" not in _cell_sentinel_text(out)


def test_justified_cell_without_a_background_is_still_tagged():
    """The sentinel used to be emitted only for a cell with a fill."""
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(_justified_cell("both")))}))

    text = _cell_sentinel_text(out)
    assert text.startswith(SENTINEL_OPEN)
    assert "ha=justify" in text
    assert "bg=" not in text


def test_justification_and_background_share_one_sentinel():
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(_justified_cell("both", fill="F2F2F2")))}))

    text = _cell_sentinel_text(out)
    assert text.count(SENTINEL_OPEN) == 1
    assert "ha=justify" in text
    assert "bg=F2F2F2" in text


def test_tagging_a_justified_cell_is_idempotent():
    packed = _pack({"word/document.xml": _doc(_table(_justified_cell("both", fill="F2F2F2")))})
    once = docx_table_pre_process.preprocess(packed)

    assert _cell_sentinel_text(docx_table_pre_process.preprocess(once)) == _cell_sentinel_text(once)


def test_nested_table_justification_does_not_leak_to_the_outer_cell():
    """A cell that opens with a nested table has no paragraph of its own.

    A recursive walk would take the inner cell's first paragraph, attribute its
    alignment to the outer cell, and then make the outer cell grow a paragraph
    purely to hold the sentinel.
    """
    inner = _table(_justified_cell("both", text="inner"))
    outer = f"<w:tbl><w:tr><w:tc>{inner}<w:p/></w:tc></w:tr></w:tbl>"
    source = _pack({"word/document.xml": _doc(outer)})

    out = docx_table_pre_process.preprocess(source)

    body = _body(out)
    cells = list(body.iter(f"{{{W_NS}}}tc"))
    outer_tc, inner_tc = cells[0], cells[1]

    assert [c.tag.split("}")[-1] for c in outer_tc] == ["tbl", "p"], "the outer cell grew a paragraph"
    assert SENTINEL_OPEN not in "".join(t.text or "" for p in outer_tc.findall(f"{{{W_NS}}}p") for t in p.iter(f"{{{W_NS}}}t")), "the outer cell was tagged from the inner cell's paragraph"
    assert "ha=justify" in (_first_run_text(inner_tc.find(f"{{{W_NS}}}p")) or ""), "the inner cell lost its own tag"


def test_cell_with_a_leading_paragraph_before_a_nested_table_is_read_from_it():
    """The cell's own paragraph still counts when it precedes the nested table."""
    inner = _table(_justified_cell(None, text="inner"))
    outer = f'<w:tbl><w:tr><w:tc><w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>outer</w:t></w:r></w:p>{inner}<w:p/></w:tc></w:tr></w:tbl>'
    source = _pack({"word/document.xml": _doc(outer)})

    out = docx_table_pre_process.preprocess(source)

    outer_tc = next(_body(out).iter(f"{{{W_NS}}}tc"))
    assert "ha=justify" in (_first_run_text(outer_tc.find(f"{{{W_NS}}}p")) or "")


@pytest.mark.parametrize(
    ("section", "fraction"),
    [
        # Letter landscape, 1 inch margins: 15840 - 2 x 1440 = 12960 twips of text
        ('<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr>', "0.3472"),
        # A4 portrait, 2 cm margins: 11906 - 2 x 1134 = 9638 twips of text
        ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1134" w:right="1134"/></w:sectPr>', "0.4669"),
        # No section: Letter portrait, 9360 twips of text
        ("", "0.4808"),
        # Margins which leave no text: Letter portrait stands in
        ('<w:sectPr><w:pgSz w:w="2880"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr>', "0.4808"),
    ],
)
def test_an_absolute_width_is_a_share_of_the_text_of_the_first_section(section, fraction):
    """LaTeX sets the PDF at the page of the first section, a landscape one too; #279."""
    tblpr = '<w:tblW w:w="4500" w:type="dxa"/>'
    blob = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table_with_pr(tblpr, _table_cell(None, "a"), grid_widths=["4500"]), section)}))

    assert _first_cell_sentinel_kv(_body(blob))["tw"] == fraction


def test_a_table_in_the_footnotes_takes_the_line_of_the_document():
    """The footnotes have no section of their own; LaTeX sets them at the line of the body."""
    landscape = '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr>'
    table = _table_with_pr('<w:tblW w:w="4500" w:type="dxa"/>', _table_cell(None, "a"), grid_widths=["4500"])
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1">' + table + "</w:footnote></w:footnotes>").encode()

    blob = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_plain_para(), landscape), "word/footnotes.xml": footnotes}))

    with zipfile.ZipFile(io.BytesIO(blob)) as archive:
        footnotes_tree = ET.fromstring(archive.read("word/footnotes.xml"))
    assert _first_cell_sentinel_kv(footnotes_tree)["tw"] == "0.3472"


# ---- The lines of a styled table ----

W_DECL = f'xmlns:w="{W_NS}"'


def _styled_table(tblpr_extra: str = "", header: bool = False, cell_borders: str = "") -> str:
    """A table of pandoc's Table style, with an optional header row and cell borders."""
    header_row = '<w:tr><w:trPr><w:tblHeader w:val="on"/></w:trPr><w:tc><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:tc></w:tr>' if header else ""
    tcpr = f"<w:tcPr>{cell_borders}</w:tcPr>" if cell_borders else ""
    return f'<w:tbl><w:tblPr><w:tblStyle w:val="Table"/>{tblpr_extra}</w:tblPr>{header_row}<w:tr><w:tc>{tcpr}<w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'


def _rules_of(table: str, bordered: frozenset[str] = frozenset()) -> str | None:
    xml, _ = docx_table_pre_process.rewrite_part(f"<w:document {W_DECL}><w:body>{table}</w:body></w:document>".encode(), 9360, bordered)
    text = "".join(t.text or "" for t in ET.fromstring(xml).iter(f"{{{W_NS}}}t"))
    if SENTINEL_OPEN not in text:
        return None
    payload = text[text.index(SENTINEL_OPEN) + 1 : text.index(SENTINEL_CLOSE)]
    return dict(item.split("=") for item in payload.split(";")).get("rl")


def test_a_styled_table_without_header_row_draws_no_lines():
    assert _rules_of(_styled_table()) == "none"


def test_the_first_row_of_a_header_draws_the_line_under_it():
    assert _rules_of(_styled_table('<w:tblLook w:firstRow="1"/>', header=True)) == "head"


def test_the_first_row_bit_of_the_look_counts_as_well():
    assert _rules_of(_styled_table('<w:tblLook w:val="0020"/>', header=True)) == "head"


def test_a_look_value_which_is_no_number_is_no_first_row():
    assert _rules_of(_styled_table('<w:tblLook w:val="zz"/>', header=True)) == "none"


def test_a_first_row_look_without_a_header_row_draws_no_line():
    assert _rules_of(_styled_table('<w:tblLook w:firstRow="1"/>')) == "none"


def test_a_table_with_borders_keeps_the_rules_of_latex():
    borders = '<w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders>'
    assert _rules_of(_styled_table(borders)) is None


def test_borders_of_nil_draw_nothing():
    borders = '<w:tblBorders><w:top w:val="nil"/><w:insideH w:val="none"/></w:tblBorders>'
    assert _rules_of(_styled_table(borders)) == "none"


def test_a_cell_with_borders_keeps_the_rules_of_latex():
    assert _rules_of(_styled_table(cell_borders='<w:tcBorders><w:left w:val="single"/></w:tcBorders>')) is None


def test_a_style_with_borders_keeps_the_rules_of_latex():
    assert _rules_of(_styled_table(), bordered=frozenset({"Table"})) is None


def test_the_bordered_table_styles_are_read_from_the_styles():
    styles = ET.fromstring(
        f"<w:styles {W_DECL}>"
        '<w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblBorders><w:insideV w:val="single"/></w:tblBorders></w:tblPr></w:style>'
        '<w:style w:type="table" w:styleId="Table"><w:tblPr><w:tblCellMar/></w:tblPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Normal"/>'
        "</w:styles>"
    )
    assert docx_table_pre_process.bordered_table_styles(styles) == {"Grid"}
    assert docx_table_pre_process.bordered_table_styles(None) == frozenset()


def _first_cell_kv(table: str) -> dict[str, str]:
    xml, _ = docx_table_pre_process.rewrite_part(f"<w:document {W_DECL}><w:body>{table}</w:body></w:document>".encode(), 9360)
    text = "".join(t.text or "" for t in ET.fromstring(xml).iter(f"{{{W_NS}}}t"))
    if SENTINEL_OPEN not in text:
        return {}
    payload = text[text.index(SENTINEL_OPEN) + 1 : text.index(SENTINEL_CLOSE)]
    return dict(item.split("=") for item in payload.split(";"))


def test_a_table_of_compact_paragraphs_is_marked():
    compact = '<w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Compact"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    assert _first_cell_kv(compact)["cc"] == "1"


def test_a_table_of_paragraphs_without_a_style_is_not_marked():
    assert "cc" not in _first_cell_kv("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>")


GRID_BORDERS = "".join(f'<w:{side} w:val="single" w:sz="6" w:color="CCCCCC"/>' for side in ("top", "left", "bottom", "right"))


def _grid_cell(borders: str = GRID_BORDERS) -> str:
    return f"<w:tc><w:tcPr><w:tcBorders>{borders}</w:tcBorders></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"


def test_a_table_whose_cells_draw_every_side_is_a_grid():
    kv = _first_cell_kv(f"<w:tbl><w:tr>{_grid_cell()}</w:tr></w:tbl>")
    assert (kv["rl"], kv["lc"], kv["lw"]) == ("grid", "CCCCCC", "0.75")


def test_the_borders_of_the_table_stand_in_for_those_of_the_cell():
    borders = "".join(f'<w:{side} w:val="single"/>' for side in ("top", "left", "insideH", "insideV"))
    kv = _first_cell_kv(f"<w:tbl><w:tblPr><w:tblBorders>{borders}</w:tblBorders></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>")
    assert (kv["rl"], kv["lc"], kv["lw"]) == ("grid", "000000", "0.5")


def test_a_cell_without_one_side_is_no_grid():
    partial = GRID_BORDERS.replace('<w:left w:val="single" w:sz="6" w:color="CCCCCC"/>', '<w:left w:val="nil"/>')
    assert "rl" not in _first_cell_kv(f"<w:tbl><w:tr>{_grid_cell(partial)}</w:tr></w:tbl>")


def test_white_lines_are_no_grid():
    assert "rl" not in _first_cell_kv(f"<w:tbl><w:tr>{_grid_cell(GRID_BORDERS.replace('CCCCCC', 'ffffff'))}</w:tr></w:tbl>")


def test_the_width_of_a_line_has_bounds():
    assert _first_cell_kv(f"<w:tbl><w:tr>{_grid_cell(GRID_BORDERS.replace('w:sz="6"', 'w:sz="0"'))}</w:tr></w:tbl>")["lw"] == "0.25"
    assert _first_cell_kv(f"<w:tbl><w:tr>{_grid_cell(GRID_BORDERS.replace('w:sz="6"', 'w:sz="x"'))}</w:tr></w:tbl>")["lw"] == "0.5"


def _bordered_table(rows: list[list[str]], table_borders: str) -> bytes:
    """A table of the given rows of cells, each the inside of a <w:tc>, with the table's borders."""
    body = "".join("<w:tr>" + "".join(f"<w:tc>{cell}</w:tc>" for cell in row) + "</w:tr>" for row in rows)
    grid = "".join('<w:gridCol w:w="2000"/>' for _ in rows[0])
    return (f'<w:document xmlns:w="{W_NS}"><w:body><w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>{table_borders}</w:tblBorders></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl></w:body></w:document>').encode()


def _sides(eighths: int = 4, color: str = "000000", style: str = "single") -> str:
    return "".join(f'<w:{side} w:val="{style}" w:sz="{eighths}" w:space="0" w:color="{color}"/>' for side in ("top", "left", "bottom", "right", "insideH", "insideV"))


def _cell_borders(xml: bytes) -> list[str | None]:
    out, _ = docx_table_pre_process.rewrite_part(xml, 9000)
    texts = [t.text or "" for t in ET.fromstring(out).iter(f"{{{W_NS}}}t")]
    found = []
    for text in texts:
        if text.startswith(SENTINEL_OPEN):
            payload = text[len(SENTINEL_OPEN) : text.index(SENTINEL_CLOSE)]
            found.append(dict(item.split("=", 1) for item in payload.split(";")).get("bd"))
    return found


def _para(text: str) -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


def test_a_grid_of_one_line_names_no_sides_of_its_cells():
    xml = _bordered_table([[_para("a"), _para("b")], [_para("c"), _para("d")]], _sides())
    assert not any(_cell_borders(xml)), "the color and width of the table say it all"


def test_a_cell_with_borders_of_its_own_names_the_sides_of_every_cell():
    own = '<w:tcPr><w:tcBorders><w:bottom w:val="double" w:sz="12" w:space="0" w:color="CC0000"/></w:tcBorders></w:tcPr>'
    xml = _bordered_table([[own + _para("a"), _para("b")], [_para("c"), _para("d")]], _sides(8, "999999"))
    assert _cell_borders(xml) == [
        "single.8.999999/single.8.999999/double.12.CC0000/single.8.999999",
        "single.8.999999/single.8.999999/single.8.999999/single.8.999999",
        "single.8.999999/single.8.999999/single.8.999999/single.8.999999",
        "single.8.999999/single.8.999999/single.8.999999/single.8.999999",
    ]


def test_the_outer_and_the_inside_borders_of_a_table_meet_in_its_cells():
    borders = (
        '<w:top w:val="single" w:sz="12" w:color="000000"/><w:bottom w:val="single" w:sz="12" w:color="000000"/>'
        '<w:left w:val="single" w:sz="12" w:color="000000"/><w:right w:val="single" w:sz="12" w:color="000000"/>'
        '<w:insideH w:val="dotted" w:sz="4" w:color="FF0000"/><w:insideV w:val="dashed" w:sz="4" w:color="00FF00"/>'
    )
    xml = _bordered_table([[_para("a"), _para("b")]], borders)
    assert _cell_borders(xml) == ["single.12.000000/dashed.4.00FF00/single.12.000000/single.12.000000", "single.12.000000/single.12.000000/single.12.000000/dashed.4.00FF00"]


def test_a_cell_merged_down_takes_its_bottom_from_its_last_row():
    restart = '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
    last = '<w:tcPr><w:vMerge/><w:tcBorders><w:bottom w:val="single" w:sz="24" w:color="0000FF"/></w:tcBorders></w:tcPr>'
    xml = _bordered_table([[restart + _para("a"), _para("b")], [last + _para(""), _para("d")]], _sides())
    assert _cell_borders(xml)[0] == "single.4.000000/single.4.000000/single.24.0000FF/single.4.000000"


# --- cell left indent -------------------------------------------------------
#
# pandoc reads an indented paragraph as a BlockQuote, which the filter unwraps
# in a cell, so the indent rides on the sentinel in points.


def _indented_cell(ind: str) -> str:
    return f"<w:tc><w:p><w:pPr><w:ind {ind}/></w:pPr><w:r><w:t>cell</w:t></w:r></w:p></w:tc>"


@pytest.mark.parametrize(("ind", "expected"), [('w:left="142"', "li=7.1"), ('w:start="200"', "li=10")])
def test_an_indented_cell_is_tagged_with_its_indent(ind, expected):
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(_indented_cell(ind)))}))

    assert expected in _cell_sentinel_text(out)


@pytest.mark.parametrize("ind", ['w:left="0"', 'w:left="-142"', 'w:hanging="142"'])
def test_a_cell_without_a_positive_indent_is_not_tagged(ind):
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(_indented_cell(ind)))}))

    assert "li=" not in _cell_sentinel_text(out)


def test_a_cell_starting_with_a_list_paragraph_is_not_tagged():
    """A list paragraph carries the indent of its level, which LaTeX sets for the list already."""
    numbered = '<w:tc><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p></w:tc>'
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(numbered))}))

    assert "li=" not in _cell_sentinel_text(out)


def _paragraph(text: str, left: int | None) -> str:
    ppr = f'<w:pPr><w:ind w:left="{left}"/></w:pPr>' if left is not None else ""
    return f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>" if text else f"<w:p>{ppr}</w:p>"


@pytest.mark.parametrize(
    ("paragraphs", "expected"),
    [
        ((("first", 142), ("second", None)), None),
        ((("first", 142), ("second", 284)), None),
        ((("first", 142), ("second", 142)), "li=7.1"),
        ((("first", 142), ("", None)), "li=7.1"),
    ],
)
def test_a_cell_is_tagged_only_when_its_paragraphs_share_the_indent(paragraphs, expected):
    """The indent applies to the whole cell: Word indents a paragraph which states no indent less."""
    cell = "<w:tc>" + "".join(_paragraph(text, left) for text, left in paragraphs) + "</w:tc>"
    out = docx_table_pre_process.preprocess(_pack({"word/document.xml": _doc(_table(cell))}))

    text = _cell_sentinel_text(out)
    if expected is None:
        assert "li=" not in text
    else:
        assert expected in text
