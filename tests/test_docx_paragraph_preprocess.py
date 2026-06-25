"""Unit tests for ``app.DocxParagraphPreProcess``.

Each test builds a minimal DOCX zip in memory containing a single
``word/document.xml`` plus the bare ``word/styles.xml`` skeleton, runs the
preprocessor, and inspects the rewritten XML — the same synthetic-fixture
approach as ``test_docx_color_preprocess.py``.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET  # noqa: S405

from app import DocxParagraphPreProcess

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)

EMPTY_STYLES_XML = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'


def _pack(parts: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _unpack(blob: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _doc(*paras_xml: str) -> bytes:
    body = "".join(paras_xml)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + "</w:body></w:document>").encode("utf-8")


def _para(ppr_inner: str, text: str = "x") -> str:
    """A <w:p> with the given <w:pPr> inner XML and a single text run."""
    return f"<w:p><w:pPr>{ppr_inner}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"


def _styles(blob: bytes) -> ET.Element:
    return ET.fromstring(_unpack(blob)["word/styles.xml"])  # noqa: S314


def _body(blob: bytes) -> ET.Element:
    return ET.fromstring(_unpack(blob)["word/document.xml"])  # noqa: S314


def _style_ids(styles_root: ET.Element) -> list[str]:
    return [el.get(f"{{{W_NS}}}styleId") or "" for el in styles_root.findall(f"{{{W_NS}}}style")]


def _pstyle_vals(body_root: ET.Element) -> list[str]:
    return [el.get(f"{{{W_NS}}}val") or "" for el in body_root.iter(f"{{{W_NS}}}pStyle")]


def _first_pppr(body_root: ET.Element) -> ET.Element:
    return body_root.find(f".//{{{W_NS}}}pPr")  # type: ignore[return-value]


# --- pass-through ---------------------------------------------------------


def test_plain_paragraph_returns_input_unchanged():
    """No jc / ind anywhere -> bytes are returned untouched (no rezip)."""
    blob = _pack({"word/document.xml": _doc(_para("")), "word/styles.xml": EMPTY_STYLES_XML})
    assert DocxParagraphPreProcess.preprocess(blob) == blob


def test_paragraph_without_ppr_is_ignored():
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>'.encode()
    blob = _pack({"word/document.xml": doc, "word/styles.xml": EMPTY_STYLES_XML})
    assert DocxParagraphPreProcess.preprocess(blob) == blob


# --- alignment ------------------------------------------------------------


def test_center_alignment_becomes_pstyle_and_jc_stripped():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="center"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    result = DocxParagraphPreProcess.preprocess(blob)
    body = _body(result)
    assert body.find(f".//{{{W_NS}}}jc") is None, "w:jc must be stripped"
    assert _pstyle_vals(body) == ["PandocPara__ALIGN_center"]
    assert "PandocPara__ALIGN_center" in _style_ids(_styles(result))


def test_right_alignment():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="right"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    body = _body(DocxParagraphPreProcess.preprocess(blob))
    assert _pstyle_vals(body) == ["PandocPara__ALIGN_right"]


def test_left_alignment_is_not_encoded():
    """left is the default reading direction; encoding it wrapped every
    (notably table-cell) paragraph in \\raggedright for no benefit, so it is
    deliberately left untouched."""
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="left"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    assert DocxParagraphPreProcess.preprocess(blob) == blob


def test_end_folds_to_right_and_start_is_not_encoded():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="start"/>'), _para('<w:jc w:val="end"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    body = _body(DocxParagraphPreProcess.preprocess(blob))
    # start (= left) is not encoded; only the end (= right) paragraph is tagged.
    assert _pstyle_vals(body) == ["PandocPara__ALIGN_right"]


def test_justified_and_left_without_indent_are_not_encoded():
    """left/start and both/distribute all map to the default; nothing to emit."""
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="both"/>'), _para('<w:jc w:val="distribute"/>'), _para('<w:jc w:val="left"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    assert DocxParagraphPreProcess.preprocess(blob) == blob


def test_paragraphs_inside_table_cells_are_not_rewritten():
    """Pandoc renders table-cell alignment/indent from the table structure
    itself, so cell paragraphs must be left alone — rewriting them is redundant
    and the per-cell LaTeX wrapper changes table row spacing (regression)."""
    cell = "<w:tc>" + _para('<w:jc w:val="center"/>', "in-cell") + "</w:tc>"
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr>' + cell + "</w:tr></w:tbl></w:body></w:document>").encode()
    blob = _pack({"word/document.xml": doc, "word/styles.xml": EMPTY_STYLES_XML})
    # No paragraph rewritten -> bytes returned unchanged.
    assert DocxParagraphPreProcess.preprocess(blob) == blob


def test_body_paragraph_still_rewritten_alongside_table_cells():
    """A centered body paragraph outside any cell is still tagged even when the
    document also contains a (skipped) table cell."""
    cell = "<w:tbl><w:tr><w:tc>" + _para('<w:jc w:val="center"/>', "cell") + "</w:tc></w:tr></w:tbl>"
    body_para = _para('<w:jc w:val="center"/>', "body")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + cell + body_para + "</w:body></w:document>").encode()
    blob = _pack({"word/document.xml": doc, "word/styles.xml": EMPTY_STYLES_XML})
    result = DocxParagraphPreProcess.preprocess(blob)
    # Exactly one pStyle added — for the body paragraph, not the cell paragraph.
    assert _pstyle_vals(_body(result)) == ["PandocPara__ALIGN_center"]


# --- indentation ----------------------------------------------------------


def test_left_indent_becomes_pstyle_and_ind_stripped():
    blob = _pack({"word/document.xml": _doc(_para('<w:ind w:left="600"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    result = DocxParagraphPreProcess.preprocess(blob)
    body = _body(result)
    assert body.find(f".//{{{W_NS}}}ind") is None, "w:ind must be stripped so pandoc doesn't make a BlockQuote"
    assert _pstyle_vals(body) == ["PandocPara__IND_600"]
    assert "PandocPara__IND_600" in _style_ids(_styles(result))


def test_distinct_indent_levels_get_distinct_styles():
    """The bug being fixed: 40px (600) and 80px (1200) must stay distinct,
    not collapse into one BlockQuote level."""
    blob = _pack({"word/document.xml": _doc(_para('<w:ind w:left="600"/>'), _para('<w:ind w:left="1200"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    result = DocxParagraphPreProcess.preprocess(blob)
    assert _pstyle_vals(_body(result)) == ["PandocPara__IND_600", "PandocPara__IND_1200"]
    ids = _style_ids(_styles(result))
    assert "PandocPara__IND_600" in ids and "PandocPara__IND_1200" in ids


def test_zero_and_negative_indent_are_not_encoded():
    blob = _pack({"word/document.xml": _doc(_para('<w:ind w:left="0"/>'), _para('<w:ind w:left="-200"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    assert DocxParagraphPreProcess.preprocess(blob) == blob


def test_non_integer_indent_is_ignored():
    blob = _pack({"word/document.xml": _doc(_para('<w:ind w:left="abc"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    assert DocxParagraphPreProcess.preprocess(blob) == blob


# --- combined -------------------------------------------------------------


def test_alignment_and_indent_share_one_pstyle():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="right"/><w:ind w:left="600"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    result = DocxParagraphPreProcess.preprocess(blob)
    assert _pstyle_vals(_body(result)) == ["PandocPara__ALIGN_right__IND_600"]
    assert "PandocPara__ALIGN_right__IND_600" in _style_ids(_styles(result))


def test_existing_pstyle_is_replaced_and_pstyle_is_first_child():
    blob = _pack({"word/document.xml": _doc(_para('<w:pStyle w:val="BodyText"/><w:jc w:val="center"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    body = _body(DocxParagraphPreProcess.preprocess(blob))
    # The old pStyle is gone, replaced by ours.
    assert _pstyle_vals(body) == ["PandocPara__ALIGN_center"]
    # And <w:pStyle> is the first child of <w:pPr> (OOXML CT_PPr schema order).
    ppr = _first_pppr(body)
    assert list(ppr)[0].tag == f"{{{W_NS}}}pStyle"


# --- dedup / idempotency --------------------------------------------------


def test_duplicate_combination_registers_one_style():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="center"/>', "a"), _para('<w:jc w:val="center"/>', "b")), "word/styles.xml": EMPTY_STYLES_XML})
    result = DocxParagraphPreProcess.preprocess(blob)
    assert _pstyle_vals(_body(result)) == ["PandocPara__ALIGN_center", "PandocPara__ALIGN_center"]
    # styles.xml has exactly one entry for the shared style.
    assert _style_ids(_styles(result)).count("PandocPara__ALIGN_center") == 1


def test_preprocess_is_idempotent():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="center"/><w:ind w:left="600"/>')), "word/styles.xml": EMPTY_STYLES_XML})
    once = DocxParagraphPreProcess.preprocess(blob)
    # Second pass finds no jc/ind left to rewrite -> returns its input unchanged.
    assert DocxParagraphPreProcess.preprocess(once) == once


# --- robustness -----------------------------------------------------------


def test_non_docx_input_is_returned_unchanged():
    assert DocxParagraphPreProcess.preprocess(b"not a zip") == b"not a zip"


def test_docx_without_styles_part_is_skipped():
    blob = _pack({"word/document.xml": _doc(_para('<w:jc w:val="center"/>'))})
    assert DocxParagraphPreProcess.preprocess(blob) == blob
