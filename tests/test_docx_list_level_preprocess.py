"""Unit tests for ``app.DocxListLevelPreProcess``.

Each test builds a minimal DOCX zip in memory containing a single
``word/document.xml`` and inspects the rewritten XML — same synthetic-fixture
approach as ``test_docx_color_preprocess.py``.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET  # noqa: S405

from app import DocxListLevelPreProcess
from app.DocxListLevelPreProcess import SENTINEL_CLOSE, SENTINEL_OPEN

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)


def _pack(parts: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _doc(*paras_xml: str) -> bytes:
    body = "".join(paras_xml)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + "</w:body></w:document>").encode("utf-8")


def _list_para(ilvl: str | None, text: str = "item") -> str:
    """A list <w:p> with a <w:numPr> at the given ilvl (None = numPr without ilvl)."""
    ilvl_xml = f'<w:ilvl w:val="{ilvl}"/>' if ilvl is not None else ""
    return f'<w:p><w:pPr><w:numPr>{ilvl_xml}<w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


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


def _sentinel(level: int) -> str:
    return f"{SENTINEL_OPEN}{level}{SENTINEL_CLOSE}"


# --- tagging --------------------------------------------------------------


def test_list_paragraph_gets_sentinel_for_its_ilvl():
    blob = _pack({"word/document.xml": _doc(_list_para("2", "Level 3"))})
    result = DocxListLevelPreProcess.preprocess(blob)
    para = _body(result).find(f".//{{{W_NS}}}p")
    # The first run is the sentinel encoding ilvl=2; the original text follows.
    assert _first_run_text(para) == _sentinel(2)
    texts = [t.text for t in para.iter(f"{{{W_NS}}}t")]
    assert texts == [_sentinel(2), "Level 3"]


def test_numpr_without_ilvl_defaults_to_level_zero():
    blob = _pack({"word/document.xml": _doc(_list_para(None, "L1"))})
    para = _body(DocxListLevelPreProcess.preprocess(blob)).find(f".//{{{W_NS}}}p")
    assert _first_run_text(para) == _sentinel(0)


def test_sentinel_is_inserted_after_ppr():
    blob = _pack({"word/document.xml": _doc(_list_para("1"))})
    para = _body(DocxListLevelPreProcess.preprocess(blob)).find(f".//{{{W_NS}}}p")
    children = list(para)
    assert children[0].tag == f"{{{W_NS}}}pPr"
    assert children[1].tag == f"{{{W_NS}}}r"  # sentinel run right after pPr


def test_multiple_levels_tagged_independently():
    blob = _pack({"word/document.xml": _doc(_list_para("0", "a"), _list_para("2", "b"), _list_para("1", "c"))})
    paras = _body(DocxListLevelPreProcess.preprocess(blob)).findall(f".//{{{W_NS}}}p")
    assert [_first_run_text(p) for p in paras] == [_sentinel(0), _sentinel(2), _sentinel(1)]


# --- pass-through ---------------------------------------------------------


def test_non_list_paragraph_is_untouched():
    blob = _pack({"word/document.xml": _doc(_plain_para("hello"))})
    assert DocxListLevelPreProcess.preprocess(blob) == blob


def test_document_without_lists_returned_unchanged():
    blob = _pack({"word/document.xml": _doc(_plain_para("a"), _plain_para("b"))})
    assert DocxListLevelPreProcess.preprocess(blob) == blob


def test_mixed_doc_only_tags_list_paragraphs():
    blob = _pack({"word/document.xml": _doc(_plain_para("intro"), _list_para("0", "item"))})
    paras = _body(DocxListLevelPreProcess.preprocess(blob)).findall(f".//{{{W_NS}}}p")
    assert _first_run_text(paras[0]) == "intro"  # plain paragraph untouched
    assert _first_run_text(paras[1]) == _sentinel(0)


# --- robustness -----------------------------------------------------------


def test_preprocess_is_idempotent():
    blob = _pack({"word/document.xml": _doc(_list_para("2", "x"))})
    once = DocxListLevelPreProcess.preprocess(blob)
    # A second pass must not prepend a second sentinel.
    assert DocxListLevelPreProcess.preprocess(once) == once
    para = _body(once).find(f".//{{{W_NS}}}p")
    assert [t.text for t in para.iter(f"{{{W_NS}}}t")] == [_sentinel(2), "x"]


def test_non_docx_input_is_returned_unchanged():
    assert DocxListLevelPreProcess.preprocess(b"not a zip") == b"not a zip"


def test_zip_without_body_parts_returned_unchanged():
    """A package with no document/header/footer parts has nothing to tag."""
    blob = _pack({"word/styles.xml": b"<x/>", "docProps/core.xml": b"<x/>"})
    assert DocxListLevelPreProcess.preprocess(blob) == blob


def test_ilvl_element_without_val_defaults_to_zero():
    para = '<w:p><w:pPr><w:numPr><w:ilvl/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    para_el = _body(DocxListLevelPreProcess.preprocess(_pack({"word/document.xml": _doc(para)}))).find(f".//{{{W_NS}}}p")
    assert _first_run_text(para_el) == _sentinel(0)


def test_non_integer_ilvl_is_not_tagged():
    blob = _pack({"word/document.xml": _doc(_list_para("abc", "x"))})
    assert DocxListLevelPreProcess.preprocess(blob) == blob


def test_negative_ilvl_is_not_tagged():
    blob = _pack({"word/document.xml": _doc(_list_para("-1", "x"))})
    assert DocxListLevelPreProcess.preprocess(blob) == blob


def test_malformed_document_xml_returned_unchanged():
    blob = _pack({"word/document.xml": b"<w:document><unclosed>"})
    assert DocxListLevelPreProcess.preprocess(blob) == blob


def test_list_paragraph_with_leading_bookmark_is_still_tagged():
    """A list paragraph whose first child after pPr is not a run (e.g. a
    bookmark) is still tagged — the sentinel goes right after pPr."""
    para = '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>x</w:t></w:r></w:p>'
    para_el = _body(DocxListLevelPreProcess.preprocess(_pack({"word/document.xml": _doc(para)}))).find(f".//{{{W_NS}}}p")
    children = list(para_el)
    assert children[0].tag == f"{{{W_NS}}}pPr"
    assert children[1].tag == f"{{{W_NS}}}r"
    assert _first_run_text(para_el) == _sentinel(1)
