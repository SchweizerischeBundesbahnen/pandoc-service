"""Unit tests for ``app.DocxMathColorPreProcess``.

These verify the *encode* half of the DOCX -> LaTeX/PDF math-color shim: given a DOCX
whose OMML carries direct ``<w:color>`` on math runs (as ``DocxMathColorPostProcess``
writes on the HTML -> DOCX path, and as Word renders), the preprocessor wraps each
colored run's ``<m:t>`` text in ``PMCzzzRRGGBBzzz`` / ``zzzPMCENDzzz`` markers and strips
the ``<w:color>``. ``filters/docx_math_colors_to_latex.lua`` turns those markers back into
``\\color`` and is exercised, with pandoc, in the integration test.

Each test builds a minimal DOCX zip in memory containing a single ``word/document.xml``
and inspects the rewritten OMML.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET  # noqa: S405

from app import DocxMathColorPreProcess

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _pack(document_xml: bytes) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)
    return buf.getvalue()


def _unpack(blob: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        return zf.read("word/document.xml")


def _doc(omath_body: str) -> bytes:
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:body><w:p><m:oMath>{omath_body}</m:oMath></w:p></w:body></w:document>').encode()


def _run(text: str, *, color: str | None = None, m_rpr: bool = False) -> str:
    w_rpr = f'<w:rPr><w:color w:val="{color}"/></w:rPr>' if color else ""
    m_rpr_xml = '<m:rPr><m:sty m:val="p"/></m:rPr>' if m_rpr else ""
    return f"<m:r>{m_rpr_xml}{w_rpr}<m:t>{text}</m:t></m:r>"


def _texts(blob: bytes) -> list[str | None]:
    root = ET.fromstring(_unpack(blob))  # noqa: S314
    return [t.text for t in root.iter(f"{{{M_NS}}}t")]


def _colors(blob: bytes) -> list[str]:
    root = ET.fromstring(_unpack(blob))  # noqa: S314
    return [c.get(f"{{{W_NS}}}val") or "" for c in root.iter(f"{{{W_NS}}}color")]


def test_colored_run_is_wrapped_and_color_stripped() -> None:
    out = DocxMathColorPreProcess.preprocess(_pack(_doc(_run("E", color="FF0000"))))
    assert _texts(out) == ["PMCzzzFF0000zzzEzzzPMCENDzzz"]
    assert _colors(out) == []  # <w:color> removed once encoded


def test_lowercase_and_hash_hex_normalized_to_uppercase() -> None:
    out = DocxMathColorPreProcess.preprocess(_pack(_doc(_run("x", color="#00ffcc"))))
    assert _texts(out) == ["PMCzzz00FFCCzzzxzzzPMCENDzzz"]


def test_multiple_runs_each_wrapped_with_own_color() -> None:
    body = _run("E", color="FF0000") + _run("=", m_rpr=True) + _run("m", color="0000FF")
    out = DocxMathColorPreProcess.preprocess(_pack(_doc(body)))
    assert _texts(out) == ["PMCzzzFF0000zzzEzzzPMCENDzzz", "=", "PMCzzz0000FFzzzmzzzPMCENDzzz"]


def test_run_without_color_is_untouched() -> None:
    blob = _pack(_doc(_run("x") + _run("y")))
    out = DocxMathColorPreProcess.preprocess(blob)
    assert out == blob  # no colored math -> original bytes returned unchanged


def test_auto_and_theme_colors_are_skipped() -> None:
    body = _run("x", color="auto") + _run("y")
    blob = _pack(_doc(body))
    out = DocxMathColorPreProcess.preprocess(blob)
    assert out == blob  # neither run gets markers


def test_non_hex_color_is_skipped() -> None:
    blob = _pack(_doc(_run("x", color="notacolor")))
    out = DocxMathColorPreProcess.preprocess(blob)
    assert out == blob


def test_colored_run_without_text_element_does_not_crash() -> None:
    # A math run with <w:color> but no <m:t> is left as-is (nothing to wrap).
    body = '<m:r><w:rPr><w:color w:val="FF0000"/></w:rPr></m:r>'
    blob = _pack(_doc(body))
    out = DocxMathColorPreProcess.preprocess(blob)
    assert out == blob


def test_math_run_with_rpr_but_no_color_is_skipped() -> None:
    # <w:rPr> present (e.g. bold) but no <w:color> child -> run left untouched.
    body = "<m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r>"
    blob = _pack(_doc(body))
    assert DocxMathColorPreProcess.preprocess(blob) == blob


def test_color_element_without_value_is_skipped() -> None:
    # <w:color/> with no w:val attribute -> nothing to encode, run left untouched.
    body = "<m:r><w:rPr><w:color/></w:rPr><m:t>x</m:t></m:r>"
    blob = _pack(_doc(body))
    assert DocxMathColorPreProcess.preprocess(blob) == blob


def test_unparseable_document_part_is_skipped() -> None:
    # A valid zip whose word/document.xml is malformed XML -> the part is skipped
    # (parse returns None) and the package is returned unchanged.
    blob = _pack(b"<not-well-formed")
    assert DocxMathColorPreProcess.preprocess(blob) == blob


def test_invalid_zip_returned_unchanged() -> None:
    assert DocxMathColorPreProcess.preprocess(b"not a zip") == b"not a zip"
