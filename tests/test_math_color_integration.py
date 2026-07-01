"""End-to-end integration tests for the math-color shim.

Convert HTML containing colored LaTeX math (``\\color`` / ``\\textcolor``) to DOCX
through the pandoc-service container and assert the resulting OMML carries real
per-run color (``<w:color>``) and no leftover ``@@PMC`` markers.

Why an integration test (vs. the mocked encode/decode unit tests): the shim only
works if the ``\\text{}`` markers survive pandoc's ``texmath`` conversion as distinct
OMML runs, with the colored content as separate runs between them. That contract
lives in ``texmath`` and could shift with a pandoc bump; only a real round-trip
catches such a regression. See ``app/HtmlMathColorPreProcess.py`` and
``app/DocxMathColorPostProcess.py``.
"""

import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

from tests.test_container import TestParameters

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _convert_html_to_docx(test_parameters: TestParameters, html: str) -> bytes:
    url = f"{test_parameters.base_url}/convert/html/to/docx"
    response = test_parameters.request_session.post(url, data=html)
    if response.status_code // 100 != 2:
        raise AssertionError(f"pandoc-service returned {response.status_code}:\n{response.text}")
    return response.content


def _document_xml(docx_bytes: bytes) -> ET.Element:
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def _math_run_colors(tree: ET.Element) -> dict[str, str]:
    """Map each math run's text to its <w:color> value (runs without color omitted)."""
    colors: dict[str, str] = {}
    for run in tree.iter(f"{{{M_NS}}}r"):
        text_el = run.find(f"{{{M_NS}}}t")
        color_el = run.find(f"{{{W_NS}}}rPr/{{{W_NS}}}color")
        if text_el is not None and text_el.text and color_el is not None:
            colors[text_el.text] = color_el.get(f"{{{W_NS}}}val") or ""
    return colors


def _math_html(latex: str) -> str:
    return f'<html><body><p><script type="math/tex; mode=display">{latex}</script></p></body></html>'


def test_textcolor_produces_omml_color(test_parameters: TestParameters):
    docx = _convert_html_to_docx(test_parameters, _math_html("\\textcolor{red}{x}"))
    tree = _document_xml(docx)
    # A real Word equation is produced (texmath would otherwise leak \textcolor as text).
    assert tree.find(f".//{{{M_NS}}}oMath") is not None
    assert _math_run_colors(tree).get("x") == "FF0000"


def test_color_content_is_colored(test_parameters: TestParameters):
    docx = _convert_html_to_docx(test_parameters, _math_html("\\color{Red}{b}"))
    assert _math_run_colors(_document_xml(docx)).get("b") == "FF0000"


def test_markers_do_not_leak_into_output(test_parameters: TestParameters):
    docx = _convert_html_to_docx(test_parameters, _math_html("\\textcolor{red}{x}"))
    with zipfile.ZipFile(BytesIO(docx)) as zf:
        assert b"@@PMC" not in zf.read("word/document.xml")


def test_nested_colors_round_trip(test_parameters: TestParameters):
    docx = _convert_html_to_docx(test_parameters, _math_html("\\color{red}{a\\color{blue}{b}c}"))
    colors = _math_run_colors(_document_xml(docx))
    assert colors.get("a") == "FF0000"
    assert colors.get("b") == "0000FF"
    assert colors.get("c") == "FF0000"
