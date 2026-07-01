"""Unit tests for ``app.DocxMathColorPostProcess.apply_math_colors``.

These verify the *decode* half of the math-color shim: given a ``Document`` whose OMML
carries ``@@PMC:RRGGBB@@`` / ``@@PMCEND@@`` marker runs (as pandoc emits them from the
``\\text{}`` markers ``HtmlMathColorPreProcess`` injects), ``apply_math_colors`` adds
``<w:color>`` to the runs between each marker pair and deletes the markers, mutating the
document in place. The encoder is tested in ``test_html_math_color_preprocess.py`` and the
full round-trip through pandoc in ``test_math_color_integration.py``.
"""

from __future__ import annotations

import io

from docx import Document
from docx.oxml import parse_xml
from lxml import etree

from app.DocxMathColorPostProcess import apply_math_colors

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _run(text: str, *, with_m_rpr: bool = False) -> str:
    rpr = "<m:rPr></m:rPr>" if with_m_rpr else ""
    return f"<m:r>{rpr}<m:t>{text}</m:t></m:r>"


def _doc(omath_body: str) -> Document:
    """A Document whose body holds one <m:oMath> built from the given run fragments."""
    doc = Document()
    body = doc.element.find(f"{{{W_NS}}}body")
    body.append(parse_xml(f'<w:p xmlns:w="{W_NS}" xmlns:m="{M_NS}"><m:oMath>{omath_body}</m:oMath></w:p>'))
    return doc


def _runs(doc: Document) -> list[etree._Element]:
    return list(doc.element.iter(f"{{{M_NS}}}r"))


def _text(run: etree._Element) -> str | None:
    t = run.find(f"{{{M_NS}}}t")
    return None if t is None else t.text


def _color(run: etree._Element) -> str | None:
    color = run.find(f"{{{W_NS}}}rPr/{{{W_NS}}}color")
    return None if color is None else color.get(f"{{{W_NS}}}val")


def test_single_colored_run() -> None:
    doc = _doc(_run("@@PMC:FF0000@@") + _run("x") + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    runs = _runs(doc)
    assert len(runs) == 1  # both markers removed
    assert _text(runs[0]) == "x"
    assert _color(runs[0]) == "FF0000"


def test_multiple_content_runs_all_colored() -> None:
    doc = _doc(_run("@@PMC:00FF00@@") + _run("b") + _run("2") + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    runs = _runs(doc)
    assert [_text(r) for r in runs] == ["b", "2"]
    assert all(_color(r) == "00FF00" for r in runs)


def test_nested_colors_use_innermost() -> None:
    body = _run("@@PMC:FF0000@@") + _run("a") + _run("@@PMC:0000FF@@") + _run("b") + _run("@@PMCEND@@") + _run("c") + _run("@@PMCEND@@")
    doc = _doc(body)
    apply_math_colors(doc)
    colors = {_text(r): _color(r) for r in _runs(doc)}
    assert colors == {"a": "FF0000", "b": "0000FF", "c": "FF0000"}


def test_content_nested_in_fraction_is_colored() -> None:
    # The colored content is a fraction, so its runs live inside <m:num>/<m:den>,
    # between the marker runs in document order.
    frac = "<m:f><m:num>" + _run("a") + "</m:num><m:den>" + _run("b") + "</m:den></m:f>"
    doc = _doc(_run("@@PMC:EA1B2C@@") + frac + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    runs = _runs(doc)
    assert [_text(r) for r in runs] == ["a", "b"]
    assert all(_color(r) == "EA1B2C" for r in runs)


def test_existing_math_run_properties_are_kept() -> None:
    # A content run carrying <m:rPr> must keep it; <w:rPr> is inserted after it.
    doc = _doc(_run("@@PMC:FF0000@@") + _run("x", with_m_rpr=True) + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    (run,) = _runs(doc)
    children = list(run)
    assert children[0].tag == f"{{{M_NS}}}rPr"
    assert children[1].tag == f"{{{W_NS}}}rPr"
    assert _color(run) == "FF0000"


def test_document_without_markers_is_unchanged() -> None:
    doc = _doc(_run("x") + _run("y"))
    before = etree.tostring(doc.element)
    apply_math_colors(doc)
    assert etree.tostring(doc.element) == before


def test_end_marker_without_open_is_ignored() -> None:
    # A stray end marker (empty color stack) is removed without error; following runs
    # stay uncolored.
    doc = _doc(_run("@@PMCEND@@") + _run("x"))
    apply_math_colors(doc)
    runs = _runs(doc)
    assert [_text(r) for r in runs] == ["x"]  # end marker removed
    assert _color(runs[0]) is None


def test_content_run_without_text_element_is_colored() -> None:
    # A math run carrying no <m:t> (text is None) is neither marker; while a color is
    # active it still gets <w:color>, and the walk does not crash on the missing text.
    doc = _doc(_run("@@PMC:FF0000@@") + "<m:r></m:r>" + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    (run,) = _runs(doc)
    assert _text(run) is None
    assert _color(run) == "FF0000"


def test_existing_run_color_is_overwritten() -> None:
    # A content run that already carries <w:rPr> with a <w:color> keeps that single
    # rPr/color pair, with the value replaced by the active marker color.
    existing = '<m:r><w:rPr><w:color w:val="000000"/></w:rPr><m:t>x</m:t></m:r>'
    doc = _doc(_run("@@PMC:FF0000@@") + existing + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    (run,) = _runs(doc)
    assert len(run.findall(f"{{{W_NS}}}rPr")) == 1
    assert len(run.find(f"{{{W_NS}}}rPr").findall(f"{{{W_NS}}}color")) == 1
    assert _color(run) == "FF0000"


def test_color_survives_document_save() -> None:
    # The injected <w:color> must survive a python-docx save/reload round-trip.
    doc = _doc(_run("@@PMC:FF0000@@") + _run("x") + _run("@@PMCEND@@"))
    apply_math_colors(doc)
    buf = io.BytesIO()
    doc.save(buf)
    reloaded = Document(io.BytesIO(buf.getvalue()))
    runs = _runs(reloaded)
    assert len(runs) == 1
    assert _color(runs[0]) == "FF0000"
    assert b"@@PMC" not in etree.tostring(reloaded.element)
