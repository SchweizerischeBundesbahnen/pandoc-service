"""End-to-end integration tests for the caption pipeline with a localized
caption sequence.

Polarion caption sequences are arbitrary names — a German instance emits
``<span data-sequence="Tabelle" ...>`` and the paragraph label reads
"Tabelle 1 ...", which does NOT start with "Table". The old text-prefix
heuristic silently produced an empty Table of Tables for such documents.
The current pipeline (filters/html_captions.lua marks captions by the
Polarion caption span; DocxReferencesPostProcess keys on the Caption style)
must handle them like any other caption.

These tests run the whole chain through the service container and assert on
the final document.xml.
"""

import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

from tests.test_container import TestParameters

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _convert_html_to_docx(test_parameters: TestParameters, html: str) -> bytes:
    url = f"{test_parameters.base_url}/convert/html/to/docx"
    response = test_parameters.request_session.post(url, data=html)
    if response.status_code // 100 != 2:
        raise AssertionError(f"pandoc-service returned {response.status_code}:\n{response.text}")
    return response.content


def _document_xml(test_parameters: TestParameters, html: str) -> ET.Element:
    docx_bytes = _convert_html_to_docx(test_parameters, html)
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        return ET.fromstring(zf.read("word/document.xml"))


def _instr_texts(doc: ET.Element) -> list[str]:
    return [i.text or "" for i in doc.iter(f"{{{W_NS}}}instrText")]


def _paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t"))


def _paragraphs_with_text(doc: ET.Element, needle: str) -> list[ET.Element]:
    found = [p for p in doc.iter(f"{{{W_NS}}}p") if needle in _paragraph_text(p)]
    if not found:
        raise AssertionError(f"no <w:p> contained {needle!r}")
    return found


def _style_of(p: ET.Element) -> str | None:
    p_style = p.find(f".//{{{W_NS}}}pStyle")
    return p_style.get(f"{{{W_NS}}}val") if p_style is not None else None


_TABELLE_DOC = (
    "<p>TOT_PLACEHOLDER</p>"
    "<p>Table 3 shows the false-positive scenario.</p>"
    '<p class="polarion-rte-caption-paragraph" style="text-align: left;">'
    'Tabelle <span data-sequence="Tabelle" class="polarion-rte-caption">1</span> -- erste Tabelle</p>'
    "<table><tr><td>A</td><td>B</td></tr></table>"
    '<p class="polarion-rte-caption-paragraph" style="text-align: left;">'
    'Tabelle <span data-sequence="Tabelle" class="polarion-rte-caption">2</span> -- zweite Tabelle</p>'
    "<table><tr><td>C</td><td>D</td></tr></table>"
)


def test_localized_sequence_captions_end_up_in_table_of_tables(test_parameters: TestParameters):
    """German "Tabelle" captions get SEQ fields (keeping their own sequence
    identifier) and are listed in the pre-filled Table of Tables."""
    doc = _document_xml(test_parameters, _TABELLE_DOC)
    instr = _instr_texts(doc)

    # Caption numbers became SEQ fields with the Polarion sequence identifier
    assert sum("SEQ Tabelle" in t for t in instr) == 2, f"expected 2 SEQ Tabelle fields, instructions: {instr!r}"

    # The ToT placeholder became a TOC \f T field with pre-filled entries
    assert any("TOC \\h \\z \\f T" in t for t in instr), "TOT placeholder was not replaced with a TOC field"
    assert sum("PAGEREF" in t for t in instr) == 2, "pre-filled ToT entries (PAGEREF hyperlinks) missing"
    assert "TOT_PLACEHOLDER" not in "".join(_paragraph_text(p) for p in doc.iter(f"{{{W_NS}}}p"))

    # Caption paragraphs carry the Caption style. The caption text also
    # appears in the pre-filled ToT list (styled TOC1), so look for the
    # Caption-styled occurrence among all matches.
    for needle in ("erste Tabelle", "zweite Tabelle"):
        styles = [_style_of(p) for p in _paragraphs_with_text(doc, needle)]
        assert "Caption" in styles, f"caption {needle!r} did not get the Caption style (styles found: {styles!r})"


def test_localized_sequence_does_not_enable_text_heuristic(test_parameters: TestParameters):
    """A body paragraph starting with "Table" must NOT be restyled as a
    caption or listed in the Table of Tables."""
    doc = _document_xml(test_parameters, _TABELLE_DOC)

    styles = [_style_of(p) for p in _paragraphs_with_text(doc, "false-positive scenario")]
    assert "Caption" not in styles, "body text starting with 'Table' was restyled as a caption"
    assert not any("false-positive" in t for t in _instr_texts(doc)), "body text leaked into TC entries"
