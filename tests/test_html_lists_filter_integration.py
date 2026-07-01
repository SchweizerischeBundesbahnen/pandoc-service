"""End-to-end integration tests for the html-lists pipeline.

These exercise the preprocessor (``app.HtmlListsPreProcess``) plus the Lua
filter (``filters/html_lists.lua``) together by converting HTML → DOCX through
the pandoc-service container (which applies both the preprocessor and the Lua
filter automatically for HTML→DOCX conversions), then inspecting the resulting
``word/document.xml`` to confirm the stray marker paragraph is gone.

The unit test for the preprocessor lives in
``test_html_lists_preprocess.py`` — this file is specifically about the
contract between the two halves: the sentinel <span> placed by the
preprocessor must reach pandoc as a Span with the right class, and the Lua
filter must turn the corresponding ListItem into a no-marker first block.
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET  # noqa: S405

import pytest

from tests.test_container import TestParameters

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/html_lists.lua"


def _convert_html_to_docx(test_parameters: TestParameters, html: str) -> bytes:
    """Convert HTML to DOCX via the pandoc-service container API.

    The service automatically applies the HtmlListsPreProcess preprocessor
    and the html_lists.lua filter for HTML→DOCX conversions.

    Returns the raw DOCX bytes.
    """
    url = f"{test_parameters.base_url}/convert/html/to/docx"
    response = test_parameters.request_session.post(url, data=html)
    if response.status_code // 100 != 2:
        raise AssertionError(f"pandoc-service returned {response.status_code}:\n{response.text}")
    return response.content


def _list_paragraphs(doc_xml: bytes) -> list[tuple[str, str, str]]:
    """Return [(ilvl, numId, text), ...] for every <w:p> in document.xml that
    carries a <w:numPr>. Paragraphs without numPr (regular body text) are
    omitted — we only care about the numbered/bulleted ones for these tests.
    """
    doc = ET.fromstring(doc_xml)
    out: list[tuple[str, str, str]] = []
    for p in doc.iter(f"{{{W_NS}}}p"):
        numPr = p.find(f".//{{{W_NS}}}numPr")
        if numPr is None:
            continue
        ilvl_el = numPr.find(f"{{{W_NS}}}ilvl")
        numId_el = numPr.find(f"{{{W_NS}}}numId")
        ilvl = ilvl_el.get(f"{{{W_NS}}}val") if ilvl_el is not None else ""
        numId = numId_el.get(f"{{{W_NS}}}val") if numId_el is not None else ""
        text = "".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t"))
        out.append((ilvl, numId, text))
    return out


def test_polarion_orphan_ol_does_not_emit_stray_marker(test_parameters: TestParameters):
    """The bug we set out to fix: the Polarion-shaped input must yield
    exactly three list paragraphs (Level 1 / Level 3 / Level 2), with no
    empty marker paragraph between Level 1 and Level 3.

    Before the fix, pandoc emitted four list paragraphs: the canonical
    three plus an empty <w:p> at ilvl=1 that Word renders as a stray "a."
    marker.
    """
    html = '<ol id="polarion_1"><li>Level 1<ol><ol><li>Level 3</li></ol><li>Level 2</li></ol></li></ol>'
    docx_bytes = _convert_html_to_docx(test_parameters, html)
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        doc_xml = zf.read("word/document.xml")

    paragraphs = _list_paragraphs(doc_xml)
    assert len(paragraphs) == 3, f"expected exactly 3 numbered paragraphs (Level 1 / Level 3 / Level 2); got {len(paragraphs)}: {paragraphs!r} — a stray empty marker paragraph was emitted, the html_lists.lua filter did not strip it"
    # And every list paragraph has real text content — no empties.
    for ilvl, numId, text in paragraphs:
        assert text.strip(), f"empty list paragraph found at ilvl={ilvl} numId={numId} — filter didn't suppress the marker"

    # Sanity: depths are 1 / 3 / 2 (zero-indexed: 0 / 2 / 1).
    ilvls = [ilvl for ilvl, _, _ in paragraphs]
    assert ilvls == ["0", "2", "1"], f"expected ilvls 0/2/1, got {ilvls}"

    # Sanity: text content is in the right order.
    texts = [text for _, _, text in paragraphs]
    assert texts == ["Level 1", "Level 3", "Level 2"], f"texts: {texts}"


def test_well_formed_nested_list_keeps_all_markers(test_parameters: TestParameters):
    """Regression guard: a normal nested list with real <li> wrappers must
    NOT lose markers — the filter only acts on the sentinel left by the
    preprocessor."""
    html = "<ol><li>One<ol><li>One.A</li><li>One.B</li></ol></li><li>Two</li></ol>"
    docx_bytes = _convert_html_to_docx(test_parameters, html)
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        doc_xml = zf.read("word/document.xml")

    paragraphs = _list_paragraphs(doc_xml)
    texts = [text for _, _, text in paragraphs]
    assert texts == ["One", "One.A", "One.B", "Two"], f"texts: {texts}"
    # Every item has content (no empty marker paragraphs).
    for ilvl, _, text in paragraphs:
        assert text.strip(), f"empty list paragraph found at ilvl={ilvl} — filter erroneously suppressed a real marker"


def test_orphan_ul_pattern_also_handled(test_parameters: TestParameters):
    """The same pattern with <ul>/<li> instead of <ol> must also work."""
    html = "<ul><li>One<ul><ul><li>Deep</li></ul><li>Two</li></ul></li></ul>"
    docx_bytes = _convert_html_to_docx(test_parameters, html)
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        doc_xml = zf.read("word/document.xml")

    paragraphs = _list_paragraphs(doc_xml)
    texts = [text for _, _, text in paragraphs]
    assert texts == ["One", "Deep", "Two"], f"texts: {texts}"
    for ilvl, _, text in paragraphs:
        assert text.strip(), f"empty bullet paragraph found at ilvl={ilvl}"


def test_sentinel_span_without_orphan_pattern_is_a_no_op(test_parameters: TestParameters):
    """If the sentinel <span> shows up inside a normal list item (e.g.
    a hand-authored input that accidentally collides with our class name),
    the filter still does its job — strips the marker — but at least the
    user's text survives. Document the actual behavior so future edits
    don't break the assumption silently."""
    html = '<ol><li><span class="pandoc-suppress-marker"></span><ol><li>nested</li></ol></li><li>after</li></ol>'
    # Note: we bypass the preprocessor here — we WANT to test the filter
    # in isolation, against a hand-authored sentinel. Use docker exec to
    # run pandoc directly inside the container with only this filter.
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(
        ["sh", "-c", f"cat > /tmp/test/in.html << 'HEREDOC_EOF'\n{html}\nHEREDOC_EOF"],
    )
    exit_code, stderr = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f html -t docx --lua-filter={FILTER_PATH} -o /tmp/test/out.docx /tmp/test/in.html"],
    )
    assert exit_code == 0, f"pandoc failed: {stderr.decode()}"

    # Read the output file
    exit_code, docx_bytes = container.exec_run(["cat", "/tmp/test/out.docx"])
    assert exit_code == 0

    with zipfile.ZipFile(BytesIO(docx_bytes)) as zf:
        doc_xml = zf.read("word/document.xml")

    paragraphs = _list_paragraphs(doc_xml)
    texts = [text for _, _, text in paragraphs]
    # The item containing the sentinel + nested list has no own text; only
    # "nested" and "after" should produce visible list paragraphs.
    assert "nested" in texts and "after" in texts

    # Cleanup
    container.exec_run(["rm", "-rf", "/tmp/test"])
