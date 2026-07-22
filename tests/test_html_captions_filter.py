"""Integration tests for ``filters/html_captions.lua``.

Runs the real ``pandoc`` binary (html -> docx) with the filter and inspects the
paragraph styles in the produced ``document.xml``. The filter marks a paragraph
with the "Caption" style iff it carries Polarion's caption counter span
(``<span data-sequence=... class="polarion-rte-caption">``); everything else —
headings, cross-references, size labels — must be left unstyled so
``DocxReferencesPostProcess`` never mistakes them for captions.
"""

from __future__ import annotations

import io
import re
import shutil
import subprocess
import zipfile

import pytest

_PANDOC = shutil.which("pandoc")
pytestmark = pytest.mark.skipif(_PANDOC is None, reason="pandoc binary not available")

_FILTER = "filters/html_captions.lua"


def _styles_by_text(body: str) -> dict[str, str | None]:
    html = f"<html><head><title>t</title></head><body>{body}</body></html>"
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "html", "-t", "docx", "--lua-filter", _FILTER, "-o", "-"],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    document_xml = zipfile.ZipFile(io.BytesIO(completed.stdout)).read("word/document.xml").decode()
    result: dict[str, str | None] = {}
    for match in re.finditer(r"<w:p\b.*?</w:p>", document_xml, re.S):
        para = match.group(0)
        style = re.search(r'<w:pStyle w:val="([^"]+)"', para)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para)).strip()
        if text:
            result[text] = style.group(1) if style else None
    return result


_CAPTION_P = '<p class="polarion-rte-caption-paragraph">Table <span data-sequence="Table" class="polarion-rte-caption">1</span> Real caption</p>'
_FIGURE_CAPTION_P = '<p class="polarion-rte-caption-paragraph">Figure <span data-sequence="Figure" class="polarion-rte-caption">2</span> A figure</p>'


def test_real_caption_gets_caption_style():
    styles = _styles_by_text(_CAPTION_P)
    assert styles["Table 1 Real caption"] == "Caption"


def test_figure_caption_gets_caption_style():
    styles = _styles_by_text(_FIGURE_CAPTION_P)
    assert styles["Figure 2 A figure"] == "Caption"


def test_prose_starting_with_table_and_number_is_not_captioned():
    """The exact false positive the text heuristic produced: a cross-reference."""
    styles = _styles_by_text("<p>Table 1 shows the results discussed above.</p>")
    assert styles["Table 1 shows the results discussed above."] != "Caption"


def test_size_label_is_not_captioned():
    styles = _styles_by_text("<p>Table 50px</p>")
    assert styles["Table 50px"] != "Caption"


def test_heading_div_title_is_not_captioned():
    styles = _styles_by_text('<div class="title">Table test III</div>')
    assert styles["Table test III"] != "Caption"


def test_only_the_caption_is_marked_among_lookalikes():
    body = _CAPTION_P + "<p>Table 1 is described below.</p>" + '<div class="title">Table overview</div>'
    styles = _styles_by_text(body)
    assert styles["Table 1 Real caption"] == "Caption"
    assert styles["Table 1 is described below."] != "Caption"
    assert styles["Table overview"] != "Caption"


def _document_xml(body: str) -> str:
    """Run pandoc with the caption filter and return raw document.xml."""
    html = f"<html><head><title>t</title></head><body>{body}</body></html>"
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "html", "-t", "docx", "--lua-filter", _FILTER, "-o", "-"],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    return zipfile.ZipFile(io.BytesIO(completed.stdout)).read("word/document.xml").decode()


def test_table_caption_contains_seq_field():
    """Caption number must be wrapped in a SEQ Table field, not plain text."""
    xml = _document_xml(_CAPTION_P)
    assert 'SEQ Table' in xml
    assert 'w:fldChar' in xml


def test_figure_caption_contains_seq_field():
    """Figure caption number must use SEQ Figure field."""
    xml = _document_xml(_FIGURE_CAPTION_P)
    assert 'SEQ Figure' in xml
    assert 'w:fldChar' in xml


def test_seq_field_has_cached_number():
    """The SEQ field should contain the original number as cached display value."""
    xml = _document_xml(_CAPTION_P)
    # Between fldChar separate and end there should be a <w:t>1</w:t>
    assert re.search(r'fldCharType="separate".*?<w:t[^>]*>1</w:t>.*?fldCharType="end"', xml, re.S)


def test_non_docx_target_is_untouched():
    """The filter only acts for the docx writer (defensive FORMAT gate)."""
    html = f"<html><body>{_CAPTION_P}</body></html>"
    completed = subprocess.run(  # noqa: S603
        [_PANDOC, "-f", "html", "-t", "latex", "--lua-filter", _FILTER],
        input=html.encode(),
        capture_output=True,
        check=True,
    )
    # No custom-style Div wrapping leaks into LaTeX output.
    assert "Caption" not in completed.stdout.decode()
