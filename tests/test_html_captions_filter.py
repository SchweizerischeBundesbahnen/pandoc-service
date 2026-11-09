"""Integration tests for ``filters/html/html_captions.lua``.

Runs the ``pandoc`` of the image (html -> docx) with the filter and inspects the
paragraph styles in the produced ``document.xml``. The filter marks a paragraph
with the "Caption" style iff it carries Polarion's caption counter span
(``<span data-sequence=... class="polarion-rte-caption">``); everything else —
headings, cross-references, size labels — must be left unstyled so
``docx_references_post_process`` never mistakes them for captions.
"""

from __future__ import annotations

import io
import re
import zipfile
from typing import TYPE_CHECKING

from tests.pandoc_in_container import filter_path, run_pandoc

if TYPE_CHECKING:
    from docker.models.containers import Container

    from tests.test_container import TestParameters

_FILTER = filter_path("html/html_captions.lua")


def _styles_by_text(container: Container, body: str) -> dict[str, str | None]:
    document_xml = _document_xml(container, body)
    result: dict[str, str | None] = {}
    for match in re.finditer(r"<w:p\b.*?</w:p>", document_xml, re.DOTALL):
        para = match.group(0)
        style = re.search(r'<w:pStyle w:val="([^"]+)"', para)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para)).strip()
        if text:
            result[text] = style.group(1) if style else None
    return result


_CAPTION_P = '<p class="polarion-rte-caption-paragraph">Table <span data-sequence="Table" class="polarion-rte-caption">1</span> Real caption</p>'
_FIGURE_CAPTION_P = '<p class="polarion-rte-caption-paragraph">Figure <span data-sequence="Figure" class="polarion-rte-caption">2</span> A figure</p>'


def test_real_caption_gets_caption_style(test_parameters: TestParameters):
    styles = _styles_by_text(test_parameters.container, _CAPTION_P)
    assert styles["Table 1 Real caption"] == "Caption"


def test_figure_caption_gets_caption_style(test_parameters: TestParameters):
    styles = _styles_by_text(test_parameters.container, _FIGURE_CAPTION_P)
    assert styles["Figure 2 A figure"] == "Caption"


def test_prose_starting_with_table_and_number_is_not_captioned(test_parameters: TestParameters):
    """The exact false positive the text heuristic produced: a cross-reference."""
    styles = _styles_by_text(test_parameters.container, "<p>Table 1 shows the results discussed above.</p>")
    assert styles["Table 1 shows the results discussed above."] != "Caption"


def test_size_label_is_not_captioned(test_parameters: TestParameters):
    styles = _styles_by_text(test_parameters.container, "<p>Table 50px</p>")
    assert styles["Table 50px"] != "Caption"


def test_heading_div_title_is_not_captioned(test_parameters: TestParameters):
    styles = _styles_by_text(test_parameters.container, '<div class="title">Table test III</div>')
    assert styles["Table test III"] != "Caption"


def test_only_the_caption_is_marked_among_lookalikes(test_parameters: TestParameters):
    body = _CAPTION_P + "<p>Table 1 is described below.</p>" + '<div class="title">Table overview</div>'
    styles = _styles_by_text(test_parameters.container, body)
    assert styles["Table 1 Real caption"] == "Caption"
    assert styles["Table 1 is described below."] != "Caption"
    assert styles["Table overview"] != "Caption"


def _document_xml(container: Container, body: str) -> str:
    """Run pandoc with the caption filter and return raw document.xml."""
    html = f"<html><head><title>t</title></head><body>{body}</body></html>"
    docx = run_pandoc(container, ["-f", "html", "-t", "docx", f"--lua-filter={_FILTER}"], html.encode(), "in.html")
    return zipfile.ZipFile(io.BytesIO(docx)).read("word/document.xml").decode()


def test_table_caption_contains_seq_field(test_parameters: TestParameters):
    """Caption number must be wrapped in a SEQ Table field, not plain text."""
    xml = _document_xml(test_parameters.container, _CAPTION_P)
    assert "SEQ Table" in xml
    assert "w:fldChar" in xml


def test_figure_caption_contains_seq_field(test_parameters: TestParameters):
    """Figure caption number must use SEQ Figure field."""
    xml = _document_xml(test_parameters.container, _FIGURE_CAPTION_P)
    assert "SEQ Figure" in xml
    assert "w:fldChar" in xml


def test_seq_field_has_cached_number(test_parameters: TestParameters):
    """The SEQ field should contain the original number as cached display value."""
    xml = _document_xml(test_parameters.container, _CAPTION_P)
    # Between fldChar separate and end there should be a <w:t>1</w:t>
    assert re.search(r'fldCharType="separate".*?<w:t[^>]*>1</w:t>.*?fldCharType="end"', xml, re.DOTALL)


def test_non_docx_target_is_untouched(test_parameters: TestParameters):
    """The filter only acts for the docx writer (defensive FORMAT gate)."""
    html = f"<html><body>{_CAPTION_P}</body></html>"
    latex = run_pandoc(test_parameters.container, ["-f", "html", "-t", "latex", f"--lua-filter={_FILTER}"], html.encode(), "in.html")
    # No custom-style Div wrapping leaks into LaTeX output.
    assert "Caption" not in latex.decode()
