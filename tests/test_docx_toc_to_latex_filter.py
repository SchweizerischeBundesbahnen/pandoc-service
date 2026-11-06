"""Integration tests for ``filters/docx/docx_toc_to_latex.lua`` and app/docx_toc_pre_process.py.

The filter tests run pandoc inside the pandoc-service container on markdown
Divs of the synthetic styles app/docx_toc_pre_process.py writes. The
end-to-end test converts a DOCX with a table of contents through the service.
"""

from __future__ import annotations

import io

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pypdf import PdfReader

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_toc_to_latex.lua"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _latex(container, markdown: str) -> str:
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/toc.md << 'EOF'\n{markdown}\nEOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/toc.md"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode()


def test_the_contents_marker_becomes_the_table_of_contents(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, '::: {custom-style="PandocToc__CONTENTS_2"}\nPandocToc__CONTENTS_2\n:::\n')
    assert "\\pdcTableOfContents{2}" in latex
    assert "PandocToc__CONTENTS_2" not in latex[latex.index("\\begin{document}") :]
    assert "\\usepackage{titletoc}" in latex
    assert "\\@starttoc{toc}\\endgroup\\vspace{-\\parskip}" in latex, "the empty line after it stands 5pt below the last entry"


def test_a_depth_out_of_range_is_left_as_it_is(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, '::: {custom-style="PandocToc__CONTENTS_10"}\nx\n:::\n')
    assert "\\pdcTableOfContents" not in latex
    assert "titletoc" not in latex


def test_an_entry_gets_its_page_and_the_leader(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, '::: {custom-style="PandocToc__ENTRY"}\n[Figure 1 Caption](#_Toc1)PDCPAGEREF_Toc1\n:::\n')
    body = latex[latex.index("\\begin{document}") :]
    assert "\\pdcTocEntryBegin" in body
    assert "\\pdcTocPage{_Toc1}" in body
    assert "\\pdcTocEntryEnd" in body
    assert "PDCPAGEREF" not in body


def test_a_document_without_a_table_of_contents_is_left_as_it_is(test_parameters: TestParameters):
    assert "titletoc" not in _latex(test_parameters.container, "Just text.\n")


def _toc_field_paragraph(document, levels: str = "1-3") -> None:
    """A paragraph holding an empty TOC field, as pandoc and docx-exporter write it."""
    paragraph = document.add_paragraph()
    for kind, text in (("begin", None), (None, f' TOC \\o "{levels}" \\h \\z \\u '), ("separate", None), ("end", None)):
        run = paragraph.add_run()
        if kind is None:
            instruction = OxmlElement("w:instrText")
            instruction.set(qn("xml:space"), "preserve")
            instruction.text = text
            run._r.append(instruction)
        else:
            char = OxmlElement("w:fldChar")
            char.set(qn("w:fldCharType"), kind)
            run._r.append(char)


def test_the_pdf_holds_the_table_of_contents_of_the_docx(test_parameters: TestParameters):
    document = Document()
    _toc_field_paragraph(document)
    for number in range(1, 4):
        document.add_heading(f"Chapter number {number}", level=1)
        document.add_paragraph("Text of the chapter.")
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    text = PdfReader(io.BytesIO(response.content)).pages[0].extract_text()
    for number in range(1, 4):
        assert text.count(f"Chapter number {number}") == 2, text


def test_the_table_of_contents_holds_the_nine_levels_of_word(test_parameters: TestParameters):
    document = Document()
    _toc_field_paragraph(document, "1-9")
    for level in range(1, 10):
        document.add_heading(f"Heading of level {level}", level=level)
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    text = "".join(page.extract_text() for page in PdfReader(io.BytesIO(response.content)).pages)
    for level in range(1, 10):
        assert text.count(f"Heading of level {level}") == 2, text
