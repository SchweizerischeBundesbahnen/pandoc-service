"""Integration tests for ``filters/docx/docx_toc_to_latex.lua`` and app/docx_toc_pre_process.py.

The filter tests run pandoc inside the pandoc-service container on markdown
Divs of the synthetic styles app/docx_toc_pre_process.py writes. The
end-to-end test converts a DOCX with a table of contents through the service.
"""

from __future__ import annotations

import ctypes
import io
import re

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
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
    latex = _latex(test_parameters.container, '::: {custom-style="PandocToc__CONTENTS_2"}\nPandocToc__CONTENTS_2\n:::\n\n# Heading\n')
    assert "\\pdcTableOfContents{2}" in latex
    assert "PandocToc__CONTENTS_2" not in latex[latex.index("\\begin{document}") :]
    assert "\\usepackage{titletoc}" in latex
    assert "\\@starttoc{toc}\\endgroup\\vspace{-\\parskip}" in latex, "the empty line after it stands 5pt below the last entry"


def test_a_document_without_a_heading_gets_the_message_of_word(test_parameters: TestParameters):
    """Word writes its message into the paragraph which ends the field: the empty line after it goes."""
    latex = _latex(test_parameters.container, '::: {custom-style="PandocToc__CONTENTS_2"}\nPandocToc__CONTENTS_2\n:::\n\n`\\strut`{=latex}\n\nText\n')
    body = latex[latex.index("\\begin{document}") :]
    assert "\\textbf{No table of contents entries found.}" in body
    assert "\\pdcTableOfContents" not in body
    assert body.replace("\\strut{}", "").count("\\strut") == 0


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
    assert "\\pageref{#1}\\kern0.5bp}" in latex, "the page ends at the right tab of Word, 0.5pt short of the margin"


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


def _numbered_heading_styles(document, texts: dict[int, str]) -> None:
    """Number the heading styles of the levels given through the numbering of the styles, with the text given for each.

    The levels above the deepest one given count in decimal, without a style.
    """
    numbering = document.part.numbering_part.element
    levels = ""
    for level in range(1, max(texts) + 1):
        style = f'<w:pStyle w:val="Heading{level}"/>' if level in texts else ""
        text = texts.get(level, f"%{level}")
        levels += f'<w:lvl w:ilvl="{level - 1}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>{style}<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl>'
    numbering.insert(0, parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="90"><w:multiLevelType w:val="multilevel"/>{levels}</w:abstractNum>'))
    numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))
    for level in texts:
        style = document.styles[f"Heading {level}"].element.get_or_add_pPr()
        style.insert(0, parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{level - 1}"/><w:numId w:val="90"/></w:numPr>'))


def _toc_text_x(test_parameters: TestParameters, texts: dict[int, str], headings: list[tuple[int, str]]) -> float:
    """How far right of the margin the first entry of the table of contents sets the text of its heading."""
    document = Document()
    _toc_field_paragraph(document)
    _numbered_heading_styles(document, texts)
    for level, text in headings:
        document.add_heading(text, level=level)
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    textpage = pdfium.PdfDocument(response.content)[0].get_textpage()
    found = textpage.search(headings[0][1], match_case=True).get_next()
    assert found, headings[0][1]
    x, y = ctypes.c_double(), ctypes.c_double()
    pdfium_c.FPDFText_GetCharOrigin(textpage.raw, found[0], ctypes.byref(x), ctypes.byref(y))
    return x.value - document.sections[0].left_margin.pt


def test_an_entry_of_a_numbered_heading_style_sets_its_text_at_the_tab_of_word(test_parameters: TestParameters):
    """Word sets the number, then a tab 12pt past its width in Aptos 12pt, on a grid of 12pt: 24pt for one digit."""
    assert abs(_toc_text_x(test_parameters, {1: "%1"}, [(1, "Alpha")]) - 24) < 0.3


def test_an_entry_of_a_wide_number_sets_its_text_at_a_later_tab(test_parameters: TestParameters):
    """1.1.1.1.1 is 45.8pt wide in Aptos 12pt: the text goes to 60pt, the first step of the grid past 57.8pt."""
    assert abs(_toc_text_x(test_parameters, {1: "%1.%1.%1.%1.%1"}, [(1, "Alpha")]) - 60) < 0.3


def test_an_entry_of_a_heading_style_without_a_number_has_none(test_parameters: TestParameters):
    """The level above a numbered style counts for its number, but has no number of its own."""
    assert abs(_toc_text_x(test_parameters, {2: "%1.%2"}, [(1, "Alpha"), (2, "Beta")])) < 0.3


def test_the_page_of_an_entry_ends_at_the_tab_of_word(test_parameters: TestParameters):
    """Word sets the page at a right tab 10 twips short of the margin, 0.5pt."""
    document = Document()
    _toc_field_paragraph(document)
    document.add_heading("Alpha", level=1)
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    page = pdfium.PdfDocument(response.content)[0]
    textpage = page.get_textpage()
    text = textpage.get_text_range()
    # The page of the first entry: the digit after its leader of dots, with or without a space before it
    found = re.search(r"\.\s*(1)", text)
    assert found, text
    index = found.start(1)
    right = textpage.get_charbox(index, loose=True)[2]
    section = document.sections[0]
    assert abs(right - (section.page_width.pt - section.right_margin.pt - 0.5)) < 0.1
