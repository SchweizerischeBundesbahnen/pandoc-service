"""Integration tests for ``filters/docx_page_breaks_to_latex.lua``.

The docx reader drops page and section breaks. ``app/docx_page_break_pre_process.py``
writes them as sentinels, and the filter turns those into LaTeX. The unit cases feed
pandoc a native AST inside the container; the end-to-end cases send a DOCX through
the service.
"""

from __future__ import annotations

import io

import pypdfium2 as pdfium
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_BREAK
from docx.shared import Cm
from pypdf import PageObject, PdfReader

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_page_breaks_to_latex.lua"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PAGE_BREAK = "\\ue020\\ue021"
LANDSCAPE = "\\ue02016838,11906,1440,1440,1440,1440\\ue021"


def _native_to_latex(container, native: str) -> str:
    """Convert a native AST to LaTeX through the filter, whitespace collapsed."""
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/in.native"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return " ".join(output.decode("utf-8").split())


def _str(text: str) -> str:
    """A pandoc native Str. Haskell escapes the code points as decimals."""
    return 'Str "' + text.replace("\\ue020", "\\57376\\&").replace("\\ue021", "\\57377\\&") + '"'


def test_page_break_becomes_clearpage(test_parameters: TestParameters):
    native = f'[ Para [ Str "before" ], Para [ {_str(PAGE_BREAK)} ], Para [ Str "after" ] ]'
    assert _native_to_latex(test_parameters.container, native) == "before \\clearpage{} after"


def test_page_break_inside_a_div(test_parameters: TestParameters):
    native = f'[ Para [ Str "before" ], Div ( "" , [] , [] ) [ Para [ {_str(PAGE_BREAK)} ] ] ]'
    assert "\\clearpage" in _native_to_latex(test_parameters.container, native)


def test_page_break_in_the_middle_of_a_text(test_parameters: TestParameters):
    native = f"[ Para [ {_str(f'before{PAGE_BREAK}after')} ] ]"
    assert _native_to_latex(test_parameters.container, native) == "before\\clearpage{}after"


def test_second_break_in_a_row_keeps_its_empty_page(test_parameters: TestParameters):
    """LaTeX ignores \\clearpage on an empty page, where Word starts one more page."""
    native = f'[ Para [ Str "before" ], Para [ {_str(PAGE_BREAK)} ], Para [ {_str(PAGE_BREAK)} ], Para [ Str "after" ] ]'
    assert _native_to_latex(test_parameters.container, native) == "before \\clearpage{} \\null\\clearpage{} after"


def test_break_at_the_start_keeps_its_empty_page(test_parameters: TestParameters):
    native = f'[ Para [ {_str(PAGE_BREAK)} ], Para [ Str "after" ] ]'
    assert _native_to_latex(test_parameters.container, native) == "\\null\\clearpage{} after"


def test_line_break_next_to_a_page_break_goes(test_parameters: TestParameters):
    """LaTeX has no line for it to end there."""
    native = f'[ Para [ Str "before" ], Para [ LineBreak, {_str(PAGE_BREAK)} ] ]'
    assert _native_to_latex(test_parameters.container, native) == "before \\clearpage{}"


def test_bookmark_next_to_a_page_break_stays(test_parameters: TestParameters):
    native = f'[ Para [ Span ( "_Toc1" , [ "anchor" ] , [] ) [], {_str(PAGE_BREAK)} ] ]'
    assert "_Toc1" in _native_to_latex(test_parameters.container, native)


def test_section_break_sets_the_page(test_parameters: TestParameters):
    native = f'[ Para [ Str "before" ], Para [ {_str(LANDSCAPE)} ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert "\\clearpage{}\\pdfpagewidth=11.6931in\\pdfpageheight=8.2681in" in flat
    assert "\\newgeometry{layoutwidth=11.6931in,layoutheight=8.2681in,top=1.0000in,right=1.0000in,bottom=1.0000in,left=1.0000in}" in flat


def test_section_break_without_room_for_text_only_breaks(test_parameters: TestParameters):
    """The sentinel is document text, so numbers LaTeX cannot set must not reach it."""
    for numbers in ("100,100,1440,1440,1440,1440", "999999,11906,0,0,0,0", "1,2"):
        marker = f"\\ue020{numbers}\\ue021"
        flat = _native_to_latex(test_parameters.container, f'[ Para [ Str "before" ], Para [ {_str(marker)} ] ]')
        assert flat == "before \\clearpage{}", flat


def test_sentinel_outside_the_body_paragraphs_goes(test_parameters: TestParameters):
    native = f"[ BulletList [ [ Plain [ {_str(f'item{PAGE_BREAK}')} ] ] ] ]"
    flat = _native_to_latex(test_parameters.container, native)
    assert "item" in flat
    assert "clearpage" not in flat


def _pdf(test_parameters: TestParameters, document) -> bytes:
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    return response.content


def _pages(test_parameters: TestParameters, document) -> list[PageObject]:
    return list(PdfReader(io.BytesIO(_pdf(test_parameters, document))).pages)


def _is_landscape(page: PageObject) -> bool:
    return page.mediabox.width > page.mediabox.height


def _text_left_edge(pdf: bytes, page_index: int, text: str) -> float:
    """Where the text starts on the page, in points."""
    text_page = pdfium.PdfDocument(pdf)[page_index].get_textpage()
    start = text_page.get_text_range().index(text)
    return text_page.get_charbox(start)[0]


def test_page_breaks_reach_the_pdf(test_parameters: TestParameters):
    document = Document()
    document.add_paragraph("one")
    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    document.add_paragraph("two")
    pages = _pages(test_parameters, document)
    assert len(pages) == 2
    assert "one" in pages[0].extract_text()
    assert "two" in pages[1].extract_text()


def test_landscape_section_reaches_the_pdf(test_parameters: TestParameters):
    document = Document()
    document.add_paragraph("portrait")
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    document.add_paragraph("landscape")
    pages = _pages(test_parameters, document)
    assert len(pages) == 2
    assert not _is_landscape(pages[0])
    assert _is_landscape(pages[1])
    assert "landscape" in pages[1].extract_text()


def test_table_of_a_landscape_section_spans_its_text_width(test_parameters: TestParameters):
    """The table states the width of the landscape text, which is more than the first section has."""
    document = Document()
    document.sections[0].left_margin = document.sections[0].right_margin = Cm(2)
    document.add_paragraph("portrait")
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.autofit = False
    width = section.page_width - section.left_margin - section.right_margin
    for cell in table.rows[0].cells:
        cell.width = width // 2
    table.rows[0].cells[0].text = "left"
    table.rows[0].cells[1].text = "right"
    pdf = _pdf(test_parameters, document)
    page = PdfReader(io.BytesIO(pdf)).pages[1]
    assert _text_left_edge(pdf, 1, "right") > float(page.mediabox.width) / 2 - 20
