"""Integration tests for ``filters/docx_page_breaks_to_latex.lua``.

The docx reader drops page and section breaks. ``app/docx_page_break_pre_process.py``
writes them as sentinels, and the filter turns those into LaTeX. The unit cases feed
pandoc a native AST inside the container; the end-to-end cases send a DOCX through
the service.
"""

from __future__ import annotations

import io
import json

import pypdfium2 as pdfium
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm
from pypdf import PageObject, PdfReader

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_page_breaks_to_latex.lua"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
OPEN = "\\ue020"
CLOSE = "\\ue021"
PAGE_BREAK = f"{OPEN}{CLOSE}"
LANDSCAPE = f"{OPEN}16838,11906,1440,1440,1440,1440,nextPage{CLOSE}"


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
    return 'Str "' + text.replace(OPEN, "\\57376\\&").replace(CLOSE, "\\57377\\&") + '"'


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
    assert "\\clearpage{} \\pdfpagewidth=11.6931in\\pdfpageheight=8.2681in" in flat
    assert "\\newgeometry{layoutwidth=11.6931in,layoutheight=8.2681in,top=1.0000in,right=1.0000in,bottom=1.0000in,left=1.0000in}" in flat


def test_page_setup_of_a_section_break_goes_after_the_div(test_parameters: TestParameters):
    """docx_paragraphs_to_latex wraps the Div of a formatted paragraph in a TeX group, which would undo the setup."""
    native = f'[ Div ( "" , [] , [] ) [ Para [ {_str(f"closing{LANDSCAPE}")} ] ], Para [ Str "next" ] ]'
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f native -t json --lua-filter={FILTER_PATH} /tmp/test/in.native"])
    assert exit_code == 0, output.decode()
    blocks = json.loads(output)["blocks"]
    assert [block["t"] for block in blocks] == ["Div", "RawBlock", "Para"]
    assert "\\newgeometry" not in json.dumps(blocks[0])
    assert "\\newgeometry" in blocks[1]["c"][1]


def test_section_break_without_room_for_text_only_breaks(test_parameters: TestParameters):
    """The sentinel is document text, so numbers LaTeX cannot set must not reach it."""
    for numbers in ("100,100,1440,1440,1440,1440,nextPage", "999999,11906,0,0,0,0,nextPage", "1,2,nextPage", "16838,11906,0,0,0,0,unknown", "16838,11906,0,0,0,0"):
        marker = f"{OPEN}{numbers}{CLOSE}"
        flat = _native_to_latex(test_parameters.container, f'[ Para [ Str "before" ], Para [ {_str(marker)} ] ]')
        assert flat == "before \\clearpage{}", flat


def test_odd_section_adds_an_empty_page_on_an_even_number(test_parameters: TestParameters):
    marker = f"{OPEN}16838,11906,1440,1440,1440,1440,oddPage{CLOSE}"
    flat = _native_to_latex(test_parameters.container, f'[ Para [ Str "before" ], Para [ {_str(marker)} ] ]')
    assert "\\clearpage{} \\ifodd\\value{page}\\else\\null\\clearpage\\fi\\pdfpagewidth" in flat


def test_section_start_stays_when_the_numbers_are_no_page(test_parameters: TestParameters):
    marker = f"{OPEN}100,100,1440,1440,1440,1440,evenPage{CLOSE}"
    flat = _native_to_latex(test_parameters.container, f'[ Para [ Str "before" ], Para [ {_str(marker)} ] ]')
    assert flat == "before \\clearpage{} \\ifodd\\value{page}\\null\\clearpage\\fi"


def test_page_break_at_the_start_of_a_heading_goes_before_it(test_parameters: TestParameters):
    """A heading argument is copied to the table of contents, so no break stays in it."""
    native = f'[ Para [ Str "before" ], Header 1 ( "" , [] , [] ) [ {_str(f"{PAGE_BREAK}Title")} ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.startswith("before \\clearpage{} \\section{Title}"), flat


def test_page_break_after_the_text_of_a_heading_goes_after_it(test_parameters: TestParameters):
    native = f'[ Header 1 ( "" , [] , [] ) [ {_str(f"Title{PAGE_BREAK}")} ], Para [ Str "after" ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.startswith("\\section{Title}"), flat
    assert flat.endswith("\\clearpage{} after"), flat


def test_page_break_in_the_middle_of_a_heading_goes_after_it(test_parameters: TestParameters):
    native = f'[ Header 1 ( "" , [] , [] ) [ {_str(f"Title{PAGE_BREAK}tail")} ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.startswith("\\section{Titletail}"), flat
    assert flat.endswith("\\clearpage{}"), flat


def test_break_after_a_heading_with_a_moved_break_keeps_its_empty_page(test_parameters: TestParameters):
    """The text after the moved break stands before it in the output, so the page after it is empty."""
    native = f'[ Header 1 ( "" , [] , [] ) [ {_str(f"Title{PAGE_BREAK}tail")} ], Para [ {_str(PAGE_BREAK)} ], Para [ Str "after" ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.endswith("\\clearpage{} \\null\\clearpage{} after"), flat


def test_second_break_moved_behind_a_heading_keeps_its_empty_page(test_parameters: TestParameters):
    native = f'[ Header 1 ( "" , [] , [] ) [ {_str(f"Title{PAGE_BREAK}mid{PAGE_BREAK}tail")} ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.endswith("\\clearpage{} \\null\\clearpage{}"), flat


def test_break_before_a_heading_on_an_empty_page_keeps_it(test_parameters: TestParameters):
    native = f'[ Para [ Str "before" ], Para [ {_str(PAGE_BREAK)} ], Header 1 ( "" , [] , [] ) [ {_str(f"{PAGE_BREAK}Title")} ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.startswith("before \\clearpage{} \\null\\clearpage{} \\section{Title}"), flat


def test_page_break_inside_a_span_of_a_heading_leaves_it(test_parameters: TestParameters):
    native = f'[ Header 1 ( "" , [] , [] ) [ Span ( "" , [] , [] ) [ Str "Title", {_str(PAGE_BREAK)} ] ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.endswith("\\clearpage{}"), flat
    assert "clearpage" not in flat.split("\\label")[0].split("}")[0], flat


def test_heading_holding_only_a_break_goes(test_parameters: TestParameters):
    native = f'[ Para [ Str "before" ], Header 1 ( "" , [] , [] ) [ {_str(PAGE_BREAK)} ] ]'
    assert _native_to_latex(test_parameters.container, native) == "before \\clearpage{}"


def test_section_break_of_a_heading_sets_the_page_after_it(test_parameters: TestParameters):
    native = f'[ Header 1 ( "" , [] , [] ) [ {_str(f"Title{LANDSCAPE}")} ] ]'
    flat = _native_to_latex(test_parameters.container, native)
    assert flat.index("\\section{Title}") < flat.index("\\clearpage{}") < flat.index("\\pdfpagewidth"), flat


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


def _document_with_section_start(start: WD_SECTION):
    document = Document()
    document.add_paragraph("one")
    document.add_section(start)
    document.add_paragraph("two")
    return document


def test_odd_section_starts_on_an_odd_page(test_parameters: TestParameters):
    pages = _pages(test_parameters, _document_with_section_start(WD_SECTION.ODD_PAGE))
    assert len(pages) == 3
    assert "two" in pages[2].extract_text()


def test_even_section_starts_on_an_even_page(test_parameters: TestParameters):
    pages = _pages(test_parameters, _document_with_section_start(WD_SECTION.EVEN_PAGE))
    assert len(pages) == 2
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


def test_landscape_follows_a_centered_paragraph_holding_the_section_break(test_parameters: TestParameters):
    """Word keeps the sectPr in the last paragraph of a section, and a centered paragraph lands in a TeX group."""
    document = Document()
    closing = document.add_paragraph("closing")
    closing.alignment = WD_ALIGN_PARAGRAPH.CENTER
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    document.add_paragraph("landscape")
    holder = closing._p.getnext()
    closing._p.get_or_add_pPr().append(holder.pPr.sectPr)
    document.element.body.remove(holder)
    pages = _pages(test_parameters, document)
    assert len(pages) == 2
    assert _is_landscape(pages[1])


def test_page_break_after_a_heading_reaches_the_pdf(test_parameters: TestParameters):
    document = Document()
    heading = document.add_heading("Chapter", 1)
    heading.add_run().add_break(WD_BREAK.PAGE)
    document.add_paragraph("body")
    pages = _pages(test_parameters, document)
    assert len(pages) == 2
    assert "Chapter" in pages[0].extract_text()
    assert "body" in pages[1].extract_text()


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
