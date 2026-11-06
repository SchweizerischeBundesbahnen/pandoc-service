"""Integration tests for ``filters/docx_document_styles_to_latex.lua`` and app/docx_document_styles.py.

The filter tests run pandoc inside the pandoc-service container with the
metadata app/docx_document_styles.py passes, and read the preamble of the
standalone LaTeX. The end-to-end tests convert a DOCX through the service.
"""

from __future__ import annotations

import io
import shlex

from docx import Document
from docx.shared import Pt
from pypdf import PdfReader

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_document_styles_to_latex.lua"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
HEADER_FILTER = "/tmp/test/header.lua"
# A filter before ours which adds header-includes, as the other docx filters do
HEADER_SCRIPT = r'function Meta(m) m["header-includes"] = pandoc.MetaList({pandoc.MetaBlocks({pandoc.RawBlock("latex", "\\newcommand{\\other}{}")})}) return m end'


def _latex(container, *metadata: str, other_preamble: bool = False, native: str | None = None) -> str:
    """The standalone LaTeX of a one-line document, or of the native document given, with the metadata given."""
    if native is None:
        container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf 'text\\n' > /tmp/test/in.md"])
    else:
        container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/in.md << 'EOF'\n{native}\nEOF"])
    filters = ""
    if other_preamble:
        container.exec_run(["sh", "-c", f"cat > {HEADER_FILTER} << 'EOF'\n{HEADER_SCRIPT}\nEOF"])
        filters = f"--lua-filter={HEADER_FILTER} "
    options = " ".join(f"-M {shlex.quote(item)}" for item in metadata)
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f {'markdown' if native is None else 'native'} -t latex {filters}--lua-filter={FILTER_PATH} {options} /tmp/test/in.md"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode("utf-8")


def _preamble(latex: str) -> str:
    return latex[: latex.index("\\begin{document}")]


def test_word_layout_goes_into_the_preamble(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-paragraph-skip=9", "docx-body-justified=false", "docx-body-hyphenated=false"))
    assert "\\setlength{\\parskip}{9pt}" in preamble
    assert "\\AtBeginDocument{\\raggedright}" in preamble
    assert "\\hyphenpenalty=10000" in preamble
    assert "\\exhyphenpenalty" not in preamble, "Word breaks a line after a hyphen of the text"
    assert "\\frenchspacing" in preamble, "Word puts no wider space after a full stop"


def test_a_justified_hyphenated_style_keeps_latex_paragraphs(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-paragraph-skip=0", "docx-body-justified=true", "docx-body-hyphenated=true"))
    assert "\\setlength{\\parskip}{0pt}" in preamble
    assert "\\raggedright" not in preamble
    assert "\\hyphenpenalty" not in preamble


def test_without_metadata_nothing_changes(test_parameters: TestParameters):
    assert _latex(test_parameters.container) == _latex(test_parameters.container, "unrelated=1")


def test_a_skip_which_is_no_number_is_left_out(test_parameters: TestParameters):
    assert _latex(test_parameters.container, "docx-body-paragraph-skip=wide", "docx-body-justified=true") == _latex(test_parameters.container)


def test_the_preamble_of_other_filters_stays(test_parameters: TestParameters):
    """Setting header-includes as a variable replaces what filters add, the filter must add to it instead."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-justified=false", other_preamble=True))
    assert "\\newcommand{\\other}{}" in preamble
    assert "\\AtBeginDocument{\\raggedright}" in preamble


def test_the_metadata_does_not_reach_the_document(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "docx-body-paragraph-skip=9", "docx-body-justified=false", "docx-body-hyphenated=false")
    assert "docx-body" not in latex


def _docx_to_pdf(test_parameters: TestParameters, document) -> bytes:
    docx = io.BytesIO()
    document.save(docx)
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), DOCX_MIME)})
    assert response.status_code == 200, response.text
    return response.content


def _fonts(pdf: bytes) -> set[str]:
    """The base fonts the pages of a PDF use, without their subset tag."""
    fonts: set[str] = set()
    for page in PdfReader(io.BytesIO(pdf)).pages:
        resources = page.get("/Resources")
        font_dictionary = resources.get_object().get("/Font") if resources is not None else None
        if font_dictionary is None:
            continue
        for font in font_dictionary.get_object().values():
            fonts.add(str(font.get_object()["/BaseFont"]).lstrip("/").split("+")[-1])
    return fonts


def test_the_pdf_sets_the_body_in_the_font_of_the_docx(test_parameters: TestParameters):
    document = Document()
    document.styles["Normal"].font.name = "Liberation Sans"
    document.styles["Normal"].font.size = Pt(12)
    document.add_paragraph("The body text of the document.")
    fonts = _fonts(_docx_to_pdf(test_parameters, document))
    assert any(font.startswith("LiberationSans") for font in fonts), fonts


def test_a_font_the_container_lacks_keeps_latin_modern(test_parameters: TestParameters):
    document = Document()
    document.styles["Normal"].font.name = "No Such Font"
    document.add_paragraph("The body text of the document.")
    fonts = _fonts(_docx_to_pdf(test_parameters, document))
    assert any(font.startswith("LMRoman") for font in fonts), fonts


def test_a_heading_style_becomes_its_titlesec_format(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading1=font:DejaVu Sans,size:20,line:24.4,bold:false,italic:true,color:0F4761,before:18,after:9,align:center"))
    assert "\\usepackage{titlesec}" in preamble
    assert "\\newfontfamily{\\docxHeadingFontA}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}]{DejaVu Sans}" in preamble
    assert "\\titleformat{\\section}[block]{\\normalfont\\docxHeadingFontA\\fontsize{20bp}{24.4bp}\\selectfont\\itshape\\color[HTML]{0F4761}\\filcenter}{}{0pt}{}" in preamble
    assert "\\titlespacing*{\\section}{0pt}{18pt}{9pt}" in preamble


DEEP_HEADINGS = '[Header 6 ("six",[],[]) [Str "Six"], Header 7 ("",[],[]) [Str "Seven"], Header 10 ("ten",[],[]) [Str "Ten"]]'


def test_a_heading_below_level_5_gets_a_level_of_its_own(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, native=DEEP_HEADINGS)
    preamble, body = _preamble(latex), latex[latex.index("\\begin{document}") :]
    assert "\\titleclass{\\pdcHeadingF}{straight}[\\subparagraph]" in preamble
    assert "\\titleclass{\\pdcHeadingI}{straight}[\\pdcHeadingH]" in preamble
    assert "\\pdcHeadingF{Six}\\label{six}" in body
    assert "\\pdcHeadingG{Seven}\n" in body, "no label without an identifier"
    assert "\\pdcHeadingI{Ten}\\label{ten}" in body, "Word has nine levels"


def test_a_deep_heading_takes_the_format_of_its_style(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading7=size:12,italic:true,color:595959,before:2,after:0", native=DEEP_HEADINGS))
    assert "\\titleformat{\\pdcHeadingG}[block]{\\normalfont\\fontsize{12bp}{14.4bp}\\selectfont\\itshape\\color[HTML]{595959}\\filright}{}{0pt}{}" in preamble
    assert "\\titlespacing*{\\pdcHeadingG}{0pt}{2pt}{0pt}" in preamble


def test_a_document_without_deep_headings_gets_no_levels_of_them(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading7=size:12"))
    assert "pdcHeading" not in preamble
    assert "titlesec" not in preamble


def test_fields_which_could_carry_latex_are_left_out(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading2=font:\\evil,size:x,color:12345G,align:sideways"))
    assert "\\evil" not in preamble
    assert "\\color" not in preamble
    assert "\\titleformat{\\subsection}[block]{\\normalfont\\filright}{}{0pt}{}" in preamble, "a style of no valid field keeps the plain font, left"


def test_the_title_style_sets_the_title_and_a_later_title_paragraph(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf -- '---\\ntitle: The Title\\n---\\n\\n::: {custom-style=\"Title\"}\\nLater Title\\n:::\\n' > /tmp/test/title.md"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f markdown -t latex --lua-filter={FILTER_PATH} -M 'docx-style-Title=size:28,bold:false,italic:false,before:0,after:9,align:center' /tmp/test/title.md"])
    assert exit_code == 0, output.decode()
    latex = output.decode()
    assert "\\newcommand{\\pdcTitleStyle}{\\centering\\normalfont\\fontsize{28bp}{33.6bp}\\selectfont}" in latex
    assert "\\renewcommand{\\maketitle}" in latex
    body = latex[latex.index("\\begin{document}") :]
    assert "\\begingroup\\pdcTitleStyle" in body
    assert "Later Title" in body


def test_the_body_size_and_line_redefine_normalsize(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65"))
    assert "\\renewcommand{\\normalsize}{\\fontsize{12bp}{14.65bp}\\selectfont" in preamble
    assert "\\normalsize\n" in preamble


def test_an_equation_gets_no_space_of_a_display_around_it(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65"))
    for skip in ("abovedisplayskip", "belowdisplayskip", "abovedisplayshortskip", "belowdisplayshortskip"):
        assert f"\\{skip}=0pt" in preamble


def test_a_body_size_without_its_line_is_left_out(test_parameters: TestParameters):
    assert "\\renewcommand{\\normalsize}" not in _preamble(_latex(test_parameters.container, "docx-body-size=12"))


def test_the_link_colour_is_defined_for_the_template(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-link-color=156082"))
    assert "\\definecolor{pdcLink}{HTML}{156082}" in preamble


def test_a_link_colour_which_is_no_hex_colour_is_left_out(test_parameters: TestParameters):
    assert "pdcLink" not in _latex(test_parameters.container, "docx-link-color=}\\evil")


def test_a_document_without_headers_has_no_page_numbers(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-page-numbers=false"))
    assert "\\pagestyle{empty}" in preamble
    assert "\\let\\ps@plain\\ps@empty" in preamble


def test_a_table_gets_the_space_and_rows_of_word(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-table-cell-space=3.6", "docx-body-baseline=14.65", "docx-body-size=12"))
    assert "\\ifdefined\\LTpre\\setlength{\\LTpre}{\\parskip}\\setlength{\\LTpost}{0pt}\\fi" in preamble
    assert "\\ifdefined\\aboverulesep\\setlength{\\aboverulesep}{0pt}\\setlength{\\belowrulesep}{0pt}\\fi" in preamble
    assert "\\renewcommand{\\arraystretch}{1.2457}" in preamble


def test_a_document_without_tables_still_compiles(test_parameters: TestParameters):
    """The lengths of the table packages exist only with a table: each is set only where it is defined."""
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf 'Just text.\\n' > /tmp/test/plain.md"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"cd /tmp/test && {PANDOC_PATH} -f markdown -o plain.pdf --pdf-engine=tectonic --lua-filter={FILTER_PATH} -M docx-table-cell-space=3.6 -M docx-body-baseline=14.65 -M docx-body-size=12 plain.md"]
    )
    assert exit_code == 0, output.decode()


def test_a_caption_is_set_in_the_caption_style(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf '::: {custom-style=\"PandocCaption\"}\\nTable 1 A caption\\n:::\\n' > /tmp/test/caption.md"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f markdown -t latex --lua-filter={FILTER_PATH} -M 'docx-style-Caption=size:12,bold:false,italic:true,before:9,after:9,align:left' /tmp/test/caption.md"])
    assert exit_code == 0, output.decode()
    latex = output.decode()
    assert "\\newcommand{\\pdcCaptionStyle}{\\normalfont\\fontsize{12bp}{14.4bp}\\selectfont\\itshape}" in latex
    body = latex[latex.index("\\begin{document}") :]
    assert "\\par\\begingroup\\pdcCaptionStyle" in body
    assert "Table 1 A caption" in body


def test_without_a_caption_style_a_caption_stays_plain(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf '::: {custom-style=\"PandocCaption\"}\\nTable 1 A caption\\n:::\\n' > /tmp/test/caption.md"])
    _, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/caption.md"])
    assert "pdcCaptionStyle" not in output.decode()


def test_a_page_break_becomes_clearpage(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf 'one\\n\\n::: {custom-style=\"PandocPageBreak\"}\\nPandocPageBreak\\n:::\\n\\ntwo\\n' > /tmp/test/break.md"])
    _, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/break.md"])
    latex = output.decode()
    assert "\\clearpage" in latex
    assert "PandocPageBreak" not in latex


def test_the_pdf_breaks_the_page_where_the_docx_does(test_parameters: TestParameters):
    from docx.enum.text import WD_BREAK

    document = Document()
    document.add_paragraph("First page.")
    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    document.add_paragraph("Second page.")
    reader = PdfReader(io.BytesIO(_docx_to_pdf(test_parameters, document)))
    assert len(reader.pages) == 2
    assert "Second page." in reader.pages[1].extract_text()


def test_a_section_marker_switches_the_page_geometry(test_parameters: TestParameters):
    container = test_parameters.container
    container.exec_run(["sh", "-c", 'mkdir -p /tmp/test && printf \'::: {custom-style="PandocSection"}\\n842 595 72 72 72 72\\n:::\\n\\n::: {custom-style="PandocSection"}\\n842 x 72 72 72 72\\n:::\\n\' > /tmp/test/section.md'])
    _, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/section.md"])
    latex = output.decode()
    body = latex[latex.index("\\begin{document}") :]
    assert "\\pdcPageGeometry{842}{595}{72}{72}{72}{72}" in body
    assert body.count("\\clearpage") == 1, "a marker which is no geometry only breaks the page"
    assert "\\newcommand{\\pdcPageGeometry}[6]" in latex


def test_the_pdf_gives_a_section_its_own_page(test_parameters: TestParameters):
    from docx.enum.section import WD_ORIENT, WD_SECTION

    document = Document()
    document.add_paragraph("Portrait page.")
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    document.add_paragraph("Landscape page.")
    reader = PdfReader(io.BytesIO(_docx_to_pdf(test_parameters, document)))
    sizes = [(round(float(page.mediabox.width)), round(float(page.mediabox.height))) for page in reader.pages]
    assert len(sizes) == 2
    assert sizes[1][0] > sizes[1][1] > 0, sizes
    assert sizes[0][1] > sizes[0][0], sizes
    assert "Landscape page." in reader.pages[1].extract_text()


def test_a_straight_quote_of_the_text_stays_straight(test_parameters: TestParameters):
    """fontspec's TeX ligatures turn " into a closing quote; Word prints the character the text holds."""
    document = Document()
    document.styles["Normal"].font.name = "Liberation Sans"
    document.add_paragraph('Click "Login" to sign in.')
    text = PdfReader(io.BytesIO(_docx_to_pdf(test_parameters, document))).pages[0].extract_text()
    assert 'Click "Login" to sign in.' in text, text


def test_a_cell_gets_the_margin_of_the_table_style(test_parameters: TestParameters):
    assert "\\setlength{\\tabcolsep}{5.4pt}" in _preamble(_latex(test_parameters.container, "docx-table-cell-margin=5.4", "docx-table-cell-space=3.6"))
