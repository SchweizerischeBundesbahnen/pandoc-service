"""Integration tests for ``filters/docx/docx_document_styles_to_latex.lua`` and app/docx_document_styles.py.

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
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_document_styles_to_latex.lua"
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
    assert "\\AtBeginDocument{\\raggedright\\let\\\\\\@normalcr}" in preamble
    assert "\\hyphenpenalty=10000" in preamble
    assert "\\exhyphenpenalty=0" in preamble, "Word breaks a line after a hyphen as freely as after a space"
    assert "\\frenchspacing" in preamble, "Word puts no wider space after a full stop"
    assert "\\clubpenalty=10000\\widowpenalty=10000" in preamble, "Word leaves no line of a paragraph alone on a page"


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
    assert "\\AtBeginDocument{\\raggedright\\let\\\\\\@normalcr}" in preamble


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


HEADING_STYLES = ("docx-body-paragraph-skip=9", "docx-style-Heading1=size:20,before:18,after:4", "docx-style-Heading2=size:16,before:8,after:4")


def test_next_to_body_text_a_heading_gets_at_least_the_paragraph_skip(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, *HEADING_STYLES))
    assert "\\titlespacing*{\\section}{0pt}{18pt}{9pt}" in preamble
    assert "\\titlespacing*{\\subsection}{0pt}{9pt}{9pt}" in preamble


def test_between_two_headings_only_their_own_spaces_count(test_parameters: TestParameters):
    native = '[Header 1 ("one",[],[]) [Str "One"], Header 2 ("two",[],[]) [Str "Two"], Para [Str "Text"]]'
    body = _latex(test_parameters.container, *HEADING_STYLES, native=native)
    body = body[body.index("\\begin{document}") :]
    assert body.index("\\titlespacing*{\\section}{0pt}{18pt}{8pt}") < body.index("One") < body.index("\\titlespacing*{\\section}{0pt}{18pt}{9pt}")
    assert "\\titlespacing*{\\subsection}" not in body, "a heading before body text keeps the spacing of the preamble"


WORD_HEADINGS = (
    "docx-body-paragraph-skip=9",
    "docx-body-descent=3.38",
    "docx-style-Heading1=size:20,line:24.41,ascent:18.78,before:18,after:4",
    "docx-style-Heading2=size:16,line:19.53,ascent:15.02,before:8,after:4",
)


def test_the_space_before_a_heading_depends_on_the_depth_of_the_line_before(test_parameters: TestParameters):
    """Less by as much as the line before reaches less deep than the line of the heading's font."""
    native = '[Para [Str "Text"], Header 1 ("one",[],[]) [Str "One"]]'
    latex = _latex(test_parameters.container, *WORD_HEADINGS, native=native)
    assert "\\titlespacing*{\\section}{0pt}{\\pdcBefore}{9pt}" in _preamble(latex)
    assert "\\newcommand{\\pdcSetBefore}" in _preamble(latex)
    body = latex[latex.index("\\begin{document}") :]
    assert body.index("\\pdcSetBefore{18bp}{5.63bp}") < body.index("\\section")
    assert "One\\texorpdfstring{\\strut}{}" in body


def test_a_heading_has_a_strut_of_the_line_of_word(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, *WORD_HEADINGS))
    assert "\\setbox\\strutbox\\hbox{\\vrule height 18.78bp depth 5.63bp width 0pt}" in preamble


def test_between_two_headings_their_struts_keep_their_lines_apart(test_parameters: TestParameters):
    native = '[Header 1 ("one",[],[]) [Str "One"], Header 2 ("two",[],[]) [Str "Two"], Para [Str "Text"]]'
    body = _latex(test_parameters.container, *WORD_HEADINGS, native=native)
    body = body[body.index("\\begin{document}") :]
    assert body.index("\\titlespacing*{\\section}{0pt}{\\pdcBefore}{8pt}") < body.index("One") < body.index("\\titlespacing*{\\section}{0pt}{\\pdcBefore}{9pt}")


SPACING = ("docx-body-paragraph-skip=9", "docx-paragraph-spacing=BodyText:Body Text:9:9;Normal:Normal:0:10")


def test_paragraphs_of_normal_stand_the_larger_of_their_spaces_apart(test_parameters: TestParameters):
    """Normal has 10pt after it: one more than the paragraph skip, which Body Text sets."""
    native = '[Div ("",[],[("custom-style","Body Text")]) [Para [Str "Body"]], Para [Str "First"], Para [Str "Second"], Div ("",[],[("pdc-style","Normal")]) [Para [Str "Indented"]]]'
    body = _latex(test_parameters.container, *SPACING, native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    assert "Body First" in body, "after Body Text, Normal has no more space before it than the skip"
    assert "First \\vspace{1bp} Second \\vspace{1bp} Indented" in body


def test_paragraphs_of_one_style_in_one_div_get_their_gap_each(test_parameters: TestParameters):
    native = '[Div ("",[],[("custom-style","Wide")]) [Para [Str "One"], Para [Str "Two"]]]'
    body = _latex(test_parameters.container, "docx-body-paragraph-skip=9", "docx-paragraph-spacing=Wide:Wide:0:12", native=native)
    assert "One\n\n\\vspace{3bp}\n\nTwo" in body


def test_a_table_ends_as_deep_as_word_sets_its_last_row(test_parameters: TestParameters):
    """A stretched row reaches 3.4 * 18.25 / 14.65 = 4.24 deep, Word's 3.4 + 1.8: the rest goes below the table."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", "docx-table-cell-space=3.6"))
    assert "\\AddToHook{env/longtable/after}{\\par\\prevdepth=\\dp\\strutbox}" in preamble
    assert "\\newcommand{\\pdcCompactCellBottom}{0.9645bp}" in preamble
    assert "\\newcommand{\\pdcCompactCellDepth}{\\rule[-\\dp\\@arstrutbox]{0pt}{0pt}}" in preamble


def test_a_paragraph_after_a_table_gets_only_its_own_space_before(test_parameters: TestParameters):
    """A table has no space of its own: Normal after it starts right below, a table after Normal 10pt below."""
    cell = 'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [Str "cell"]]'
    body = f'TableBody ("",[],[]) (RowHeadColumns 0) [] [Row ("",[],[]) [{cell}]]'
    table = f'Table ("",[],[]) (Caption Nothing []) [(AlignDefault,ColWidthDefault)] (TableHead ("",[],[]) []) [{body}] (TableFoot ("",[],[]) [])'
    native = f'[{table}, Para [Str "After"], {table}]'
    body = _latex(test_parameters.container, *SPACING, native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    after = body.index("After")
    assert body.index("\\end{longtable}") < body.index("\\vspace{-9bp}") < after
    assert after < body.index("\\vspace{1bp}") < body.index("\\begin{longtable}", after)


def test_the_lines_of_a_paragraph_stay_together_after_a_heading(test_parameters: TestParameters):
    """LaTeX sets the club penalty back after a heading, and \\\\ under \\raggedright ends the paragraph."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-justified=false"))
    assert "\\@clubpenalty=10000" in preamble
    assert "\\let\\\\\\@normalcr" in preamble


def test_a_page_starts_as_far_down_as_word_sets_its_first_line(test_parameters: TestParameters):
    """The first line of a page at the ascent; the first of a document without a title its space before lower."""
    metadata = (*SPACING, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-ascent=11.25", "docx-body-descent=3.4")
    latex = _latex(test_parameters.container, *metadata, native='[Div ("",[],[("custom-style","Body Text")]) [Para [Str "First"]], Para [Str "Next"]]')
    assert "\\setlength{\\topskip}{11.25bp}" in _preamble(latex)
    body = latex[latex.index("\\begin{document}") :]
    assert "\\rule{0pt}{20.25bp}\\strut{}First" in body
    assert body.count("\\rule{0pt}") == 1


def test_an_entry_of_a_list_of_figures_gets_the_space_of_the_paragraph_before(test_parameters: TestParameters):
    """An entry is set without the paragraph skip, with 5pt after it: Body Text before it keeps its 9pt."""
    entry = 'Div ("",[],[("custom-style","PandocToc__ENTRY")]) [Para [Str "Figure"]]'
    native = f'[Div ("",[],[("custom-style","Body Text")]) [Para [Str "Before"]], {entry}, {entry}, Para [Str "After"]]'
    body = " ".join(_latex(test_parameters.container, *SPACING, native=native).split())
    assert body.count("\\vspace{9bp}") == 1
    assert "\\vspace{-9bp}" in body, "Normal after an entry: the larger of 5 and 0, less the skip and the 5pt the entry ends with"


def test_a_later_title_paragraph_gets_the_space_of_word(test_parameters: TestParameters):
    """The space between paragraphs puts the title's space before it, and its line comes up by what TeX adds."""
    title = "docx-style-Title=size:28,line:34.18,ascent:26.29,before:0,after:4"
    native = '[Para [Str "Before"], Div ("",[],[("custom-style","Title")]) [Para [Str "Later"]], Para [Str "After"]]'
    body = " ".join(_latex(test_parameters.container, *SPACING, title, "docx-body-descent=3.4", native=native).split())
    assert "\\vspace{1bp} \\par\\ifdim\\prevdepth<7.89bp\\ifdim\\prevdepth>-1000pt\\vskip\\dimexpr\\prevdepth-7.89bp\\relax\\fi\\fi" in body
    assert "\\par\\endgroup \\vspace{-5bp} \\strut{}After" in body, "the larger of 4 and 0, less the skip"


def test_the_items_of_a_tight_list_get_the_space_of_compact(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-list-item-space=1.8"))
    assert "\\renewcommand{\\tightlist}{\\setlength{\\itemsep}{1.8pt}\\setlength{\\parskip}{0pt}}" in preamble


def test_a_bullet_reaches_as_high_as_its_font(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-list-bullet-ascents=12.06,0,10.79,0"))
    assert "\\renewcommand{\\labelitemi}{\\rule[-\\dp\\strutbox]{0pt}{\\dimexpr12.06bp+\\dp\\strutbox\\relax}\\pdclabelitemi}" in preamble
    assert "labelitemiii}{\\rule" in preamble
    assert "labelitemii}{" not in preamble


def test_the_full_stop_which_ends_a_paragraph_hangs_in_the_margin(test_parameters: TestParameters):
    native = '[Para [Str "One", Space, Str "end."], Para [Str "No", Space, Str "stop"], Para [Str "."]]'
    body = _latex(test_parameters.container, native=native)
    body = body[body.index("\\begin{document}") :]
    assert "One end\\rlap{.}" in body
    assert "No stop\n" in body
    assert "\n.\n" in body, "a paragraph of a full stop alone keeps it"


def test_an_equation_alone_is_a_centred_paragraph(test_parameters: TestParameters):
    native = '[Para [Math DisplayMath "a^{2}"], Para [Str "Text", Space, Math DisplayMath "b"]]'
    body = _latex(test_parameters.container, native=native)
    body = body[body.index("\\begin{document}") :]
    assert "{\\centering\\renewcommand{\\arraystretch}{1}\\(\\displaystyle a^{2}\\)\\par}" in body
    assert "\\[b\\]" in body, "an equation within text stays a display"


def test_a_numbered_heading_style_sets_its_number_in_its_hanging_indent(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading1=size:20,number:{1:Roman},numberhang:21.6", "docx-style-Heading2=size:16,number:{1:Roman}.{2:arabic}"))
    assert "\\renewcommand{\\thesection}{\\Roman{section}}" in preamble
    assert "\\titleformat{\\section}[hang]{\\normalfont\\fontsize{20bp}{24bp}\\selectfont\\filright}{\\makebox[21.6pt][l]{\\thesection}}{0pt}{}" in preamble
    assert "\\renewcommand{\\thesubsection}{\\Roman{section}.\\arabic{subsection}}" in preamble
    assert "{\\thesubsection\\quad}{0pt}{}" in preamble, "a number without a hanging indent is followed by a space"
    assert "\\AtBeginDocument{\\setcounter{secnumdepth}{2}}" in preamble


def test_a_number_which_could_carry_latex_is_left_out(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading1=size:20,number:{1:input}", "docx-style-Heading2=size:16,number:\\x{1:arabic}"))
    assert "\\thesection" not in preamble
    assert "\\thesubsection" not in preamble
    assert "\\AtBeginDocument{\\setcounter{secnumdepth}" not in preamble


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


def test_the_title_stands_as_far_down_as_the_ascent_of_its_font(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, "title=Doc", "docx-style-Title=size:28,ascent:26.29,after:4", "docx-body-paragraph-skip=9")
    assert "\\pdcTitleStyle\\rule{0pt}{26.29bp}\\@title\\par\\endgroup\\vspace{9pt}" in _preamble(latex)


def test_a_table_of_contents_after_the_title_takes_only_the_space_of_the_title(test_parameters: TestParameters):
    native = '[Div ("",[],[("custom-style","PandocToc__CONTENTS_3")]) [Para [Str "PandocToc__CONTENTS_3"]]]'
    latex = _latex(test_parameters.container, "title=Doc", "docx-style-Title=size:28,after:4", "docx-body-paragraph-skip=9", native=native)
    assert "\\pdcTitleStyle\\@title\\par\\endgroup\\vspace{4pt}" in _preamble(latex)


def test_the_title_and_the_next_line_are_set_with_struts_of_word(test_parameters: TestParameters):
    """The struts keep the lines Word's distance apart: after the title only the larger space counts, less the skip TeX adds."""
    title = "docx-style-Title=size:28,line:34.18,ascent:26.29,after:4"
    preamble = _preamble(_latex(test_parameters.container, "title=Doc", title, "docx-body-paragraph-skip=9", "docx-body-descent=3.4"))
    assert "\\pdcTitleStyle\\strut\\@title\\strut\\par\\endgroup\\vspace{0pt}" in preamble


def test_a_table_of_contents_after_the_title_gets_its_space_after(test_parameters: TestParameters):
    native = '[Div ("",[],[("custom-style","PandocToc__CONTENTS_3")]) [Para [Str "PandocToc__CONTENTS_3"]]]'
    title = "docx-style-Title=size:28,line:34.18,ascent:26.29,after:4"
    latex = _latex(test_parameters.container, "title=Doc", title, "docx-body-paragraph-skip=9", "docx-body-descent=3.4", native=native)
    assert "\\@title\\strut\\par\\endgroup\\vspace{4pt}" in _preamble(latex)


def test_a_paragraph_starts_and_ends_with_a_strut(test_parameters: TestParameters):
    body = _latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", native='[Para [Str "Text"]]')
    assert "\\strut{}Text\\strut{}" in body


def test_the_caption_style_defines_its_space_before(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Caption=size:9,before:0,after:10"))
    assert "\\newcommand{\\pdcCaptionBefore}{0pt}" in preamble
    assert "\\newcommand{\\pdcCaptionAfter}{10pt}" in preamble


def test_the_body_size_and_line_redefine_normalsize(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65"))
    assert "\\renewcommand{\\normalsize}{\\fontsize{12bp}{14.65bp}\\selectfont" in preamble
    assert "\\normalsize\n" in preamble


def test_an_equation_gets_no_space_of_a_display_around_it(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65"))
    for skip in ("abovedisplayskip", "belowdisplayskip", "abovedisplayshortskip", "belowdisplayshortskip"):
        assert f"\\{skip}=0pt" in preamble


def test_a_line_of_the_body_reaches_as_deep_as_the_descent_of_its_font(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4"))
    assert "\\selectfont\\setbox\\strutbox\\hbox{\\vrule height 11.25bp depth 3.4bp width 0pt}" in preamble
    assert "\\lineskip=0pt\\lineskiplimit=0pt" in preamble


def test_a_table_starts_below_the_descent_of_the_line_before_it(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4"))
    assert "\\AddToHook{env/longtable/before}{\\par\\ifdim\\prevdepth>-1000pt\\ifdim\\prevdepth<\\dp\\strutbox" in preamble


def test_a_body_without_its_descent_keeps_the_strut_of_latex(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65"))
    assert "\\strutbox" not in preamble
    assert "\\lineskip=0pt" not in preamble


def test_a_table_of_compact_rows_starts_higher_by_the_share_of_the_stretch(test_parameters: TestParameters):
    # The stretch puts 11.25 * 3.6 / 14.65 = 2.764 above the line, Word 1.8
    preamble = _preamble(_latex(test_parameters.container, "docx-body-baseline=14.65", "docx-body-descent=3.4", "docx-table-cell-space=3.6"))
    assert "\\newcommand{\\pdcCompactCellRaise}{0.9645bp}" in preamble


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
    assert "\\newcommand{\\pdcPlainCellSpace}{3.6pt}" in preamble
    assert "\\newcommand{\\pdcCompactCellStretch}{1.2457}" in preamble


def test_a_cell_paragraph_without_a_style_takes_the_space_of_normal(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-table-cell-space=3.6", "docx-table-plain-cell-space=10", "docx-body-baseline=14.65", "docx-body-size=12"))
    assert "\\newcommand{\\pdcPlainCellSpace}{10pt}" in preamble
    assert "\\newcommand{\\pdcCompactCellStretch}{1.2457}" in preamble


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


def test_a_document_of_grid_tables_only_compiles(test_parameters: TestParameters):
    """The grid is raw LaTeX: the template must still load the packages of a table."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    document = Document()
    table = document.add_table(rows=2, cols=2)
    for cell in table._cells:
        borders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            line = OxmlElement(f"w:{side}")
            line.set(qn("w:val"), "single")
            borders.append(line)
        cell._tc.get_or_add_tcPr().append(borders)
        cell.text = "grid cell"
    text = "".join(page.extract_text() for page in PdfReader(io.BytesIO(_docx_to_pdf(test_parameters, document))).pages)
    assert text.count("grid cell") == 4, text
