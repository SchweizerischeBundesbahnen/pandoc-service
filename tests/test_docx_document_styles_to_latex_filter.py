"""Integration tests for ``filters/docx/docx_document_styles_to_latex.lua`` and app/docx_document_styles.py.

The filter tests run pandoc inside the pandoc-service container with the
metadata app/docx_document_styles.py passes, and read the preamble of the
standalone LaTeX. The end-to-end tests convert a DOCX through the service.
"""

from __future__ import annotations

import io
import shlex

import pypdfium2 as pdfium
import pytest
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
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


def test_an_equation_takes_metric_math_in_place_of_cambria_math(test_parameters: TestParameters):
    """The image builds Metric Math, Cambria Math to the layout of an equation: Latin Modern Math sets it narrower."""
    document = Document()
    math = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    document.add_paragraph()._p.append(parse_xml(f"<m:oMathPara {math}><m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath></m:oMathPara>"))
    fonts = _fonts(_docx_to_pdf(test_parameters, document))
    assert "MetricMath-Regular" in fonts, fonts


def test_a_heading_style_becomes_its_titlesec_format(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading1=font:DejaVu Sans,size:20,line:24.4,bold:false,italic:true,color:0F4761,before:18,after:9,align:center"))
    assert "\\usepackage{titlesec}" in preamble
    assert "\\newfontfamily{\\docxHeadingFontA}[Kerning=Off,Scale=1,WordSpace={1,1,0},Ligatures={TeXOff,NoCommon}]{DejaVu Sans}" in preamble
    # The colour inside the line of the heading: before it, its special would let TeX break the page after the heading
    assert "\\titleformat{\\section}[block]{\\normalfont\\docxHeadingFontA\\pdcNoScriptFont{}\\fontsize{20bp}{24.4bp}\\selectfont\\itshape\\filcenter}{}{0pt}{\\color[HTML]{0F4761}}" in preamble
    assert "\\titlespacing*{\\section}{0pt}{18pt}{9pt}" in preamble


def test_a_heading_sets_its_scripts_in_its_font(test_parameters: TestParameters):
    """A heading in another font than the body's sets its scripts as Word does in that font, or as LaTeX does."""
    with_scripts = "docx-style-Heading1=font:DejaVu Serif,size:20,script:0.6997/0.4800/0.1396/1.0000/0.2358"
    without = "docx-style-Heading2=font:No Scripts,size:16"
    preamble = _preamble(_latex(test_parameters.container, with_scripts, without))
    assert "\\docxHeadingFontA\\pdcScriptFont{0.6997}{0.4800}{0.1396}{1.0000}{0.2358}\\fontsize{20bp}" in preamble
    assert "\\docxHeadingFontB\\pdcNoScriptFont{}\\fontsize{16bp}" in preamble


def test_a_script_field_that_could_carry_latex_is_left_out(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading1=font:DejaVu Serif,size:20,script:1/2/3/4/5}\\input{y}"))
    assert "\\docxHeadingFontA\\pdcNoScriptFont{}\\fontsize{20bp}" in preamble


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
    assert "\\titleformat{\\pdcHeadingG}[block]{\\normalfont\\fontsize{12bp}{14.4bp}\\selectfont\\itshape\\filright}{}{0pt}{\\color[HTML]{595959}}" in preamble
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


def test_a_heading_right_before_a_table_keeps_only_its_own_space_after(test_parameters: TestParameters):
    """Word puts the larger of the space after a heading and before the next block between them, and a table has none."""
    cell = 'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [Str "Cell"]]'
    table = f'Table ("",[],[]) (Caption Nothing []) [(AlignDefault,ColWidthDefault)] (TableHead ("",[],[]) []) [TableBody ("",[],[]) (RowHeadColumns 0) [] [Row ("",[],[]) [{cell}]]] (TableFoot ("",[],[]) [])'
    native = '[Header 1 ("one",[],[]) [Str "One"], ' + table + ', Header 1 ("two",[],[]) [Str "Two"], Div ("",[],[("custom-style","Body Text")]) [Para [Str "Text"]]]'
    body = _latex(test_parameters.container, *HEADING_STYLES, SPACING[1], native=native)
    body = body[body.index("\\begin{document}") :]
    assert body.index("\\titlespacing*{\\section}{0pt}{18pt}{4pt}") < body.index("One") < body.index("\\titlespacing*{\\section}{0pt}{18pt}{9pt}")
    assert body.count("\\titlespacing") == 2, "before Body Text a heading keeps the spacing of the preamble"


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
    """A row of Compact reaches Word's 3.4 + 1.8 deep: the stretch makes the depth."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", "docx-table-cell-space=3.6"))
    assert "\\AddToHook{env/longtable/after}{\\par\\prevdepth=\\dp\\strutbox}" in preamble
    assert "\\newcommand{\\pdcCompactCellStretch}{1.5294}" in preamble
    assert "\\newcommand{\\pdcCompactCellDepth}{\\rule[-\\dp\\@arstrutbox]{0pt}{0pt}}" in preamble
    # Below an equation which ends a cell: what the equation is deeper than a line, up to what the row is
    space = "\\newcommand{\\pdcCompactEquationSpace}{\\dimen@\\pdcCellLineDepth\\ifdim\\dimen@>\\dp\\@arstrutbox\\dimen@\\dp\\@arstrutbox\\fi"
    assert space + "\\advance\\dimen@-\\dp\\strutbox\\ifdim\\dimen@>\\z@\\vadjust{\\kern\\dimen@}\\fi}" in preamble


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
    assert body.count("\\rule{0pt}{2") == 1


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
    latex = _latex(test_parameters.container, native=native)
    body = latex[latex.index("\\begin{document}") :]
    assert "\\def\\pdcDisplayLine#1{{\\centering\\renewcommand{\\arraystretch}{1}\\setlength{\\jot}{1.35pt}#1\\par}}" in _preamble(latex)
    assert "\\pdcDisplayLine{\\(\\displaystyle a^{2}\\)}" in body
    assert "\\[b\\]" in body, "an equation within text stays a display"


def test_an_equation_sets_its_rows_and_fractions_as_word(test_parameters: TestParameters):
    """Word sets the rows of a matrix a fixed pitch apart, the equation tight to its content and a little space
    above it, and a fraction inside a fraction as large as the outer one."""
    latex = _latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", native='[Para [Math DisplayMath "a^{2}"]]')
    assert "\\newcommand{\\pdcEquationRow}{14.1bp}\\newcommand{\\pdcEquationGap}{3bp}" in _preamble(latex)
    assert "\\lineskip\\pdcEquationGap\\baselineskip\\pdcEquationRow\\lineskiplimit\\z@" in _preamble(latex), "rows which would overlap keep a gap"
    assert "\\genfrac{}{}{}{1}{\\displaystyle #1}{\\displaystyle #2}" in _preamble(latex)
    assert "\\let\\frac\\pdcEquationFrac\\pdcEquationArrays\\sbox{\\pdcEquationBox}{#1}\\rule{0pt}{\\dimexpr\\ht\\pdcEquationBox+1.62bp\\relax}" in _preamble(latex)


def test_the_rows_of_a_matrix_are_the_line_of_the_math_font_apart(test_parameters: TestParameters):
    """Word sets the rows of a matrix the line of its math font apart: 1.2 times the size in Latin Modern Math, 14.4pt at 12pt."""
    latex = _latex(test_parameters.container, "docx-body-size=12", "docx-math-top=0.806", "docx-math-descent=0.394", native='[Para [Math DisplayMath "a^{2}"]]')
    assert "\\newcommand{\\pdcEquationRow}{14.4bp}" in _preamble(latex)


def test_an_equation_is_as_high_and_as_deep_as_a_line_of_its_font(test_parameters: TestParameters):
    """Word sets E=mc² on the line of Cambria Math, and the next line right below its descent: not a line of the body."""
    metadata = ("docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.381", "docx-math-top=0.9502", "docx-math-descent=0.2222")
    preamble = _preamble(_latex(test_parameters.container, *metadata, native='[Para [Math DisplayMath "a^{2}"]]'))
    assert "\\rule[-2.6664bp]{0pt}{14.0688bp}\\usebox{\\pdcEquationBox}\\par\\prevdepth\\dp\\strutbox}}" in preamble, preamble
    # The last paragraph of a cell: the cell ends the paragraph, and the row with it. A cell sets the line lower
    # An equation which starts its cell is in the first line of the cell, and hangs from the top of the row
    last = "\\def\\pdcDisplayLineEnd#1{\\ifpdcCellFirstLine\\else\\par\\fi\\centering\\arraybackslash"
    assert last in preamble and "\\rule[-2.6664bp]{0pt}{15.6888bp}\\usebox{\\pdcEquationBox}}}" in preamble, preamble


def _numerator_top(container, equation: str) -> float:
    """Where the first character of the PDF of a display equation stands, from the top of the page."""
    native = f'[Para [Str "Text"], Para [Math DisplayMath "{equation}"]]'
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/cancel.native << 'EOF'\n{native}\nEOF"])
    command = f"{PANDOC_PATH} -s -f native --lua-filter={FILTER_PATH} -M docx-body-size=12 -M docx-body-baseline=14.65 --pdf-engine=tectonic -o /tmp/test/cancel.pdf /tmp/test/cancel.native"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    _, pdf = container.exec_run(["cat", "/tmp/test/cancel.pdf"])
    page = pdfium.PdfDocument(pdf)[0]
    text = page.get_textpage()
    return page.get_height() - text.get_charbox(text.get_text_range().index("5"))[3]


def test_a_cancelled_term_is_as_high_as_its_box(test_parameters: TestParameters):
    """Word strikes it out within its box: the slash of cancel.sty reaches a point beyond it, which set the fraction higher."""
    cancelled = _numerator_top(test_parameters.container, "\\\\frac{\\\\cancel{5}}{3}")
    plain = _numerator_top(test_parameters.container, "\\\\frac{5}{3}")
    assert cancelled == pytest.approx(plain, abs=0.05)


def _origin_distance(container, equation: str) -> float:
    """How far right of the origin of the 1 of a display equation the origin of its 0 stands, in the PDF."""
    native = f'[Para [Math DisplayMath "{equation}"]]'
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/matrix.native << 'EOF'\n{native}\nEOF"])
    command = f"{PANDOC_PATH} -s -f native --lua-filter={FILTER_PATH} -M docx-body-size=12 -M docx-body-baseline=14.65 --pdf-engine=tectonic -o /tmp/test/matrix.pdf /tmp/test/matrix.native"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    _, pdf = container.exec_run(["cat", "/tmp/test/matrix.pdf"])
    text = pdfium.PdfDocument(pdf)[0].get_textpage()
    characters = text.get_text_range()
    return text.get_charbox(characters.index("0"))[0] - text.get_charbox(characters.index("1"))[0]


def test_the_columns_of_a_matrix_stand_an_em_apart(test_parameters: TestParameters):
    """Word sets the columns of a matrix 12pt apart in a body of 12pt, TeX two \\arraycolsep, 10pt."""
    matrix = _origin_distance(test_parameters.container, "\\\\begin{matrix} 1 & 0 \\\\end{matrix}")
    assert matrix - _origin_distance(test_parameters.container, "10") == pytest.approx(12, abs=0.05)


def _equation_baseline(container, equation: str, *metadata: str) -> float:
    """Where the bottom of the = of a display equation below a line of text stands, from the top of the page."""
    native = f'[Para [Str "Text"], Para [Math DisplayMath "{equation}"]]'
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/limit.native << 'EOF'\n{native}\nEOF"])
    options = " ".join(f"-M {item}" for item in ("docx-body-size=12", "docx-body-baseline=14.65", *metadata))
    command = f"{PANDOC_PATH} -s -f native --lua-filter={FILTER_PATH} {options} --pdf-engine=tectonic -o /tmp/test/limit.pdf /tmp/test/limit.native"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    _, pdf = container.exec_run(["cat", "/tmp/test/limit.pdf"])
    page = pdfium.PdfDocument(pdf)[0]
    text = page.get_textpage()
    return page.get_height() - text.get_charbox(text.get_text_range().index("="))[1]


def test_a_limit_above_an_equation_takes_no_gap_beyond_it(test_parameters: TestParameters):
    """Word sets a limit right above what it stands on: TeX put the gap of a stack above it, 1.44pt in Latin Modern Math."""
    plain = _equation_baseline(test_parameters.container, "A\\\\overset{\\\\frac{a}{b}}{=}B")
    word = _equation_baseline(test_parameters.container, "A\\\\overset{\\\\frac{a}{b}}{=}B", "docx-math-limit-space=0.12")
    assert plain - word == pytest.approx(1.44, abs=0.05)


def test_a_brace_below_is_a_limit_of_the_size_of_a_script(test_parameters: TestParameters):
    """Word sets a lower limit of ⏟ alone as a character, 3.9pt wide under a+a+⋯+a: TeX stretched it across."""
    native = '[Para [Math DisplayMath "\\\\underbrace{a + a + \\\\cdots + a}"]]'
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/brace.native << 'EOF'\n{native}\nEOF"])
    command = f"{PANDOC_PATH} -s -f native --lua-filter={FILTER_PATH} -M docx-body-size=12 -M docx-body-baseline=14.65 -M docx-math-limit-space=0.12 --pdf-engine=tectonic -o /tmp/test/brace.pdf /tmp/test/brace.native"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    _, pdf = container.exec_run(["cat", "/tmp/test/brace.pdf"])
    text = pdfium.PdfDocument(pdf)[0].get_textpage()
    characters = text.get_text_range()
    braces = [text.get_charbox(index) for index, character in enumerate(characters) if character == "⏟"]
    assert len(braces) == 1, characters
    assert braces[0][2] - braces[0][0] < 6


def test_a_label_below_a_brace_stands_below_the_brace(test_parameters: TestParameters):
    """pandoc reads the label of a brace of Word as a lower limit around it, which stays below the brace."""
    native = '[Para [Math DisplayMath "\\\\underset{n}{\\\\underbrace{a + a + \\\\cdots + a}}"]]'
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/label.native << 'EOF'\n{native}\nEOF"])
    command = f"{PANDOC_PATH} -s -f native --lua-filter={FILTER_PATH} -M docx-body-size=12 -M docx-body-baseline=14.65 -M docx-math-limit-space=0.12 --pdf-engine=tectonic -o /tmp/test/label.pdf /tmp/test/label.native"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    _, pdf = container.exec_run(["cat", "/tmp/test/label.pdf"])
    text = pdfium.PdfDocument(pdf)[0].get_textpage()
    characters = text.get_text_range()
    brace, label = (text.get_charbox(characters.index(character)) for character in ("⏟", "\N{MATHEMATICAL ITALIC SMALL N}"))
    assert label[3] <= brace[1], characters
    assert brace[2] - brace[0] < 6


def _bracket_height(container, equation: str) -> float:
    """How high the PDF draws the opening bracket of a display equation, the first character of its page."""
    native = f'[Para [Math DisplayMath "{equation}"]]'
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/bracket.native << 'EOF'\n{native}\nEOF"])
    command = f"{PANDOC_PATH} -s -f native --lua-filter={FILTER_PATH} -M docx-body-size=12 -M docx-body-baseline=14.65 --pdf-engine=tectonic -o /tmp/test/bracket.pdf /tmp/test/bracket.native"
    exit_code, output = container.exec_run(["sh", "-c", command])
    assert exit_code == 0, output.decode()
    _, pdf = container.exec_run(["cat", "/tmp/test/bracket.pdf"])
    _, bottom, _, top = pdfium.PdfDocument(pdf)[0].get_textpage().get_charbox(0)
    return top - bottom


def test_a_growing_bracket_covers_less_of_its_content_than_in_tex(test_parameters: TestParameters):
    """Word sets (x²) in Latin Modern Math with the bracket of the text, 11.94pt high: TeX took the next size, 13.26pt."""
    assert _bracket_height(test_parameters.container, "\\\\left( x^{2} \\\\right)") < 12.5


def test_the_brackets_of_a_fraction_without_a_bar_grow_as_any_other(test_parameters: TestParameters):
    """Word sets n over k in brackets 21.49pt high in Latin Modern Math, as a over b: TeX's \\binom took 28.96pt."""
    binomial = _bracket_height(test_parameters.container, "\\\\binom{n}{k}")
    assert binomial == pytest.approx(_bracket_height(test_parameters.container, "\\\\left( \\\\frac{a}{b} \\\\right)"), abs=0.1)


def test_a_coloured_numbered_heading_colours_its_number_and_its_text_inside_its_line(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-style-Heading1=size:20,color:0F4761,number:{1:arabic},numberhang:21.6"))
    assert "\\titleformat{\\section}[hang]{\\normalfont\\fontsize{20bp}{24bp}\\selectfont\\filright}{\\color[HTML]{0F4761}\\makebox[21.6pt][l]{\\thesection}}{0pt}{\\color[HTML]{0F4761}}" in preamble


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


def test_the_fonts_of_the_body_load_first_in_every_run_of_tectonic(test_parameters: TestParameters):
    """tectonic keeps glyph boxes by font number across its runs; a table of contents shifts the numbers from the second on."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4"))
    assert "\\AtBeginDocument{\\setbox0\\hbox{\\normalfont x\\bfseries x\\itshape x\\normalfont\\itshape x}}" in preamble


def test_a_table_starts_below_the_descent_of_the_line_before_it(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4"))
    assert "\\AddToHook{env/longtable/before}{\\par\\ifdim\\prevdepth>-1000pt\\ifdim\\prevdepth<\\dp\\strutbox" in preamble


def test_a_body_without_its_descent_keeps_the_strut_of_latex(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65"))
    assert "\\strutbox" not in preamble
    assert "\\lineskip=0pt" not in preamble


def test_a_row_of_compact_has_half_its_space_above_its_line_and_half_below(test_parameters: TestParameters):
    """Word sets 1.8 above the line of a row of Compact: the stretched strut of array.sty is that high with an extra height of
    (11.25 + 1.8) / 1.5294 - 11.25 = -2.717."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-baseline=14.65", "docx-body-descent=3.4", "docx-table-cell-space=3.6"))
    assert "\\newcommand{\\pdcCompactCellExtra}{-2.717bp}" in preamble


def test_without_the_descent_a_row_of_compact_shares_its_space_by_the_stretch(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-baseline=14.65", "docx-table-cell-space=3.6"))
    assert "\\newcommand{\\pdcCompactCellStretch}{1.2457}" in preamble
    assert "\\newcommand{\\pdcCompactCellExtra}{0pt}" in preamble


def test_a_body_size_without_its_line_is_left_out(test_parameters: TestParameters):
    assert "\\renewcommand{\\normalsize}" not in _preamble(_latex(test_parameters.container, "docx-body-size=12"))


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


def test_the_pdf_breaks_the_page_where_the_docx_does(test_parameters: TestParameters):
    from docx.enum.text import WD_BREAK

    document = Document()
    document.add_paragraph("First page.")
    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    document.add_paragraph("Second page.")
    reader = PdfReader(io.BytesIO(_docx_to_pdf(test_parameters, document)))
    assert len(reader.pages) == 2
    assert "Second page." in reader.pages[1].extract_text()


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


def test_a_table_after_a_tight_list_gets_only_the_space_after_compact(test_parameters: TestParameters):
    """The items of a tight list are Compact paragraphs: Word puts their 1.8pt after above a table."""
    cell = 'Cell ("",[],[]) AlignDefault (RowSpan 1) (ColSpan 1) [Plain [Str "cell"]]'
    body = f'TableBody ("",[],[]) (RowHeadColumns 0) [] [Row ("",[],[]) [{cell}]]'
    table = f'Table ("",[],[]) (Caption Nothing []) [(AlignDefault,ColWidthDefault)] (TableHead ("",[],[]) []) [{body}] (TableFoot ("",[],[]) [])'
    native = f'[BulletList [[Plain [Str "item"]]], {table}]'
    spacing = "docx-paragraph-spacing=BodyText:Body Text:9:9;Compact:Compact:1.8:1.8;Normal:Normal:0:10"
    body = _latex(test_parameters.container, "docx-body-paragraph-skip=9", spacing, native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    assert body.index("\\end{itemize}") < body.index("\\vspace{-7.2bp}") < body.index("\\begin{longtable}")


def test_a_list_takes_the_indents_of_word_with_its_label_aligned_left(test_parameters: TestParameters):
    """Word sets the label at the start of the hanging indent: LaTeX right-aligns it against the text."""
    preamble = _preamble(_latex(test_parameters.container, "docx-list-indents=36:18,36:18,36:9,36:18"))
    assert "\\def\\pdcListLefti{36bp}\\def\\pdcListHangi{18bp}" in preamble
    assert "\\def\\pdcListLeftiii{36bp}\\def\\pdcListHangiii{9bp}" in preamble
    assert "\\leftmargin=\\csname pdcListLeft\\romannumeral\\@listdepth\\endcsname" in preamble, "a quote keeps the margin of LaTeX"
    assert "\\def\\makelabel##1{##1\\hfil}" in preamble
    assert "\\AddToHook{env/itemize/begin}{\\global\\let\\pdcNextList\\pdcWordLabel}" in preamble


def test_a_list_without_four_levels_of_indents_keeps_those_of_latex(test_parameters: TestParameters):
    assert "pdcListLeft" not in _preamble(_latex(test_parameters.container, "docx-list-indents=36:18"))


def test_a_paragraph_after_a_caption_gets_the_larger_of_their_spaces(test_parameters: TestParameters):
    """Word puts the 6pt after the caption before a Normal paragraph, not the paragraph skip of 9pt."""
    native = '[Div ("",[],[("custom-style","PandocCaption")]) [Para [Str "Figure"]], Para [Str "After"]]'
    body = _latex(test_parameters.container, *SPACING, "docx-style-Caption=size:12,before:0,after:6", native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    assert body.index("Figure") < body.index("\\vspace{-3bp}") < body.index("After")


def test_a_picture_alone_on_its_line_goes_down_by_the_descent(test_parameters: TestParameters):
    """Word sets the line of only a picture on its bottom and the next line right below: TeX keeps the descent between."""
    image = 'Image ("",[],[]) [] ("x.png","")'
    native = f'[Para [Str "Text", LineBreak, {image}, LineBreak, Str "After"], Para [Str "Icon", Space, {image}]]'
    body = _latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    assert "Text\\strut{}\\\\ {\\raisebox{-3.4bp}{\\pandocbounded{\\includegraphics[keepaspectratio]{x.png}}}}\\strut{}\\\\ After" in body
    assert body.count("raisebox") == 1, "a picture next to text stays on the baseline"


def test_a_picture_alone_in_a_link_goes_down_by_the_descent_as_well(test_parameters: TestParameters):
    """A picture of Polarion links to its file: alone on its line it stands as low as one without a link."""
    image = 'Image ("",[],[]) [] ("x.png","")'
    native = f'[Para [Link ("",[],[]) [{image}] ("x.png","")]]'
    body = _latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    assert "{\\raisebox{-3.4bp}{\\pandocbounded{\\includegraphics[keepaspectratio]{x.png}}}}" in body, body


def test_pictures_side_by_side_on_their_line_go_down_by_the_descent(test_parameters: TestParameters):
    """Word sets a line of pictures alone as it sets a line of one: no descent of the font below them."""
    image = 'Image ("",[],[]) [] ("x.png","")'
    native = f'[Para [{image}, Space, Link ("",[],[]) [{image}] ("x.png","")], Para [{image}, Space, Str "text"]]'
    body = _latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", native=native)
    body = " ".join(body[body.index("\\begin{document}") :].split())
    assert body.count("raisebox{-3.4bp}") == 2, body


def test_a_picture_taller_than_a_page_runs_past_the_bottom_margin(test_parameters: TestParameters):
    """Word sets it at the top of a page below the lines before it there: TeX would leave a page empty."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12"))
    assert "\\setlength{\\pdcPictureRoom}{\\dimexpr\\textheight-2\\baselineskip\\relax}" in preamble
    assert "[\\pdcPictureRoom][0pt]" in preamble
    assert "\\RenewDocumentCommand{\\includegraphics}{som}{\\pdcPagePicture{" in preamble


def test_a_table_goes_on_at_the_top_margin_of_the_next_page(test_parameters: TestParameters):
    """longtable sets an empty header at the top of the next page, which TeX would move down by the ascent."""
    preamble = _preamble(_latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4"))
    assert "\\AddToHook{env/longtable/begin}{\\topskip=0pt}" in preamble


def test_a_period_at_the_end_of_a_line_hangs_past_the_margin(test_parameters: TestParameters):
    """Word lets the period which ends a line stand past the right margin: only the word before it has to fit."""
    native = '[Para [Str "One.", Space, Str "Two", Space, RawInline (Format "latex") "{\\\\large ", Str "end.", RawInline (Format "latex") "}"]]'
    body = " ".join(_latex(test_parameters.container, "docx-body-size=12", "docx-body-descent=3.4", native=native).split())
    assert "One\\rlap{.}\\hskip\\dimexpr\\fontcharwd\\font`.+\\fontdimen2\\font\\relax Two" in body
    assert "end\\rlap{.}}" in body, "the period of the paragraph hangs inside the group of its run"


def test_the_punctuation_word_lets_hang_hangs_past_the_margin(test_parameters: TestParameters):
    """Probes in Word for Mac: a colon, a closing bracket, a percent sign and their like hang as the period does."""
    native = '[Para [Str "a)", Space, Str "50%", Space, Str "b-", Space, Str "c", LineBreak, Str "d\u201d", LineBreak, Str "handling):"]]'
    body = " ".join(_latex(test_parameters.container, "docx-body-size=12", "docx-body-descent=3.4", native=native).split())
    assert "a\\rlap{)}\\hskip\\dimexpr\\fontcharwd\\font`)+\\fontdimen2\\font\\relax 50" in body, body
    assert "50\\rlap{\\%}\\hskip\\dimexpr\\fontcharwd\\font`\\%+" in body, "a percent sign goes in escaped"
    assert "b- c" in body, "a hyphen does not hang"
    assert "d\\rlap{\u201d}\\strut{}\\\\" in body, "a mark before a line break hangs"
    assert "handling)\\rlap{:}" in body, "the mark which ends the paragraph hangs"


def test_an_equation_has_no_room_for_an_empty_limit(test_parameters: TestParameters):
    """Word sets no limit where a sum has an empty one: TeX keeps room for it."""
    body = _latex(test_parameters.container, "docx-body-size=12", "docx-body-baseline=14.65", native='[Para [Math DisplayMath "\\\\sum_{i}^{}P_{}"]]')
    assert "\\sum_{i}P" in body, body


def test_an_equation_without_the_size_of_the_body_keeps_the_fractions_of_latex(test_parameters: TestParameters):
    """\\pdcEquationFrac comes with the size of the body: without it an equation uses no command it lacks."""
    latex = _latex(test_parameters.container, native='[Para [Math DisplayMath "\\\\frac{a}{b}"]]')
    assert "pdcEquation" not in latex
    assert "\\frac{a}{b}" in latex


def test_every_line_a_break_ends_reaches_the_descent_of_the_font(test_parameters: TestParameters):
    """Word sets a picture on the next line below the descent of the font, whatever the letters before the break."""
    metadata = ("docx-body-size=12", "docx-body-baseline=14.64", "docx-body-ascent=11.28", "docx-body-descent=3.36")
    native = '[Para [Str "Not", Space, Str "resized:", LineBreak, Strong [Str "bold", LineBreak], Str "end"]]'
    body = " ".join(_latex(test_parameters.container, *metadata, native=native).split())
    assert "resized\\rlap{:}\\strut{}\\\\ \\textbf{bold\\strut{}\\\\ }end" in body, body


def test_a_page_ends_where_the_descent_of_its_last_line_reaches_the_margin(test_parameters: TestParameters):
    """Word keeps the whole last line of a page above the bottom margin: TeX counts its whole depth, also at a penalty."""
    metadata = ("docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4")
    preamble = _preamble(_latex(test_parameters.container, *metadata))
    assert "\\newcommand{\\pdcPageDepth}{\\global\\maxdepth=\\z@\\global\\@maxdepth=\\z@}" in preamble
    assert "\\AddToHook{begindocument/end}{\\pdcPageDepth}" in preamble


def test_a_paragraph_in_the_line_of_its_runs_gets_its_struts_inside(test_parameters: TestParameters):
    """docx_colors_to_latex opens the line of the font of the runs: the struts are of that font, so inside it."""
    metadata = ("docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4")
    native = '[Para [RawInline (Format "latex") "{\\\\pdcParLine{12.8bp}{10.2bp}{2.6bp}", Str "text", RawInline (Format "latex") "\\\\par}"]]'
    body = " ".join(_latex(test_parameters.container, *metadata, native=native).split())
    assert "{\\pdcParLine{12.8bp}{10.2bp}{2.6bp}\\strut{}text\\strut{}\\par}" in body, body


def test_a_character_of_a_shared_glyph_carries_its_own_text(test_parameters: TestParameters):
    """A reader takes the ( of Segoe UI for U+FD3E, which the font maps to the same glyph: the PDF says it is a (."""
    body = _latex(test_parameters.container, "docx-actual-text=0028", native='[Para [Str "(a)"]]')
    body = body[body.index("\\begin{document}") :]
    assert "\\texorpdfstring{\\special{pdf:content /Span<</ActualText<FEFF0028>>>BDC}(\\special{pdf:content EMC}}{(}a\\rlap{)}" in body, body


def test_a_hanging_mark_of_a_shared_glyph_carries_its_own_text(test_parameters: TestParameters):
    """The ) of Segoe UI shares its glyph with U+FD3F: past the margin it keeps its own text as well."""
    body = _latex(test_parameters.container, "docx-actual-text=0029", native='[Para [Str "a)", Space, Str "b)"]]')
    body = " ".join(body[body.index("\\begin{document}") :].split())
    shared = "\\rlap{\\special{pdf:content /Span<</ActualText<FEFF0029>>>BDC})\\special{pdf:content EMC}}"
    assert f"a{shared}\\hskip" in body, body
    assert f"b{shared}" in body, "the mark which ends the paragraph too"


def test_the_hyphen_of_a_font_which_shares_its_glyph_copies_out_as_a_hyphen(test_parameters: TestParameters):
    """Noto Sans CJK maps - and U+2011 to one glyph, which xdvipdfmx writes as U+2011, NON-BREAKING HYPHEN."""
    document = Document()
    document.styles["Normal"].font.name = "Noto Sans CJK SC"
    document.add_paragraph("e-mail")
    docx = io.BytesIO()
    document.save(docx)
    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/pdf", files={"source": ("in.docx", docx.getvalue(), mime)})
    assert response.status_code == 200, response.text
    assert "e-mail" in pdfium.PdfDocument(response.content)[0].get_textpage().get_text_range()


def test_a_line_which_fits_but_for_its_closing_colon_stays_one_line(test_parameters: TestParameters):
    """The line of test_formulas_conversion is 468.475pt wide in Liberation Sans, the text 468pt: Word hangs the colon.

    TeX moved its last word to a line of its own, and everything below the paragraph down by a line.
    """
    document = Document()
    # As the documents of the corpus: Word hangs punctuation with Simplified Chinese as the East Asian language
    defaults = document.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    lang = defaults.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        defaults.append(lang)
    lang.set(qn("w:eastAsia"), "zh-CN")
    section = document.sections[0]
    section.left_margin = section.right_margin = Inches(1)
    section.page_width = Inches(8.5)
    document.styles["Normal"].font.name = "Liberation Sans"
    document.styles["Normal"].font.size = Pt(12)
    document.add_paragraph("Plain-TeX matrix macro + \\cr, multi-line (\\pmatrix -> pmatrix env, \\cr -> \\\\, <br> handling):")
    text = pdfium.PdfDocument(_docx_to_pdf(test_parameters, document))[0].get_textpage().get_text_range()
    assert "<br> handling):" in text, text


def test_no_punctuation_hangs_where_word_lets_none(test_parameters: TestParameters):
    native = '[Para [Str "a)", Space, Str "end."]]'
    body = " ".join(_latex(test_parameters.container, "docx-body-size=12", "docx-hanging-punctuation=false", native=native).split())
    assert "a) end." in body, body
    assert "rlap" not in body


def test_a_space_which_starts_the_line_after_a_break_stays(test_parameters: TestParameters):
    """Word keeps the space after a line break, TeX drops it with the break: the attribute line of a work item starts a space in."""
    native = '[Para [Str "Title", LineBreak, Str "\\8203\\8203", Space, Str "[Draft]", LineBreak, Str "next", LineBreak, Space, Str "spaced"]]'
    body = " ".join(_latex(test_parameters.container, native=native).split())
    assert "\\\\ \\hbox{}\\hspace{0pt}\\hspace{0pt} {[}Draft" in body, body
    assert "\\\\ next" in body, "a line which starts with text gets nothing"
    assert "\\\\ \\hbox{} spaced" in body


def test_an_equation_in_a_cell_sets_the_line_of_its_font_lower(test_parameters: TestParameters):
    """Word sets the line of the font of an equation in a cell, whose paragraph has no space before it, 0.135 times the size lower.

    In the body the space before its paragraph takes it. test_test_case_tables: the equation of a cell sat 2pt high.
    """
    metadata = ("docx-body-size=12", "docx-body-baseline=14.65", "docx-body-descent=3.4", "docx-math-top=0.9502", "docx-math-descent=0.2222")
    preamble = _preamble(_latex(test_parameters.container, *metadata, native='[Para [Math DisplayMath "x"]]'))
    assert "\\def\\pdcDisplayLine#1{" in preamble
    body = preamble[preamble.index("\\def\\pdcDisplayLine#1{") :].split("\n")[0]
    cell = preamble[preamble.index("\\def\\pdcCellDisplayLine#1{") :].split("\n")[0]
    end = preamble[preamble.index("\\def\\pdcDisplayLineEnd#1{") :].split("\n")[0]
    assert "\\rule[-2.6664bp]{0pt}{14.0688bp}" in body, body
    assert "\\rule[-2.6664bp]{0pt}{15.6888bp}" in cell, cell
    assert "\\rule[-2.6664bp]{0pt}{15.6888bp}" in end, end
    # In a cell the line may hang from the top of the row, in the body it does not
    assert "\\pdcCellHang{\\rule{0pt}" in cell and "\\pdcCellHang{\\rule{0pt}" in end
    assert "\\pdcCellHang" not in body
