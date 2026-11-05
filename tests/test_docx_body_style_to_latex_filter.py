"""Integration tests for ``filters/docx_body_style_to_latex.lua`` and app/docx_body_style.py.

The filter tests run pandoc inside the pandoc-service container with the
metadata app/docx_body_style.py passes, and read the preamble of the
standalone LaTeX. The end-to-end tests convert a DOCX through the service.
"""

from __future__ import annotations

import io

from docx import Document
from docx.shared import Pt
from pypdf import PdfReader

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_body_style_to_latex.lua"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
HEADER_FILTER = "/tmp/test/header.lua"
# A filter before ours which adds header-includes, as the other docx filters do
HEADER_SCRIPT = r'function Meta(m) m["header-includes"] = pandoc.MetaList({pandoc.MetaBlocks({pandoc.RawBlock("latex", "\\newcommand{\\other}{}")})}) return m end'


def _latex(container, *metadata: str, other_preamble: bool = False) -> str:
    """The standalone LaTeX of a one-line document, with the metadata given."""
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test && printf 'text\\n' > /tmp/test/in.md"])
    filters = ""
    if other_preamble:
        container.exec_run(["sh", "-c", f"cat > {HEADER_FILTER} << 'EOF'\n{HEADER_SCRIPT}\nEOF"])
        filters = f"--lua-filter={HEADER_FILTER} "
    options = " ".join(f"-M {item}" for item in metadata)
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -s -f markdown -t latex {filters}--lua-filter={FILTER_PATH} {options} /tmp/test/in.md"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode("utf-8")


def _preamble(latex: str) -> str:
    return latex[: latex.index("\\begin{document}")]


def test_word_layout_goes_into_the_preamble(test_parameters: TestParameters):
    preamble = _preamble(_latex(test_parameters.container, "docx-body-paragraph-skip=9", "docx-body-justified=false", "docx-body-hyphenated=false"))
    assert "\\setlength{\\parskip}{9pt}" in preamble
    assert "\\AtBeginDocument{\\raggedright}" in preamble
    assert "\\hyphenpenalty=10000" in preamble
    assert "\\exhyphenpenalty=10000" in preamble


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
