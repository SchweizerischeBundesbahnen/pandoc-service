"""Tests for the filters ``docx_math_colors_to_latex.lua`` and ``docx_math_styles_to_latex.lua``: the markers of math runs become LaTeX."""

from __future__ import annotations

import ctypes

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c
import pytest

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_math_colors_to_latex.lua"


def _md_to_latex(test_parameters: TestParameters, md: str) -> str:
    container = test_parameters.container
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.md << 'HEREDOC_EOF'\n{md}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/in.md"])
    assert exit_code == 0, output.decode()
    return " ".join(output.decode("utf-8").split())


def _marked(color: str, text: str) -> str:
    return f"PMCzzz{color}zzz{text}zzzPMCENDzzz"


def test_colored_letter_becomes_color_group(test_parameters: TestParameters):
    latex = _md_to_latex(test_parameters, f"${_marked('FF0000', 'a')}$")
    assert "{\\color[HTML]{FF0000} a}" in latex, latex


def test_colored_operators_keep_their_spacing(test_parameters: TestParameters):
    """A group is an ordinary atom, so a colored + in one is set as tight as a letter. Word spaces it as an operator."""
    md = f"${_marked('FF0000', 'a')}{_marked('FF0000', ' + ')}{_marked('0000FF', 'b')}{_marked('0000FF', ' = ')}{_marked('0000FF', 'c')}$"
    latex = _md_to_latex(test_parameters, md)
    assert "\\mathbin{\\color[HTML]{FF0000}+}" in latex, latex
    assert "\\mathrel{\\color[HTML]{0000FF}=}" in latex, latex


STYLES_FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_math_styles_to_latex.lua"


def test_style_markers_become_upright_and_bold_italic(test_parameters: TestParameters):
    """texmath sets an upright d of "dx" italic and a bold italic v upright. Word sets them as marked."""
    container = test_parameters.container
    md = "$PMSzzzpzzzdzzzPMSENDzzzx + PMSzzzbizzzvzzzPMSENDzzz$"
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && printf '%s' '{md}' > /tmp/test/style.md"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={STYLES_FILTER_PATH} /tmp/test/style.md"])
    assert exit_code == 0, output.decode()
    assert "\\symup{d}x + \\symbfit{v}" in output.decode(), output.decode()


def test_alphabets_of_texmath_come_from_the_math_font(test_parameters: TestParameters):
    """unicode-math sets \\mathtt and its kin in the text font. Word sets them in the math font."""
    container = test_parameters.container
    md = r"$\mathbf{v} \cdot \mathit{t} + \mathsf{S} + \mathtt{code} + \mathrm{d} + \mathbb{R} + \mathcal{L}$"
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && printf '%s' '{md}' > /tmp/test/alphabets.md"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={STYLES_FILTER_PATH} /tmp/test/alphabets.md"])
    assert exit_code == 0, output.decode()
    assert r"\symbf{v} \cdot \symit{t} + \symsf{S} + \symtt{code} + \symup{d} + \mathbb{R} + \mathcal{L}" in output.decode(), output.decode()


def _script_rise(test_parameters: TestParameters, math: str) -> float:
    """How far above the baseline of the letter b of the equation given its superscript 2 sits, in the PDF."""
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/rise.md << 'HEREDOC_EOF'\n$${math}$$\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"cd /tmp/test && {PANDOC_PATH} -f markdown --lua-filter={FILTER_PATH} --pdf-engine=tectonic -V mathfont='DejaVu Math TeX Gyre' rise.md -o rise.pdf"])
    assert exit_code == 0, output.decode()
    exit_code, pdf = container.exec_run(["cat", "/tmp/test/rise.pdf"])
    textpage = pdfium.PdfDocument(pdf)[0].get_textpage()
    # The letter b comes first on the page, a math italic outside the 16 bits pdfium gives a character; the 2 next
    superscript = next(index for index in range(textpage.count_chars()) if pdfium_c.FPDFText_GetUnicode(textpage.raw, index) == ord("2"))
    baselines = []
    for index in (0, superscript):
        x, y = ctypes.c_double(), ctypes.c_double()
        pdfium_c.FPDFText_GetCharOrigin(textpage.raw, index, ctypes.byref(x), ctypes.byref(y))
        baselines.append(y.value)
    return baselines[1] - baselines[0]


def test_a_colored_base_raises_its_superscript_as_a_letter_does(test_parameters: TestParameters):
    """With an OpenType math font TeX raises the script of a group from its top, of a letter less: Word sets a colored b² as a black one."""
    colored = _script_rise(test_parameters, f"{{{_marked('FF0000', 'b')}}}^{{{_marked('FF0000', '2')}}}")
    plain = _script_rise(test_parameters, "b^{2}")
    assert abs(colored - plain) < 0.05, (colored, plain)


def test_a_script_of_black_content_stays_out_of_a_colored_base(test_parameters: TestParameters):
    """The color of the base would reach a script inside its group."""
    latex = _md_to_latex(test_parameters, f"${{{_marked('0000FF', 'c')}}}^{{2}}$")
    assert "{{\\color[HTML]{0000FF} c}}^{2}" in latex, latex


def test_the_colored_limits_of_an_integral_stay_apart(test_parameters: TestParameters):
    """A colored group after _ is a script of the integral, not the base of the script after it."""
    latex = _md_to_latex(test_parameters, f"$\\int_{{{_marked('008000', '0')}}}^{{{_marked('008000', '1')}}}$")
    assert "\\int_{{\\color[HTML]{008000} 0}}^{{\\color[HTML]{008000} 1}}" in latex, latex


def test_both_scripts_of_a_colored_base_move_into_its_group(test_parameters: TestParameters):
    latex = _md_to_latex(test_parameters, f"${{{_marked('00FF00', 'x')}}}_{{{_marked('00FF00', 'i')}}}^{{{_marked('00FF00', '2')}}}$")
    assert "{\\color[HTML]{00FF00} x_{{\\color[HTML]{00FF00} i}}^{{\\color[HTML]{00FF00} 2}}}" in latex, latex


@pytest.mark.parametrize("command", ["\\sqrt", "\\sqrt[3]", "\\frac{a}", "\\symbf"])
def test_a_script_after_the_colored_argument_of_a_command_stays_out_of_it(test_parameters: TestParameters, command: str):
    """The power of a root of a colored x belongs to the root: inside the argument it would be the power of x."""
    latex = _md_to_latex(test_parameters, f"${command}{{{_marked('FF0000', 'x')}}}^{{{_marked('FF0000', '2')}}}$")
    assert f"{command}{{{{\\color[HTML]{{FF0000}} x}}}}^{{{{\\color[HTML]{{FF0000}} 2}}}}" in latex, latex


def test_a_colored_base_after_a_colored_letter_takes_its_script(test_parameters: TestParameters):
    """E = mc²: the group of the m before the c is no argument."""
    latex = _md_to_latex(test_parameters, f"${_marked('0000FF', 'm')}{{{_marked('0000FF', 'c')}}}^{{{_marked('0000FF', '2')}}}$")
    assert "{\\color[HTML]{0000FF} m}{\\color[HTML]{0000FF} c^{{\\color[HTML]{0000FF} 2}}}" in latex, latex
