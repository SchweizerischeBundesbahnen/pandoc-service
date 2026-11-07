"""Tests for the filters ``docx_math_colors_to_latex.lua`` and ``docx_math_styles_to_latex.lua``: the markers of math runs become LaTeX."""

from __future__ import annotations

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
