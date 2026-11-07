"""Tests for ``filters/docx/docx_math_colors_to_latex.lua``: the color markers of math runs become color groups."""

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
