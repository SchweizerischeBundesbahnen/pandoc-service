"""Integration tests for ``filters/docx_paragraphs_to_latex.lua``.

Runs pandoc inside the pandoc-service container with the filter loaded and
asserts the LaTeX output for a ``Div`` carrying the synthetic
``custom-style="PandocPara__..."`` attribute the ``docx_paragraph_pre_process``
preprocessor produces. The AST shape is built via a markdown fenced div
(``::: {custom-style="..."}``), which pandoc turns into exactly that ``Div``
node — so the test is pandoc-only and needs no DOCX fixture or tectonic.
"""

from __future__ import annotations

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_paragraphs_to_latex.lua"


def _md_to_latex(container, md: str) -> str:
    """Convert a markdown snippet to LaTeX through the filter; collapse
    whitespace so assertions don't depend on pandoc's line wrapping."""
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.md << 'HEREDOC_EOF'\n{md}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={FILTER_PATH} /tmp/test/in.md"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return " ".join(output.decode("utf-8").split())


def _div(style: str, text: str = "content") -> str:
    return f'::: {{custom-style="{style}"}}\n{text}\n:::\n'


def test_center_alignment_emits_centering(test_parameters: TestParameters):
    flat = _md_to_latex(test_parameters.container, _div("PandocPara__ALIGN_center"))
    assert "\\centering" in flat, flat
    assert "\\par}" in flat, flat


def test_right_alignment_emits_raggedleft(test_parameters: TestParameters):
    assert "\\raggedleft" in _md_to_latex(test_parameters.container, _div("PandocPara__ALIGN_right"))


def test_left_alignment_emits_raggedright(test_parameters: TestParameters):
    assert "\\raggedright" in _md_to_latex(test_parameters.container, _div("PandocPara__ALIGN_left"))


def test_indent_emits_leftskip_in_points(test_parameters: TestParameters):
    """600 twips = 30pt, 1200 twips = 60pt — distinct, not collapsed."""
    assert "\\leftskip=30.00pt" in _md_to_latex(test_parameters.container, _div("PandocPara__IND_600"))
    assert "\\leftskip=60.00pt" in _md_to_latex(test_parameters.container, _div("PandocPara__IND_1200"))


def test_alignment_and_indent_combined(test_parameters: TestParameters):
    flat = _md_to_latex(test_parameters.container, _div("PandocPara__ALIGN_right__IND_600"))
    assert "\\leftskip=30.00pt" in flat and "\\raggedleft" in flat, flat
    # Balanced group: one opening brace before content, \par} after.
    assert flat.count("\\par}") == 1, flat


def test_invalid_indent_value_is_dropped(test_parameters: TestParameters):
    """A non-integer / out-of-range twips value must not reach \\leftskip."""
    flat = _md_to_latex(test_parameters.container, _div("PandocPara__IND_99999999"))  # exceeds the 31680 cap
    assert "\\leftskip" not in flat, flat


def test_non_pandocpara_div_is_left_alone(test_parameters: TestParameters):
    flat = _md_to_latex(test_parameters.container, _div("SomeOtherStyle"))
    assert "\\centering" not in flat
    assert "\\raggedleft" not in flat
    assert "\\leftskip" not in flat
    assert "content" in flat


SHOW_STYLE = 'function Div(div) if div.attributes["pdc-style"] then return {pandoc.RawBlock("latex", "STYLE=" .. div.attributes["pdc-style"]), table.unpack(div.content)} end end'


def test_the_style_the_paragraph_had_stays_on_its_div(test_parameters: TestParameters):
    """docx_document_styles_to_latex puts the space of that style around the paragraph."""
    container = test_parameters.container
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/show.lua << 'HEREDOC_EOF'\n{SHOW_STYLE}\nHEREDOC_EOF"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.md << 'HEREDOC_EOF'\n{_div('PandocPara__ALIGN_center__STYLE_BodyText')}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f markdown -t latex --lua-filter={FILTER_PATH} --lua-filter=/tmp/test/show.lua /tmp/test/in.md"])
    assert exit_code == 0, output.decode()
    assert "STYLE=BodyText" in output.decode()
    assert "\\centering" in output.decode()
