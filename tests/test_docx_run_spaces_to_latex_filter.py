"""Integration tests for ``filters/docx/docx_run_spaces_to_latex.lua``.

app/docx_run_spaces_pre_process.py writes a space at the edge of a formatted
run as U+E030, which pandoc's DOCX reader keeps inside the formatting. The
filter turns it back into a space where it stands.
"""

from __future__ import annotations

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx/docx_run_spaces_to_latex.lua"


def _latex(container, native: str) -> str:
    container.exec_run(["sh", "-c", f"mkdir -p /tmp/test && cat > /tmp/test/spaces.native << 'EOF'\n{native}\nEOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/spaces.native"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode().strip()


def test_a_space_of_a_struck_run_stays_struck(test_parameters: TestParameters):
    latex = _latex(test_parameters.container, '[Para [Str "a", Strikeout [Str "\\57392mo", Space, Str "du\\57392\\57392"], Str "b"]]')
    assert latex == "a\\st{ mo du }b", latex


def test_text_without_the_sentinel_stays_as_it_is(test_parameters: TestParameters):
    assert _latex(test_parameters.container, '[Para [Str "a", Space, Emph [Str "b"]]]') == "a \\emph{b}"


def test_a_space_of_a_code_run_stays_a_space(test_parameters: TestParameters):
    """pandoc reads a run of Verbatim Char as code, from its text as it is."""
    latex = _latex(test_parameters.container, '[Para [Str "a", Code ("",[],[]) "x\\57392", Str "b"], CodeBlock ("",[],[]) "\\57392y"]')
    assert "\ue030" not in latex, latex
    assert "\\texttt{x\\ }b" in latex or "\\texttt{x }b" in latex, latex
    assert "\\begin{verbatim}\n y\n\\end{verbatim}" in latex, latex
