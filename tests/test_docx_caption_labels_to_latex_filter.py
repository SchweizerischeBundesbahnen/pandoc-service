"""Integration tests for the Figure handling of ``filters/docx_caption_labels_to_latex.lua``.

Runs the real ``pandoc`` binary inside the pandoc-service container with the
filter on a native AST that mimics what the docx reader produces when a
Caption-styled paragraph directly precedes a paragraph holding only an image:
it pairs the two into a Figure. Word shows them as two ordinary paragraphs, so
the PDF must too, without a centered float or a "Figure N:" label of its own.
"""

from __future__ import annotations

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_caption_labels_to_latex.lua"

_IMAGE = 'Image ( "" , [] , [] ) [] ( "media/picture.png" , "" )'


def _native_to_latex(container, native: str) -> str:
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/in.native"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode("utf-8")


def _figure(caption: str) -> str:
    return f'[ Figure ( "" , [] , [ ( "custom-style" , "Body Text" ) ] )\n    (Caption Nothing [{caption}])\n    [ Plain [ {_IMAGE} ] ] ]\n'


_CAPTION = 'Div ( "" , [] , [ ( "custom-style" , "Caption" ) ] ) [ Para [ Str "Figure" , Space , Str "1" , Space , Str "First" ] ]'


def test_captioned_figure_becomes_caption_then_image(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, _figure(_CAPTION))
    assert "\\begin{figure}" not in latex, latex
    assert "\\caption" not in latex, latex
    assert latex.index("Figure 1 First") < latex.index("\\includegraphics"), latex


def test_figure_without_caption_is_untouched(test_parameters: TestParameters):
    """Only a caption the reader paired with the image is taken apart."""
    latex = _native_to_latex(test_parameters.container, _figure(""))
    assert "\\begin{figure}" in latex, latex
