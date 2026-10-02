"""Integration tests for the Figure handling of ``filters/docx_caption_labels_to_latex.lua``.

Runs the real ``pandoc`` binary inside the pandoc-service container with the
filter on a native AST that mimics what the docx reader produces when a
Caption-styled paragraph directly precedes a paragraph holding only an image:
it pairs the two into a Figure. Word shows them as two ordinary paragraphs, so
the PDF must too, without a centered float or a "Figure N:" label of its own.

The service test at the end covers the real reader: it exports Polarion-shaped
HTML to DOCX and converts that DOCX to LaTeX, both through the service.
"""

from __future__ import annotations

import base64
import struct
import zlib

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


def _png() -> str:
    """A 1x1 PNG as a data URI."""

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b"")
    return "data:image/png;base64," + base64.b64encode(png).decode()


def _polarion_figure(number: int, title: str) -> str:
    """An image paragraph and the Polarion caption below it."""
    return f'<p><img src="{_png()}" style="width: 100px; height: 50px;"/></p><p class="polarion-rte-caption-paragraph">Figure <span class="polarion-rte-caption" data-sequence="Figure">{number}</span> {title}</p>'


def test_service_keeps_figure_captions_below_their_images(test_parameters: TestParameters):
    """The caption of the first image precedes the second image, and the reader must not pair the two."""
    html = f"<html><body><p>Figures</p>{_polarion_figure(1, 'First Picture')}{_polarion_figure(2, 'Second Picture')}</body></html>"
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/html/to/docx", data=html)
    assert response.status_code == 200, response.text

    files = {"source": ("in.docx", response.content, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
    response = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/latex", files=files)
    assert response.status_code == 200, response.text
    latex = response.text

    assert "\\begin{figure}" not in latex, latex
    assert "\\caption" not in latex, latex
    first_image = latex.index("\\includegraphics")
    second_image = latex.index("\\includegraphics", first_image + 1)
    assert first_image < latex.index("First Picture") < second_image < latex.index("Second Picture"), latex
