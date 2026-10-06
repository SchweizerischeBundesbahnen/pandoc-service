"""Integration tests for ``filters/docx_image_layout_to_latex.lua``.

The filter runs on the real ``pandoc`` binary inside the pandoc-service container: first on
a native AST holding the marker ``app/docx_image_layout_pre_process.py`` writes into an image
title, then on the whole round trip, where an HTML icon converted to DOCX keeps its offset
and gap in the LaTeX the same DOCX converts to.
"""

from __future__ import annotations

import base64
import struct
import zlib

import pytest
from docker.models.containers import Container

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
LAYOUT = "/usr/local/share/pandoc/filters/docx_image_layout_to_latex.lua"
GRAPHIC = "\\pandocbounded{\\includegraphics[keepaspectratio]{pic.png}}"


def _native_to_latex(container: Container, native: str) -> str:
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(["sh", "-c", f"{PANDOC_PATH} -f native -t latex /tmp/test/in.native --lua-filter={LAYOUT}"])
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return " ".join(output.decode("utf-8").split())


def _image(title: str) -> str:
    return f'[ Para [ Str "a" , Image ( "" , [] , [] ) [] ( "pic.png" , "{title}" ) , Str "b" ] ]'


@pytest.mark.parametrize(
    ("title", "expected"),
    [
        pytest.param("{{PICLAYOUT:-6|0|19050}}", "a\\raisebox{-3pt}{" + GRAPHIC + "}\\hspace{1.5pt}b", id="shift-and-gap"),
        pytest.param("{{PICLAYOUT:4|12700|0}}", "a\\hspace{1pt}\\raisebox{2pt}{" + GRAPHIC + "}b", id="raise-and-left-gap"),
        pytest.param("{{PICLAYOUT:0|0|0}}", "a" + GRAPHIC + "b", id="nothing-to-do"),
        # %g would write these as 1e+04 and 7.874e-05, which TeX cannot read
        pytest.param("{{PICLAYOUT:-40000|1|127000000}}", "a\\hspace{0pt}\\raisebox{-16000pt}{" + GRAPHIC + "}\\hspace{10000pt}b", id="extreme-values"),
    ],
)
def test_marker_sets_the_picture_in_raisebox_and_hspace(test_parameters: TestParameters, title: str, expected: str):
    assert expected in _native_to_latex(test_parameters.container, _image(title))


def test_marker_gives_the_title_back(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, _image("{{PICLAYOUT:-6|0|0}}Draft"))

    assert "PICLAYOUT" not in latex
    assert "\\raisebox{-3pt}" in latex


def test_image_without_a_marker_is_left_alone(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, _image("Draft"))

    assert "\\raisebox" not in latex
    assert "\\hspace" not in latex


def _png_data_uri(width: int, height: int) -> str:
    raw = b"".join(b"\x00" + b"\x80\x80\x80" * width for _ in range(height))

    def chunk(tag: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + tag + payload + struct.pack(">I", zlib.crc32(tag + payload))

    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    return "data:image/png;base64," + base64.b64encode(png).decode()


def test_icon_layout_survives_from_html_through_docx_into_latex(test_parameters: TestParameters):
    """An enum icon converted to DOCX and that DOCX to LaTeX, as the visual tests of docx-exporter do."""
    icon = f'<img style="vertical-align:bottom;margin-right:2px;" src="{_png_data_uri(16, 16)}"/>'
    html = f'<p><span style="font-weight:bold;"><span class="polarion-JSEnumOption">{icon}Draft</span></span></p>'

    docx = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/html/to/docx", data=html)
    assert docx.status_code == 200, docx.text
    latex = test_parameters.request_session.post(f"{test_parameters.base_url}/convert/docx/to/latex", files={"source": ("file.docx", docx.content)})
    assert latex.status_code == 200, latex.text

    text = " ".join(latex.text.split())
    # A quarter of pandoc's default 12pt body text down, the 2px margin after
    assert "\\raisebox{-3pt}{" in text, text
    assert "\\hspace{1.5pt}" in text, text
    assert "PICLAYOUT" not in text
