"""Unit tests for ``app.docx_image_layout_pre_process``.

The preprocessor puts the ``<w:position>`` of a picture's run and the left and right
``<wp:effectExtent>`` of its inline in front of the picture's ``<wp:docPr title>``, the one
picture property pandoc's DOCX reader keeps. ``filters/docx_image_layout_to_latex.lua``
reads it back and is exercised with pandoc in ``test_docx_image_layout_to_latex_filter.py``.
"""

from __future__ import annotations

import io
import zipfile
from xml.etree import ElementTree as ET

import pytest

from app import docx_image_layout_pre_process, docx_latex_pre_process

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _document(runs: str) -> bytes:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W_NS}" xmlns:wp="{WP_NS}"><w:body><w:p>{runs}</w:p></w:body></w:document>'.encode()


def _picture_run(*, position: str | None = None, effect_extent: tuple[str, str] | None = None, title: str | None = None) -> str:
    r_pr = f'<w:rPr><w:position w:val="{position}"/></w:rPr>' if position is not None else ""
    extent = f'<wp:effectExtent l="{effect_extent[0]}" t="0" r="{effect_extent[1]}" b="0"/>' if effect_extent else ""
    title_attr = f' title="{title}"' if title is not None else ""
    return f'<w:r>{r_pr}<w:drawing><wp:inline><wp:extent cx="152400" cy="152400"/>{extent}<wp:docPr id="1" name="Picture"{title_attr}/></wp:inline></w:drawing></w:r>'


def _titles(xml: bytes) -> list[str | None]:
    return [doc_pr.get("title") for doc_pr in ET.fromstring(xml).iter(f"{{{WP_NS}}}docPr")]


@pytest.mark.parametrize(
    ("run", "expected"),
    [
        pytest.param(_picture_run(position="-6", effect_extent=("0", "19050")), "{{PICLAYOUT:-6|0|19050}}", id="shift-and-gap"),
        pytest.param(_picture_run(position="4"), "{{PICLAYOUT:4|0|0}}", id="raise-only"),
        pytest.param(_picture_run(effect_extent=("12700", "0"), title="Draft"), "{{PICLAYOUT:0|12700|0}}Draft", id="keeps-the-title"),
        pytest.param(_picture_run(position="x", effect_extent=("-5", "19050")), "{{PICLAYOUT:0|0|19050}}", id="ignores-invalid-values"),
    ],
)
def test_rewrite_part_marks_a_shifted_or_spaced_picture(run: str, expected: str):
    xml, changed = docx_image_layout_pre_process.rewrite_part(_document(run))

    assert changed
    assert _titles(xml) == [expected]


@pytest.mark.parametrize(
    "run",
    [
        pytest.param(_picture_run(), id="plain-picture"),
        pytest.param(_picture_run(position="0", effect_extent=("0", "0"), title="Draft"), id="zero-values"),
        pytest.param(_picture_run(position="-6", title="{{PICLAYOUT:-6|0|0}}"), id="already-marked"),
        pytest.param('<w:r><w:rPr><w:position w:val="-6"/></w:rPr><w:t>text</w:t></w:r>', id="text-run"),
    ],
)
def test_rewrite_part_leaves_other_runs_alone(run: str):
    source = _document(run)

    xml, changed = docx_image_layout_pre_process.rewrite_part(source)

    assert not changed
    assert xml == source


def test_rewrite_part_skips_unparseable_xml():
    assert docx_image_layout_pre_process.rewrite_part(b"<w:document") == (b"<w:document", False)


def test_docx_latex_preprocess_runs_the_image_layout_rewrite():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr("word/document.xml", _document(_picture_run(position="-6")))

    result = docx_latex_pre_process.preprocess(buffer.getvalue())

    with zipfile.ZipFile(io.BytesIO(result)) as package:
        assert _titles(package.read("word/document.xml")) == ["{{PICLAYOUT:-6|0|0}}"]
