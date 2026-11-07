"""Unit tests for ``app.docx_math_style_pre_process``: upright and bold italic letter runs of Office math get style markers."""

from __future__ import annotations

from xml.etree import ElementTree as ET

from app import docx_math_style_pre_process

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _run(text: str, properties: str = "") -> str:
    return f"<m:r><m:rPr>{properties}</m:rPr><m:t>{text}</m:t></m:r>"


def _rewrite(*runs: str) -> tuple[list[str | None], int]:
    xml = f'<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:body><w:p><m:oMath>{"".join(runs)}</m:oMath></w:p></w:body></w:document>'.encode()
    out, _ = docx_math_style_pre_process.rewrite_part(xml)
    root = ET.fromstring(out)
    return [t.text for t in root.iter(f"{{{M_NS}}}t")], len(list(root.iter(f"{{{M_NS}}}sty")))


def test_upright_and_bold_italic_letters_are_marked() -> None:
    texts, styles = _rewrite(_run("d", '<m:sty m:val="p"/>'), _run("x"), _run("v", '<m:sty m:val="bi"/>'))
    assert texts == ["PMSzzzpzzzdzzzPMSENDzzz", "x", "PMSzzzbizzzvzzzPMSENDzzz"]
    assert styles == 0, "texmath must not read the style of a marked run"


def test_operators_functions_scripts_and_other_styles_stay() -> None:
    runs = (
        _run("+", '<m:sty m:val="p"/>'),
        _run("sin", '<m:sty m:val="p"/>'),
        _run("R", '<m:scr m:val="double-struck"/><m:sty m:val="p"/>'),
        _run("t", '<m:sty m:val="i"/>'),
        _run("v", '<m:sty m:val="b"/>'),
        _run("km", '<m:nor/><m:sty m:val="p"/>'),
    )
    texts, styles = _rewrite(*runs)
    assert texts == ["+", "sin", "R", "t", "v", "km"]
    assert styles == len(runs)


def test_part_without_styles_is_returned_as_is() -> None:
    xml = f'<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:body/></w:document>'.encode()
    assert docx_math_style_pre_process.rewrite_part(xml) == (xml, False)
