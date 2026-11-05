"""Tests for app/docx_body_style.py: the body text of a DOCX as pandoc arguments for LaTeX."""

from __future__ import annotations

import io
import subprocess
import zipfile

import pytest

from app import docx_body_style
from app.docx_body_style import body_style_options

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
THEME = (
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme>'
    '<a:majorFont><a:latin typeface="Major Sans"/></a:majorFont><a:minorFont><a:latin typeface="Minor Serif"/></a:minorFont>'
    "</a:fontScheme></a:themeElements></a:theme>"
)
# The defaults and styles pandoc's own reference document carries, shortened to what counts here.
DEFAULTS = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200"/></w:pPr></w:pPrDefault></w:docDefaults>'
NORMAL = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
BODY_TEXT = '<w:style w:type="paragraph" w:styleId="BodyText"><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="180" w:after="180"/></w:pPr></w:style>'


def _docx(styles: str, theme: str | None = THEME, settings: str | None = None) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr("word/document.xml", f"<w:document {W}><w:body/></w:document>")
        package.writestr("word/styles.xml", f"<w:styles {W}>{styles}</w:styles>")
        if theme is not None:
            package.writestr("word/theme/theme1.xml", theme)
        if settings is not None:
            package.writestr("word/settings.xml", f"<w:settings {W}>{settings}</w:settings>")
    return buffer.getvalue()


def _options(docx: bytes) -> dict[str, str]:
    """The variables and metadata as one dict, keyed by name."""
    arguments = body_style_options(docx)
    assert set(arguments[::2]) <= {"-V", "-M"}
    return dict(item.split("=", 1) for item in arguments[1::2])


@pytest.fixture(autouse=True)
def installed(monkeypatch: pytest.MonkeyPatch) -> set[str]:
    families = {"Minor Serif", "Major Sans", "Body Font"}
    monkeypatch.setattr(docx_body_style, "installed_families", lambda: frozenset(families))
    return families


def test_sets_pandoc_template_like_word():
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert options == {
        "mainfont": "Minor Serif",
        "fontsize": "12pt",
        "docx-body-paragraph-skip": "9",
        "docx-body-justified": "false",
        "docx-body-hyphenated": "false",
    }


def test_a_font_named_in_the_style_goes_before_the_theme():
    body = BODY_TEXT.replace("</w:pPr>", '</w:pPr><w:rPr><w:rFonts w:ascii="Body Font"/></w:rPr>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["mainfont"] == "Body Font"


def test_the_major_theme_font_is_read_for_majorhansi():
    defaults = DEFAULTS.replace('w:asciiTheme="minorHAnsi"', 'w:asciiTheme="majorHAnsi"')
    assert _options(_docx(defaults + NORMAL + BODY_TEXT))["mainfont"] == "Major Sans"


def test_a_font_the_container_lacks_keeps_the_latex_font(installed: set[str]):
    installed.clear()
    assert "mainfont" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))


def test_a_name_which_is_no_installed_family_never_reaches_latex():
    body = BODY_TEXT.replace("</w:pPr>", '</w:pPr><w:rPr><w:rFonts w:ascii="}\\input{/etc/passwd"/></w:rPr>')
    assert "mainfont" not in _options(_docx(DEFAULTS + NORMAL + body))


def test_a_document_without_a_theme_has_no_theme_font():
    assert "mainfont" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT, theme=None))


@pytest.mark.parametrize(("half_points", "size"), [("20", "10pt"), ("22", "11pt"), ("24", "12pt"), ("28", "12pt"), ("16", "10pt")])
def test_the_size_is_the_nearest_size_of_the_article_class(half_points: str, size: str):
    defaults = DEFAULTS.replace('<w:sz w:val="24"/>', f'<w:sz w:val="{half_points}"/>')
    assert _options(_docx(defaults + NORMAL + BODY_TEXT))["fontsize"] == size


def test_no_size_leaves_the_size_of_the_template():
    defaults = DEFAULTS.replace('<w:sz w:val="24"/>', "")
    assert "fontsize" not in _options(_docx(defaults + NORMAL + BODY_TEXT))


def test_the_paragraph_skip_is_the_larger_of_before_and_after():
    body = BODY_TEXT.replace('w:before="180" w:after="180"', 'w:before="60" w:after="240"')
    assert _options(_docx(DEFAULTS + NORMAL + body))["docx-body-paragraph-skip"] == "12"


def test_a_side_the_style_leaves_out_comes_from_the_style_it_is_based_on():
    normal = NORMAL.replace("</w:style>", '<w:pPr><w:spacing w:before="300"/></w:pPr></w:style>')
    body = BODY_TEXT.replace('<w:spacing w:before="180" w:after="180"/>', '<w:spacing w:after="100" w:line="276"/>')
    assert _options(_docx(DEFAULTS + normal + body))["docx-body-paragraph-skip"] == "15"


def test_without_body_text_the_normal_style_and_the_defaults_count():
    assert _options(_docx(DEFAULTS + NORMAL))["docx-body-paragraph-skip"] == "10"


@pytest.mark.parametrize(("alignment", "justified"), [("both", "true"), ("distribute", "true"), ("left", "false"), ("center", "false")])
def test_justified_only_when_the_style_justifies(alignment: str, justified: str):
    body = BODY_TEXT.replace("</w:pPr>", f'<w:jc w:val="{alignment}"/></w:pPr>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["docx-body-justified"] == justified


@pytest.mark.parametrize(("setting", "hyphenated"), [("<w:autoHyphenation/>", "true"), ('<w:autoHyphenation w:val="false"/>', "false"), ("", "false")])
def test_hyphenated_only_when_the_document_turns_it_on(setting: str, hyphenated: str):
    assert _options(_docx(DEFAULTS + NORMAL + BODY_TEXT, settings=setting))["docx-body-hyphenated"] == hyphenated


def test_a_loop_of_based_on_ends():
    normal = NORMAL.replace("</w:style>", '<w:basedOn w:val="BodyText"/></w:style>')
    assert _options(_docx(DEFAULTS + normal + BODY_TEXT))["docx-body-paragraph-skip"] == "9"


@pytest.mark.parametrize("source", [b"not a zip", _docx("")[:20]])
def test_an_unreadable_document_gets_no_options(source: bytes):
    assert body_style_options(source) == []


def test_a_package_without_styles_gets_no_options():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr("word/document.xml", f"<w:document {W}><w:body/></w:document>")
    assert body_style_options(buffer.getvalue()) == []


def test_an_empty_body_text_style_is_the_style_and_not_normal():
    """A style without children is falsy for ElementTree: it must still be the style read, inheriting no Normal."""
    normal = NORMAL.replace("</w:style>", '<w:pPr><w:spacing w:after="120"/></w:pPr></w:style>')
    empty_body = '<w:style w:type="paragraph" w:styleId="BodyText"/>'
    assert _options(_docx(DEFAULTS + normal + empty_body))["docx-body-paragraph-skip"] == "10"


def test_installed_families_reads_every_family_of_fc_list(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.undo()
    docx_body_style.installed_families.cache_clear()
    output = "Aptos\nDejaVu Sans,DejaVu Sans Light\n\n Liberation Serif \n"
    monkeypatch.setattr(subprocess, "run", lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, stdout=output))
    try:
        assert docx_body_style.installed_families() == {"Aptos", "DejaVu Sans", "DejaVu Sans Light", "Liberation Serif"}
    finally:
        docx_body_style.installed_families.cache_clear()


def test_installed_families_is_empty_when_fc_list_fails(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.undo()
    docx_body_style.installed_families.cache_clear()

    def fail(*_args, **_kwargs):
        raise FileNotFoundError("fc-list")

    monkeypatch.setattr(subprocess, "run", fail)
    try:
        assert docx_body_style.installed_families() == frozenset()
    finally:
        docx_body_style.installed_families.cache_clear()
