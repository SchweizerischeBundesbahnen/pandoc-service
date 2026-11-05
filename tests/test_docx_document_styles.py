"""Tests for app/docx_document_styles.py: the body text of a DOCX as pandoc arguments for LaTeX."""

from __future__ import annotations

import io
import subprocess
import zipfile

import pytest

from app import docx_document_styles
from app.docx_document_styles import document_style_options

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
    arguments = document_style_options(docx)
    assert set(arguments[::2]) <= {"-V", "-M"}
    return dict(item.split("=", 1) for item in arguments[1::2])


@pytest.fixture(autouse=True)
def no_font_files(monkeypatch: pytest.MonkeyPatch):
    """fc-match finds no file, so a font gets LaTeX's line unless a test gives it metrics."""

    def missing(*_args, **_kwargs):
        raise FileNotFoundError("fc-match")

    docx_document_styles.line_ratio.cache_clear()
    monkeypatch.setattr(subprocess, "run", missing)
    yield
    docx_document_styles.line_ratio.cache_clear()


@pytest.fixture(autouse=True)
def installed(monkeypatch: pytest.MonkeyPatch) -> set[str]:
    families = {"Minor Serif", "Major Sans", "Body Font"}
    monkeypatch.setattr(docx_document_styles, "installed_families", lambda: frozenset(families))
    return families


def test_sets_pandoc_template_like_word():
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert options == {
        "mainfont": "Minor Serif",
        "mainfontoptions": "Kerning=Off",
        "fontsize": "12pt",
        "docx-body-size": "12",
        "docx-body-baseline": "14.4",
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
    assert document_style_options(source) == []


def test_a_package_without_styles_gets_no_options():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        package.writestr("word/document.xml", f"<w:document {W}><w:body/></w:document>")
    assert document_style_options(buffer.getvalue()) == []


def test_an_empty_body_text_style_is_the_style_and_not_normal():
    """A style without children is falsy for ElementTree: it must still be the style read, inheriting no Normal."""
    normal = NORMAL.replace("</w:style>", '<w:pPr><w:spacing w:after="120"/></w:pPr></w:style>')
    empty_body = '<w:style w:type="paragraph" w:styleId="BodyText"/>'
    assert _options(_docx(DEFAULTS + normal + empty_body))["docx-body-paragraph-skip"] == "10"


def test_installed_families_reads_every_family_of_fc_list(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.undo()
    docx_document_styles.installed_families.cache_clear()
    output = "Aptos\nDejaVu Sans,DejaVu Sans Light\n\n Liberation Serif \n"
    monkeypatch.setattr(subprocess, "run", lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, stdout=output))
    try:
        assert docx_document_styles.installed_families() == {"Aptos", "DejaVu Sans", "DejaVu Sans Light", "Liberation Serif"}
    finally:
        docx_document_styles.installed_families.cache_clear()


def test_installed_families_is_empty_when_fc_list_fails(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.undo()
    docx_document_styles.installed_families.cache_clear()

    def fail(*_args, **_kwargs):
        raise FileNotFoundError("fc-list")

    monkeypatch.setattr(subprocess, "run", fail)
    try:
        assert docx_document_styles.installed_families() == frozenset()
    finally:
        docx_document_styles.installed_families.cache_clear()


HEADING1 = (
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:keepNext/><w:spacing w:before="360" w:after="80"/></w:pPr>'
    '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:color w:val="0f4761"/><w:sz w:val="40"/></w:rPr></w:style>'
)
HEADING3 = (
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>'
    '<w:rPr><w:rFonts w:eastAsiaTheme="majorEastAsia"/><w:i/><w:color w:val="auto"/><w:sz w:val="28"/></w:rPr></w:style>'
)
TITLE = '<w:style w:type="paragraph" w:styleId="Title"><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="80"/><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:b/><w:sz w:val="56"/></w:rPr></w:style>'


def _fields(value: str) -> dict[str, str]:
    return dict(field.split(":", 1) for field in value.split(","))


def test_a_heading_takes_the_font_size_colour_and_space_of_its_style():
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT + HEADING1))
    assert _fields(options["docx-style-Heading1"]) == {
        "font": "Major Sans",
        "size": "20",
        "line": "24",
        "bold": "false",
        "italic": "false",
        "color": "0F4761",
        "before": "18",
        "after": "9",
        "align": "left",
    }


def test_a_heading_inherits_the_font_its_style_names_no_ascii_of():
    """Heading 3 of pandoc's template sets only the East Asian theme font: the Latin one comes from Normal."""
    fields = _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + HEADING3))["docx-style-Heading3"])
    assert fields["font"] == "Minor Serif"
    assert fields["italic"] == "true"
    assert "color" not in fields, "auto is the colour of the text"


def test_the_title_is_centred_and_bold_when_its_style_says_so():
    fields = _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + TITLE))["docx-style-Title"])
    assert (fields["align"], fields["bold"], fields["size"]) == ("center", "true", "28")


def test_a_style_the_document_lacks_is_left_out():
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert not [name for name in options if name.startswith("docx-style-")]


def test_bold_switched_off_by_the_style_is_off():
    title = TITLE.replace("<w:b/>", '<w:b w:val="0"/>')
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + title))["docx-style-Title"])["bold"] == "false"


def test_a_family_with_a_separator_of_the_fields_is_left_out(installed: set[str]):
    installed.add("Odd, Family")
    heading = HEADING1.replace('<w:rFonts w:asciiTheme="majorHAnsi"/>', '<w:rFonts w:ascii="Odd, Family"/>')
    assert "font" not in _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + heading))["docx-style-Heading1"])


def test_a_colour_which_is_no_hex_colour_is_left_out():
    heading = HEADING1.replace('w:val="0f4761"', 'w:val="}\\input{x"')
    assert "color" not in _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + heading))["docx-style-Heading1"])


def _font_file(path, ascent: int, descent: int, line_gap: int) -> str:
    """A font of 1000 units an em with the hhea metrics given."""
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen

    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder([".notdef"])
    builder.setupCharacterMap({})
    builder.setupGlyf({".notdef": TTGlyphPen(None).glyph()})
    builder.setupHorizontalMetrics({".notdef": (500, 0)})
    builder.setupHorizontalHeader(ascent=ascent, descent=descent, lineGap=line_gap)
    builder.setupNameTable({"familyName": "Test", "styleName": "Regular"})
    builder.setupOS2()
    builder.setupPost()
    file = path / "test.ttf"
    builder.save(str(file))
    return str(file)


@pytest.fixture
def font_metrics(monkeypatch: pytest.MonkeyPatch, tmp_path):
    """fc-match answers with a font of 1.25 lines an em."""
    file = _font_file(tmp_path, 900, -250, 100)
    docx_document_styles.line_ratio.cache_clear()
    monkeypatch.setattr(subprocess, "run", lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, stdout=file))
    yield
    docx_document_styles.line_ratio.cache_clear()


def test_the_line_of_a_font_is_its_hhea_ascent_descent_and_gap(font_metrics):
    assert docx_document_styles.line_ratio("Minor Serif") == pytest.approx(1.25)


def test_the_body_size_is_in_points_and_its_line_comes_from_its_font(font_metrics):
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert (options["docx-body-size"], options["docx-body-baseline"]) == ("12", "15")


def test_a_line_spacing_of_the_style_multiplies_the_line(font_metrics):
    body = BODY_TEXT.replace('<w:spacing w:before="180" w:after="180"/>', '<w:spacing w:before="180" w:after="180" w:line="360" w:lineRule="auto"/>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["docx-body-baseline"] == "22.5"


@pytest.mark.parametrize(("rule", "line", "baseline"), [("exact", "200", "10"), ("atLeast", "200", "15"), ("atLeast", "400", "20")])
def test_an_exact_or_least_line_counts_in_twips(font_metrics, rule: str, line: str, baseline: str):
    body = BODY_TEXT.replace('<w:spacing w:before="180" w:after="180"/>', f'<w:spacing w:before="180" w:after="180" w:line="{line}" w:lineRule="{rule}"/>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["docx-body-baseline"] == baseline


def test_a_heading_gets_the_line_of_its_size(font_metrics):
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + HEADING1))["docx-style-Heading1"])["line"] == "25"


def test_a_font_without_metrics_gets_the_line_of_latex(monkeypatch: pytest.MonkeyPatch):
    docx_document_styles.line_ratio.cache_clear()
    monkeypatch.setattr(subprocess, "run", lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, stdout="/no/such/font.ttf"))
    try:
        assert docx_document_styles.line_ratio("Minor Serif") == docx_document_styles.DEFAULT_LINE_RATIO
    finally:
        docx_document_styles.line_ratio.cache_clear()


def test_the_latin_modern_body_gets_the_line_of_latex(installed: set[str]):
    installed.clear()
    assert _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))["docx-body-baseline"] == "14.4"
