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
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>'
    '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:accent1><a:srgbClr val="156082"/></a:accent1>'
    '<a:hlink><a:srgbClr val="467886"/></a:hlink></a:clrScheme><a:fontScheme>'
    '<a:majorFont><a:latin typeface="Major Sans"/></a:majorFont><a:minorFont><a:latin typeface="Minor Serif"/></a:minorFont>'
    "</a:fontScheme></a:themeElements></a:theme>"
)
# The defaults and styles pandoc's own reference document carries, shortened to what counts here.
DEFAULTS = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200"/></w:pPr></w:pPrDefault></w:docDefaults>'
NORMAL = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
BODY_TEXT = '<w:style w:type="paragraph" w:styleId="BodyText"><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="180" w:after="180"/></w:pPr></w:style>'


def _docx(styles: str, theme: str | None = THEME, settings: str | None = None, section: str = "<w:sectPr/>", numbering: str | None = None) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as package:
        if numbering is not None:
            package.writestr("word/numbering.xml", f"<w:numbering {W}>{numbering}</w:numbering>")
        package.writestr("word/document.xml", f"<w:document {W}><w:body>{section}</w:body></w:document>")
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
    docx_document_styles.ascent_ratio.cache_clear()
    monkeypatch.setattr(subprocess, "run", missing)
    yield
    docx_document_styles.line_ratio.cache_clear()
    docx_document_styles.ascent_ratio.cache_clear()


@pytest.fixture(autouse=True)
def installed(monkeypatch: pytest.MonkeyPatch) -> set[str]:
    families = {"Minor Serif", "Major Sans", "Body Font"}
    monkeypatch.setattr(docx_document_styles, "installed_families", lambda: frozenset(families))
    return families


def test_sets_pandoc_template_like_word():
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert options == {
        "mainfont": "Minor Serif",
        "mainfontoptions": "Ligatures={TeXOff,NoCommon}",
        "fontsize": "12pt",
        "docx-body-size": "12",
        "docx-body-baseline": "14.4",
        "docx-body-paragraph-skip": "9",
        "docx-body-justified": "false",
        "docx-body-hyphenated": "false",
        "docx-page-numbers": "false",
        "docx-table-plain-cell-space": "10",
        "docx-paragraph-spacing": "Normal:Normal:0:10",
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


def test_a_word_font_the_container_lacks_takes_its_open_substitute(installed: set[str]):
    installed.add("Metric Sans")
    body = BODY_TEXT.replace("</w:pPr>", '</w:pPr><w:rPr><w:rFonts w:ascii="Aptos"/></w:rPr>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["mainfont"] == "Metric Sans"


def test_the_word_font_goes_before_its_substitute(installed: set[str]):
    installed.update({"Aptos", "Metric Sans"})
    body = BODY_TEXT.replace("</w:pPr>", '</w:pPr><w:rPr><w:rFonts w:ascii="Aptos"/></w:rPr>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["mainfont"] == "Aptos"


@pytest.mark.parametrize(
    ("name", "substitute"),
    [("segoe ui", "Metric UI"), ("Calibri", "Carlito"), ("Georgia", "Metric Serif"), ("Aptos Display", "Metric Sans Display"), ("Arial", "Liberation Sans")],
)
def test_a_font_a_run_names_takes_its_substitute(installed: set[str], name: str, substitute: str):
    installed.add(substitute)
    assert docx_document_styles.installed_family(name) == substitute


def test_a_substitute_the_container_lacks_is_no_font(installed: set[str]):
    assert docx_document_styles.installed_family("Aptos") is None


def test_the_equations_take_cambria_math_when_it_is_installed(installed: set[str]):
    installed.add("Cambria Math")
    assert _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))["mathfont"] == "Cambria Math"


def test_without_cambria_math_the_equations_keep_the_latex_font():
    assert "mathfont" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))


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
        "after": "4",
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
    docx_document_styles.ascent_ratio.cache_clear()
    monkeypatch.setattr(subprocess, "run", lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, stdout=file))
    yield
    docx_document_styles.line_ratio.cache_clear()
    docx_document_styles.ascent_ratio.cache_clear()


def test_the_line_of_a_font_is_its_hhea_ascent_descent_and_gap(font_metrics):
    assert docx_document_styles.line_ratio("Minor Serif") == pytest.approx(1.25)


def test_the_body_size_is_in_points_and_its_line_comes_from_its_font(font_metrics):
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert (options["docx-body-size"], options["docx-body-baseline"]) == ("12", "15")


def test_the_body_gets_where_its_line_starts_and_ends(font_metrics):
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert (options["docx-body-ascent"], options["docx-body-descent"]) == ("10.8", "4.2")


def test_a_body_font_without_metrics_gets_no_ascent():
    assert "docx-body-ascent" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))


def test_a_line_spacing_of_the_style_multiplies_the_line(font_metrics):
    body = BODY_TEXT.replace('<w:spacing w:before="180" w:after="180"/>', '<w:spacing w:before="180" w:after="180" w:line="360" w:lineRule="auto"/>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["docx-body-baseline"] == "22.5"


@pytest.mark.parametrize(("rule", "line", "baseline"), [("exact", "200", "10"), ("atLeast", "200", "15"), ("atLeast", "400", "20")])
def test_an_exact_or_least_line_counts_in_twips(font_metrics, rule: str, line: str, baseline: str):
    body = BODY_TEXT.replace('<w:spacing w:before="180" w:after="180"/>', f'<w:spacing w:before="180" w:after="180" w:line="{line}" w:lineRule="{rule}"/>')
    assert _options(_docx(DEFAULTS + NORMAL + body))["docx-body-baseline"] == baseline


def test_a_heading_gets_the_line_of_its_size(font_metrics):
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + HEADING1))["docx-style-Heading1"])["line"] == "25"


def test_a_heading_gets_the_ascent_of_its_font(font_metrics):
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + HEADING1))["docx-style-Heading1"])["ascent"] == "18"


def test_a_font_without_metrics_gets_no_ascent(monkeypatch: pytest.MonkeyPatch):
    docx_document_styles.ascent_ratio.cache_clear()
    monkeypatch.setattr(subprocess, "run", lambda *_args, **_kwargs: subprocess.CompletedProcess([], 0, stdout="/no/such/font.ttf"))
    try:
        assert docx_document_styles.ascent_ratio("Minor Serif") is None
        assert docx_document_styles.ascent_ratio(None) is None
    finally:
        docx_document_styles.ascent_ratio.cache_clear()


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


HYPERLINK = '<w:style w:type="character" w:styleId="Hyperlink"><w:rPr><w:color w:val="4F81BD" w:themeColor="accent1"/></w:rPr></w:style>'


def test_links_get_no_colour_of_their_own():
    """Word colours the text of a link only by its runs: the Hyperlink style, which docx_color_pre_process carries over."""
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT + HYPERLINK))
    assert "docx-link-color" not in options
    assert "colorlinks" not in options


def test_a_theme_shade_darkens_as_word_does():
    heading = HEADING1.replace('<w:color w:val="0f4761"/>', '<w:color w:val="000000" w:themeColor="accent1" w:themeShade="BF"/>')
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + heading))["docx-style-Heading1"])["color"] == "0F4761"


def test_a_theme_tint_lightens():
    heading = HEADING1.replace('<w:color w:val="0f4761"/>', '<w:color w:val="000000" w:themeColor="text1" w:themeTint="80"/>')
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + heading))["docx-style-Heading1"])["color"] == "7F7F7F"


def test_a_theme_colour_the_theme_lacks_falls_back_to_the_value():
    heading = HEADING1.replace('<w:color w:val="0f4761"/>', '<w:color w:val="123456" w:themeColor="accent6"/>')
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + heading))["docx-style-Heading1"])["color"] == "123456"


def test_a_document_without_a_hyperlink_style_leaves_links_alone():
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))
    assert "docx-link-color" not in options
    assert "colorlinks" not in options


def test_a_document_with_a_footer_keeps_its_page_numbers():
    section = '<w:sectPr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:footerReference w:type="default" r:id="rId9"/></w:sectPr>'
    assert "docx-page-numbers" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT, section=section))


COMPACT = '<w:style w:type="paragraph" w:styleId="Compact"><w:basedOn w:val="BodyText"/><w:pPr><w:spacing w:before="36" w:after="36"/></w:pPr></w:style>'


def test_a_cell_holds_the_space_before_and_after_its_compact_paragraph():
    assert _options(_docx(DEFAULTS + NORMAL + BODY_TEXT + COMPACT))["docx-table-cell-space"] == "3.6"


def test_a_document_without_the_compact_style_leaves_tables_alone():
    assert "docx-table-cell-space" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT))


TABLE_STYLE = (
    '<w:style w:type="table" w:default="1" w:styleId="Table"><w:tblPr><w:tblCellMar>'
    '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
    "</w:tblCellMar></w:tblPr></w:style>"
)


def test_a_cell_has_the_margin_of_the_default_table_style():
    assert _options(_docx(DEFAULTS + NORMAL + BODY_TEXT + TABLE_STYLE))["docx-table-cell-margin"] == "5.4"


def test_margins_which_differ_or_are_no_twips_are_left_out():
    uneven = TABLE_STYLE.replace('<w:right w:w="108" w:type="dxa"/>', '<w:right w:w="200" w:type="dxa"/>')
    percent = TABLE_STYLE.replace('w:w="108" w:type="dxa"', 'w:w="5" w:type="pct"')
    no_default = TABLE_STYLE.replace(' w:default="1"', "")
    for style in (uneven, percent, no_default):
        assert "docx-table-cell-margin" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT + style))


NUMBERED_HEADING1 = HEADING1.replace("<w:keepNext/>", '<w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>')


def _numbering(text: str = "%1.%2", second_format: str = "lowerLetter") -> str:
    return (
        '<w:abstractNum w:abstractNumId="7">'
        '<w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1"/></w:lvl>'
        f'<w:lvl w:ilvl="1"><w:numFmt w:val="{second_format}"/><w:lvlText w:val="{text}"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl>'
        '</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="7"/></w:num>'
    )


def test_a_heading_style_with_numbering_carries_its_number():
    fields = _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + NUMBERED_HEADING1, numbering=_numbering()))["docx-style-Heading1"])
    assert (fields["number"], fields["numberhang"]) == ("{1:Roman}.{2:alph}", "21.6")


@pytest.mark.parametrize(("text", "second_format"), [("%1.%2", "bullet"), ("§%2", "decimal"), ("%1.%3", "decimal")])
def test_a_number_this_cannot_set_is_left_out(text: str, second_format: str):
    fields = _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + NUMBERED_HEADING1, numbering=_numbering(text, second_format)))["docx-style-Heading1"])
    assert "number" not in fields


def test_a_heading_without_the_numbering_part_has_no_number():
    assert "number" not in _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + NUMBERED_HEADING1))["docx-style-Heading1"])


def test_a_localised_heading_style_is_found_by_its_name():
    localised = HEADING1.replace('w:styleId="Heading1">', 'w:styleId="berschrift1"><w:name w:val="heading 1"/>')
    assert _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + localised))["docx-style-Heading1"])["size"] == "20"


def test_the_items_of_a_list_get_the_larger_space_of_compact():
    assert _options(_docx(DEFAULTS + NORMAL + BODY_TEXT + COMPACT))["docx-list-item-space"] == "1.8"


def _bullets(font: str = "Body Font") -> str:
    plain = '<w:abstractNum w:abstractNumId="990"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>'
    used = f'<w:abstractNum w:abstractNumId="991"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:rPr><w:rFonts w:ascii="{font}"/></w:rPr></w:lvl></w:abstractNum>'
    return plain + used + '<w:num w:numId="1"><w:abstractNumId w:val="990"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="991"/></w:num>'


LIST_ITEM = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:p><w:sectPr/>'


def test_a_bullet_reaches_as_high_as_its_font_in_the_list_of_the_document(font_metrics):
    """Word sets the line of an item as high as the bullet's font: the first level of the list the document uses."""
    options = _options(_docx(DEFAULTS + NORMAL + BODY_TEXT, section=LIST_ITEM, numbering=_bullets()))
    assert options["docx-list-bullet-ascents"] == "10.8,0,0,0"


def test_a_bullet_in_a_font_which_is_not_installed_gets_no_ascent(font_metrics):
    assert "docx-list-bullet-ascents" not in _options(_docx(DEFAULTS + NORMAL + BODY_TEXT, section=LIST_ITEM, numbering=_bullets("Missing")))


def test_the_spacing_of_normal_and_the_styles_of_the_body_is_handed_over():
    body = '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p><w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr></w:p><w:sectPr/>'
    styles = DEFAULTS + NORMAL + BODY_TEXT + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>'
    spacing = _options(_docx(styles, section=body))["docx-paragraph-spacing"]
    # Normal takes its space after from the defaults; the title has spaces of its own
    assert spacing == "BodyText:BodyText:9:9;Normal:Normal:0:10"


def _indented(*levels: tuple[int, int]) -> str:
    """A list whose levels have the left and hanging indents given, in twips."""
    lvls = "".join(f'<w:lvl w:ilvl="{index}"><w:numFmt w:val="decimal"/><w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>' for index, (left, hanging) in enumerate(levels))
    return f'<w:abstractNum w:abstractNumId="991">{lvls}</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="991"/></w:num>'


def test_each_level_of_a_list_is_indented_past_the_level_above_it():
    """Word sets the label at the start of the hanging indent and the text at the left indent."""
    numbering = _indented((720, 360), (1440, 360), (2160, 180), (2880, 360))
    assert _options(_docx(DEFAULTS + NORMAL, section=LIST_ITEM, numbering=numbering))["docx-list-indents"] == "36:18,36:18,36:9,36:18"


def test_a_list_with_a_level_not_past_the_one_above_gets_the_indents_of_latex():
    numbering = _indented((720, 360), (720, 360), (2160, 360), (2880, 360))
    assert "docx-list-indents" not in _options(_docx(DEFAULTS + NORMAL, section=LIST_ITEM, numbering=numbering))


def test_a_list_indent_of_a_digit_python_cannot_read_falls_back_to_latex():
    """'²'.isdigit() is true, but int('²') fails: the conversion keeps the indents of LaTeX instead of failing."""
    numbering = _indented((720, 360), (1440, 360), (2160, 360), (2880, 360)).replace('w:left="1440"', 'w:left="²"')
    assert "docx-list-indents" not in _options(_docx(DEFAULTS + NORMAL, section=LIST_ITEM, numbering=numbering))


def test_a_heading_hanging_indent_of_a_digit_python_cannot_read_hangs_nothing():
    numbering = _numbering().replace('w:hanging="432"', 'w:hanging="²"')
    fields = _fields(_options(_docx(DEFAULTS + NORMAL + BODY_TEXT + NUMBERED_HEADING1, numbering=numbering))["docx-style-Heading1"])
    assert fields["numberhang"] == "0"
