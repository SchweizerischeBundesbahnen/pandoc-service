"""Unit tests for ``app.docx_color_pre_process``.

Each test builds a minimal DOCX zip in memory containing a single
``word/document.xml`` plus the bare ``word/styles.xml`` skeleton, runs the
preprocessor, and inspects the rewritten XML. Working against synthetic
fixtures (rather than a checked-in `.docx`) keeps the tests focused on the
specific input shape under test and makes regressions trivial to read.
"""

from __future__ import annotations

import io
import re
import time
import zipfile
from xml.etree import ElementTree as ET

from app import docx_color_pre_process

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)
NS = {"w": W_NS}

EMPTY_STYLES_XML = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'


def _pack(parts: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _unpack(blob: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _doc(*runs_xml: str) -> bytes:
    body = "".join(runs_xml)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' + body + "</w:p></w:body></w:document>").encode("utf-8")


def _run(rpr_inner: str, text: str = "x") -> str:
    return f"<w:r><w:rPr>{rpr_inner}</w:rPr><w:t>{text}</w:t></w:r>"


def _styles(blob: bytes) -> ET.Element:
    return ET.fromstring(_unpack(blob)["word/styles.xml"])


def _body(blob: bytes) -> ET.Element:
    return ET.fromstring(_unpack(blob)["word/document.xml"])


def _style_ids(styles_root: ET.Element) -> list[str]:
    return [el.get(f"{{{W_NS}}}styleId") or "" for el in styles_root.findall(f"{{{W_NS}}}style")]


def _rstyle_vals(body_root: ET.Element) -> list[str]:
    return [el.get(f"{{{W_NS}}}val") or "" for el in body_root.iter(f"{{{W_NS}}}rStyle")]


def test_preprocess_plain_docx_returns_input_unchanged():
    """No colored runs anywhere -> bytes are returned untouched (no rezip)."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run("<w:b/>", "bold word")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_preprocess_rewrites_fg_color_to_synthetic_style():
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="FF0000"/>', "red")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    body = _body(result)
    # Original <w:color> is removed.
    assert body.find(f".//{{{W_NS}}}color") is None
    # An rStyle reference now lives on the run.
    assert _rstyle_vals(body) == ["PandocColor__FG_FF0000"]
    # The matching synthetic style was registered.
    assert "PandocColor__FG_FF0000" in _style_ids(_styles(result))


def test_preprocess_rewrites_shd_to_synthetic_style():
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>', "green-bg")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    body = _body(result)
    assert body.find(f".//{{{W_NS}}}shd") is None
    assert _rstyle_vals(body) == ["PandocColor__BG_00FF00"]
    assert "PandocColor__BG_00FF00" in _style_ids(_styles(result))


def test_preprocess_rewrites_highlight_preserving_named_value():
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:highlight w:val="yellow"/>', "hl")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    body = _body(result)
    assert body.find(f".//{{{W_NS}}}highlight") is None
    assert _rstyle_vals(body) == ["PandocColor__HL_yellow"]
    assert "PandocColor__HL_yellow" in _style_ids(_styles(result))


def test_preprocess_captures_font_size():
    """Pandoc's docx reader drops run <w:sz>, so a size-only run is rewritten to
    a synthetic style carrying the half-point size (32 = 16pt)."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:sz w:val="32"/><w:szCs w:val="32"/>', "big")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    body = _body(result)
    assert body.find(f".//{{{W_NS}}}sz") is None, "direct <w:sz> must be stripped from the run"
    assert _rstyle_vals(body) == ["PandocColor__SZ_32"]
    assert "PandocColor__SZ_32" in _style_ids(_styles(result))


def test_preprocess_preserves_distinct_complex_script_size():
    """A run whose complex-script size (w:szCs) differs from w:sz must keep both
    values — szCs is encoded separately, not overwritten with the w:sz value."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:sz w:val="32"/><w:szCs w:val="28"/>', "x")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )
    result = docx_color_pre_process.preprocess(blob)
    assert _rstyle_vals(_body(result)) == ["PandocColor__SZ_32__SZCS_28"]
    # The registered style keeps the distinct szCs value (28), not 32.
    styles = _styles(result)
    style = next(el for el in styles.findall(f"{{{W_NS}}}style") if el.get(f"{{{W_NS}}}styleId") == "PandocColor__SZ_32__SZCS_28")
    assert style.find(f".//{{{W_NS}}}sz").get(f"{{{W_NS}}}val") == "32"
    assert style.find(f".//{{{W_NS}}}szCs").get(f"{{{W_NS}}}val") == "28"


def test_preprocess_combines_color_and_size():
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="FF0000"/><w:sz w:val="20"/>', "x")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )
    body = _body(docx_color_pre_process.preprocess(blob))
    assert _rstyle_vals(body) == ["PandocColor__FG_FF0000__SZ_20"]


def test_preprocess_combines_fg_bg_highlight_into_single_style():
    """A run with all three properties becomes a single combined style.
    Ordering of segments is deterministic: FG, BG, HL.
    """
    rpr_inner = '<w:color w:val="0000FF"/><w:shd w:val="clear" w:color="auto" w:fill="C0C0C0"/><w:highlight w:val="yellow"/>'
    blob = _pack(
        {
            "word/document.xml": _doc(_run(rpr_inner, "all-three")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    expected = "PandocColor__FG_0000FF__BG_C0C0C0__HL_yellow"
    assert _rstyle_vals(_body(result)) == [expected]
    assert expected in _style_ids(_styles(result))


def test_preprocess_deduplicates_styles_across_runs():
    """Two runs with identical color sets share one style entry."""
    run_xml = _run('<w:color w:val="FF0000"/>', "red") * 3
    blob = _pack(
        {
            "word/document.xml": _doc(run_xml),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_FF0000"] * 3
    style_ids = _style_ids(_styles(result))
    assert style_ids.count("PandocColor__FG_FF0000") == 1


def test_preprocess_replaces_existing_rstyle():
    """If a run already has an <w:rStyle>, the synthetic one replaces it.
    Documented behavior: the original style's name is therefore not visible
    to pandoc; its character properties would have been dropped by the
    reader regardless.
    """
    rpr_inner = '<w:rStyle w:val="MyExistingStyle"/><w:color w:val="FF0000"/>'
    blob = _pack(
        {
            "word/document.xml": _doc(_run(rpr_inner, "text")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_FF0000"]


def test_preprocess_processes_header_and_footer_parts():
    """Runs in headers/footers must also be rewritten — they share the
    <w:r>/<w:rPr> shape with the body."""
    header_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>' + _run('<w:color w:val="123456"/>', "header text") + "</w:p></w:hdr>").encode(
        "utf-8"
    )

    blob = _pack(
        {
            "word/document.xml": _doc(_run("<w:b/>", "body")),
            "word/header1.xml": header_xml,
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    header_root = ET.fromstring(_unpack(result)["word/header1.xml"])
    assert _rstyle_vals(header_root) == ["PandocColor__FG_123456"]
    assert "PandocColor__FG_123456" in _style_ids(_styles(result))


def test_preprocess_skips_w_color_auto():
    """<w:color w:val="auto"/> means "use the default colour" — there is no
    concrete color to preserve, so the run must not be rewritten."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="auto"/>', "auto")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    # Same bytes — no preprocessing happened.
    assert result == blob


def test_preprocess_skips_theme_color_reference():
    """<w:color w:themeColor="accent1"/> with no w:val attribute means the
    color comes from the theme; we cannot resolve it without parsing
    theme1.xml, so leave the run alone."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:themeColor="accent1"/>', "themed")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)
    assert result == blob


def test_preprocess_skips_invalid_zip():
    """Not a zip — return bytes unchanged."""
    result = docx_color_pre_process.preprocess(b"definitely not a docx")
    assert result == b"definitely not a docx"


def test_preprocess_skips_docx_without_styles_part():
    """Pandoc-style DOCXes always include styles.xml; if it's missing we
    don't even try to add styles to it."""
    blob = _pack({"word/document.xml": _doc(_run('<w:color w:val="FF0000"/>', "x"))})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_preprocess_synthetic_style_carries_matching_rpr():
    """The registered style includes a <w:rPr> with the original color
    properties so the intermediate DOCX still renders correctly if a human
    inspects it. Pandoc itself only needs the styleId and name."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="ABCDEF"/>', "x")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    styles = _styles(result)
    style_el = styles.find(f".//{{{W_NS}}}style[@{{{W_NS}}}styleId='PandocColor__FG_ABCDEF']")
    assert style_el is not None
    color = style_el.find(f"{{{W_NS}}}rPr/{{{W_NS}}}color")
    assert color is not None
    assert color.get(f"{{{W_NS}}}val") == "ABCDEF"


def test_preprocess_normalizes_lowercase_hex():
    """Word can emit lowercase hex; the synthetic style id uses canonical
    uppercase so equivalent runs deduplicate."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="abcdef"/>', "x"), _run('<w:color w:val="ABCDEF"/>', "y")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_ABCDEF", "PandocColor__FG_ABCDEF"]
    assert _style_ids(_styles(result)).count("PandocColor__FG_ABCDEF") == 1


def test_preprocess_idempotent_on_already_preprocessed_input():
    """Running the preprocessor twice produces the same result as running
    it once — the second pass finds no direct <w:color> to rewrite."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="FF0000"/>', "x")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    once = docx_color_pre_process.preprocess(blob)
    twice = docx_color_pre_process.preprocess(once)

    # Comparing bytes is fragile across zip member ordering / mtime so we
    # compare the parsed body instead.
    assert _rstyle_vals(_body(twice)) == ["PandocColor__FG_FF0000"]
    # And there's still only one style registered (no duplication).
    assert _style_ids(_styles(twice)).count("PandocColor__FG_FF0000") == 1


def test_preprocess_idempotent_keeps_the_size_beside_the_color():
    """A synthetic style holds a size as well as colors, so a second pass must not read it as a character style."""
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="FF0000"/><w:sz w:val="32"/>', "x")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    twice = docx_color_pre_process.preprocess(docx_color_pre_process.preprocess(blob))

    assert _rstyle_vals(_body(twice)) == ["PandocColor__FG_FF0000__SZ_32"]


def test_preprocess_preserves_drawing_namespace_prefixes():
    """Regression: ElementTree's default serialize re-numbers any namespace
    it wasn't told about as ns0/ns1/.., which breaks pandoc's DOCX reader
    (it matches drawing/relationship elements on the canonical prefix and
    silently drops <w:drawing> when wp:/a:/pic:/r: get renamed). The
    preprocessor must round-trip drawings with their original prefixes
    so images survive into the LaTeX/PDF output.
    """
    drawing_fragment = (
        '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:drawing>'
        '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<wp:extent cx="406400" cy="406400"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"/>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    body = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        "<w:document"
        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<w:body><w:p>" + drawing_fragment + "</w:p></w:body></w:document>"
    ).encode("utf-8")
    blob = _pack({"word/document.xml": body, "word/styles.xml": EMPTY_STYLES_XML})

    result = docx_color_pre_process.preprocess(blob)
    doc_xml = _unpack(result)["word/document.xml"].decode("utf-8")

    # All three drawing-related prefixes must survive serialization.
    assert "<wp:inline" in doc_xml, f"wp: prefix lost — pandoc will drop the image\n{doc_xml}"
    assert "<a:graphic" in doc_xml, f"a: prefix lost — pandoc will drop the image\n{doc_xml}"
    assert "<pic:pic" in doc_xml, f"pic: prefix lost — pandoc will drop the image\n{doc_xml}"
    # And no auto-generated ns* prefix has crept in.
    assert "<ns0:" not in doc_xml
    assert "<ns1:" not in doc_xml
    assert "xmlns:ns0=" not in doc_xml
    assert "xmlns:ns1=" not in doc_xml


def test_preprocess_real_fixture_round_trip():
    """End-to-end check against the checked-in tests/data/colored.docx fixture,
    which contains a red FG run, a green shaded run, and a yellow highlight."""
    from pathlib import Path

    blob = Path(__file__).resolve().parent.joinpath("data/colored.docx").read_bytes()

    result = docx_color_pre_process.preprocess(blob)

    body_xml = _unpack(result)["word/document.xml"].decode("utf-8")
    # Direct color/shd/highlight have been removed from body runs.
    for tag in ("w:color", "w:shd", "w:highlight"):
        # ET serializes self-closing tags as "<w:tag …/>" or "<w:tag …></w:tag>";
        # match either form by anchoring on "<w:tag" followed by space or '>'.
        assert re.search(rf"<{tag}\b[^/]*>", body_xml) is None, f"{tag} still present in body"

    style_ids = _style_ids(_styles(result))
    assert "PandocColor__FG_FF0000" in style_ids
    assert "PandocColor__BG_00FF00" in style_ids
    assert "PandocColor__HL_yellow" in style_ids


# ---- The font of a run ----

import pytest

from app import docx_document_styles


@pytest.fixture
def installed_fonts(monkeypatch: pytest.MonkeyPatch) -> set[str]:
    families = {"Segoe UI", "Calibri"}
    monkeypatch.setattr(docx_document_styles, "installed_families", lambda: frozenset(families))
    return families


def _rewrite(*runs: str) -> bytes:
    return docx_color_pre_process.preprocess(_pack({"word/document.xml": _doc(*runs), "word/styles.xml": EMPTY_STYLES_XML}))


def test_the_font_of_a_run_rides_in_its_style_as_hex(installed_fonts: set[str]):
    blob = _rewrite(_run('<w:rFonts w:ascii="segoe ui" w:hAnsi="segoe ui"/>'))
    expected = "PandocColor__FN_" + b"Segoe UI".hex().upper()
    assert _rstyle_vals(_body(blob)) == [expected]
    assert expected in _style_ids(_styles(blob))


def test_a_font_the_container_lacks_is_left_out(installed_fonts: set[str]):
    blob = _rewrite(_run('<w:rFonts w:ascii="Comic Sans"/><w:color w:val="FF0000"/>'))
    assert _rstyle_vals(_body(blob)) == ["PandocColor__FG_FF0000"]


def test_the_font_joins_the_colour_of_the_run(installed_fonts: set[str]):
    blob = _rewrite(_run('<w:rFonts w:hAnsi="Calibri"/><w:color w:val="FF0000"/>'))
    assert _rstyle_vals(_body(blob)) == ["PandocColor__FG_FF0000__FN_" + b"Calibri".hex().upper()]


def test_a_run_of_a_style_of_its_own_keeps_it_for_a_font_alone(installed_fonts: set[str]):
    blob = _rewrite(_run('<w:rStyle w:val="VerbatimChar"/><w:rFonts w:ascii="Calibri"/>'))
    assert _rstyle_vals(_body(blob)) == ["VerbatimChar"]


def test_the_installed_family_is_matched_regardless_of_case(installed_fonts: set[str]):
    assert docx_document_styles.installed_family(" SEGOE ui ") == "Segoe UI"
    assert docx_document_styles.installed_family("Segoe") is None
    assert docx_document_styles.installed_family(None) is None


def _styles_xml(*styles: str) -> bytes:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + "".join(styles) + "</w:styles>").encode("utf-8")


def _char_style(style_id: str, rpr_inner: str = "", based_on: str | None = None, name: str | None = None) -> str:
    parent = f'<w:basedOn w:val="{based_on}"/>' if based_on else ""
    return f'<w:style w:type="character" w:styleId="{style_id}"><w:name w:val="{name or style_id}"/>{parent}<w:rPr>{rpr_inner}</w:rPr></w:style>'


# The Hyperlink style of a Word template: the concrete value written beside the theme color.
HYPERLINK_STYLE = _char_style("Hyperlink", '<w:color w:val="156082" w:themeColor="accent1"/><w:u w:val="single"/>')


def test_preprocess_takes_the_color_of_the_character_style():
    """A link's runs reference the Hyperlink style and carry no color of their own.

    pandoc keeps only the style's name, so without taking its color over the link
    is printed black in the PDF while Word draws it in the template's color.
    """
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="Hyperlink"/>', "link")), "word/styles.xml": _styles_xml(HYPERLINK_STYLE)})

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_156082"]
    assert "PandocColor__FG_156082" in _style_ids(_styles(result))


def test_preprocess_follows_based_on_for_the_style_color():
    """The color can come from the style a character style is based on; the nearest one wins."""
    styles = _styles_xml(
        _char_style("Base", '<w:color w:val="00FF00"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'),
        _char_style("Middle", based_on="Base"),
        _char_style("Link", '<w:color w:val="0000FF"/>', based_on="Middle"),
    )
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="Link"/>', "link")), "word/styles.xml": styles})

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_0000FF__BG_FFFF00"]


def test_preprocess_direct_color_wins_over_the_style_color():
    """A color set on the run itself - a deleted Work Item drawn red inside a link - is kept."""
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="Hyperlink"/><w:color w:val="E74C3C"/>', "EL-105")), "word/styles.xml": _styles_xml(HYPERLINK_STYLE)})

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_E74C3C"]


def test_preprocess_direct_auto_color_is_not_replaced_by_the_style_color():
    """<w:color w:val="auto"/> on the run means "no color" for it, whatever its style says."""
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="Hyperlink"/><w:color w:val="auto"/>', "plain")), "word/styles.xml": _styles_xml(HYPERLINK_STYLE)})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_preprocess_leaves_a_style_without_color_alone():
    """A character style that sets no color (e.g. only bold) gives the run nothing to rewrite."""
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="Strong"/>', "bold")), "word/styles.xml": _styles_xml(_char_style("Strong", "<w:b/>"))})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_preprocess_survives_a_based_on_cycle():
    """A hand-edited template may chain styles into a cycle; resolving it must terminate."""
    styles = _styles_xml(_char_style("A", '<w:color w:val="FF0000"/>', based_on="B"), _char_style("B", based_on="A"))
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="B"/>', "x")), "word/styles.xml": styles})

    result = docx_color_pre_process.preprocess(blob)

    assert _rstyle_vals(_body(result)) == ["PandocColor__FG_FF0000"]


def test_character_style_colors_resolves_every_style_of_a_cycle():
    """Each style of a cycle, and a style based on one of them, sees the color held by another."""
    styles = _styles_xml(
        _char_style("A", '<w:color w:val="FF0000"/>', based_on="B"),
        _char_style("B", based_on="A"),
        _char_style("C", based_on="B"),
    )

    colors = docx_color_pre_process.character_style_colors(styles)

    assert colors == dict.fromkeys("ABC", docx_color_pre_process.StyleColors("FF0000", None, None))


def test_character_style_colors_takes_linear_time_on_a_long_chain():
    """An uploaded document decides how long a basedOn chain is, and the walk runs in a conversion slot.

    Walking the chain again for each style took about 25 s for these styles.
    """
    count = 20_000
    chain = [_char_style("S0", '<w:color w:val="FF0000"/>')] + [_char_style(f"S{index}", based_on=f"S{index - 1}") for index in range(1, count)]

    started = time.perf_counter()
    colors = docx_color_pre_process.character_style_colors(_styles_xml(*chain))

    assert time.perf_counter() - started < 5
    assert len(colors) == count
    assert colors[f"S{count - 1}"].fg == "FF0000"


def test_character_style_colors_keeps_what_a_style_clears():
    """auto and none clear what the parent style sets, they do not leave it to the parent."""
    styles = _styles_xml(
        _char_style("Base", '<w:color w:val="FF0000"/><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/><w:highlight w:val="yellow"/>'),
        _char_style("NoColor", '<w:color w:val="auto"/>', based_on="Base"),
        _char_style("NoBackground", '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>', based_on="Base"),
        _char_style("NoHighlight", '<w:highlight w:val="none"/>', based_on="Base"),
        _char_style("Nothing", '<w:color w:val="auto"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:highlight w:val="none"/>', based_on="Base"),
        _char_style("Colored", '<w:color w:val="0000FF"/>', based_on="Nothing"),
    )

    colors = docx_color_pre_process.character_style_colors(styles)

    assert colors["NoColor"] == docx_color_pre_process.StyleColors(None, "00FF00", "yellow")
    assert colors["NoBackground"] == docx_color_pre_process.StyleColors("FF0000", None, "yellow")
    assert colors["NoHighlight"] == docx_color_pre_process.StyleColors("FF0000", "00FF00", None)
    assert "Nothing" not in colors
    assert colors["Colored"] == docx_color_pre_process.StyleColors("0000FF", None, None)


def test_preprocess_leaves_highlighted_code_to_pandoc():
    """pandoc reads a run as code through its style. A synthetic style would set the token in the font of the text."""
    styles = _styles_xml(
        _char_style("VerbatimChar", '<w:sz w:val="22"/>', name="Verbatim Char"),
        _char_style("DataTypeTok", '<w:color w:val="902000"/>', based_on="VerbatimChar"),
    )
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="DataTypeTok"/>', "int")), "word/styles.xml": styles})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_preprocess_leaves_a_colored_verbatim_style_alone():
    """A template may color Verbatim Char itself, under any id and in any case."""
    styles = _styles_xml(_char_style("Quellcode", '<w:color w:val="902000"/>', name="verbatim char"))
    blob = _pack({"word/document.xml": _doc(_run('<w:rStyle w:val="Quellcode"/>', "code")), "word/styles.xml": styles})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_preprocess_keeps_translated_table_of_contents_entries_uncolored():
    """Word translates the id of a built-in style, so a German table of contents is found by the style's name."""
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Verzeichnis1"/></w:pPr><w:hyperlink w:anchor="_h1">' + _run('<w:rStyle w:val="Hyperlink"/>', "Heading") + "</w:hyperlink></w:p>"
        '<w:p><w:pPr><w:pStyle w:val="Abbildungsverzeichnis"/></w:pPr><w:hyperlink w:anchor="_f1">' + _run('<w:rStyle w:val="Hyperlink"/>', "Figure") + "</w:hyperlink></w:p>"
        "</w:body></w:document>"
    ).encode("utf-8")
    styles = _styles_xml(
        HYPERLINK_STYLE,
        '<w:style w:type="paragraph" w:styleId="Verzeichnis1"><w:name w:val="toc 1"/></w:style>',
        '<w:style w:type="paragraph" w:styleId="Abbildungsverzeichnis"><w:name w:val="table of figures"/></w:style>',
    )
    blob = _pack({"word/document.xml": document, "word/styles.xml": styles})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_table_of_contents_styles_reads_paragraph_styles_by_name():
    styles = _styles_xml(
        '<w:style w:type="paragraph" w:styleId="TM9"><w:name w:val="TOC 9"/></w:style>',
        '<w:style w:type="paragraph" w:styleId="Tabledesillustrations"><w:name w:val="table of figures"/></w:style>',
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>',
        '<w:style w:type="character" w:styleId="TM1Char"><w:name w:val="toc 1"/></w:style>',
    )

    assert docx_color_pre_process.table_of_contents_styles(styles) == {"TM9", "Tabledesillustrations"}
    assert docx_color_pre_process.table_of_contents_styles(b"not xml") == frozenset()


def test_preprocess_keeps_table_of_contents_entries_uncolored():
    """The entries of a table of contents carry the Hyperlink style but are not drawn as links."""
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_h1">' + _run('<w:rStyle w:val="Hyperlink"/>', "Heading") + "</w:hyperlink></w:p>"
        '<w:p><w:pPr><w:pStyle w:val="TableofFigures"/></w:pPr><w:hyperlink w:anchor="_f1">' + _run('<w:rStyle w:val="Hyperlink"/>', "Figure 1") + "</w:hyperlink></w:p>"
        "</w:body></w:document>"
    ).encode("utf-8")
    blob = _pack({"word/document.xml": document, "word/styles.xml": _styles_xml(HYPERLINK_STYLE)})

    result = docx_color_pre_process.preprocess(blob)

    assert result == blob


def test_character_style_colors_resolves_only_colored_character_styles():
    styles = _styles_xml(
        HYPERLINK_STYLE,
        _char_style("Strong", "<w:b/>"),
        '<w:style w:type="paragraph" w:styleId="Red"><w:name w:val="Red"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>',
    )

    colors = docx_color_pre_process.character_style_colors(styles)

    assert colors == {"Hyperlink": docx_color_pre_process.StyleColors("156082", None, None)}


THEME_XML = b'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:accent1><a:srgbClr val="156082"/></a:accent1></a:clrScheme></a:themeElements></a:theme>'


def test_character_style_colors_take_the_theme_colour_as_word_does():
    """Word draws a theme colour from the theme, whatever value is written beside it."""
    styles = _styles_xml(_char_style("Hyperlink", '<w:color w:val="4F81BD" w:themeColor="accent1"/>'))
    assert docx_color_pre_process.character_style_colors(styles, THEME_XML) == {"Hyperlink": docx_color_pre_process.StyleColors("156082", None, None)}
    assert docx_color_pre_process.character_style_colors(styles) == {"Hyperlink": docx_color_pre_process.StyleColors("4F81BD", None, None)}
