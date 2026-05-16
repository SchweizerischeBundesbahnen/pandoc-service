"""Unit tests for ``app.DocxColorPreProcess``.

Each test builds a minimal DOCX zip in memory containing a single
``word/document.xml`` plus the bare ``word/styles.xml`` skeleton, runs the
preprocessor, and inspects the rewritten XML. Working against synthetic
fixtures (rather than a checked-in `.docx`) keeps the tests focused on the
specific input shape under test and makes regressions trivial to read.
"""

from __future__ import annotations

import io
import re
import zipfile
from xml.etree import ElementTree as ET  # noqa: S405

from app import DocxColorPreProcess

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
    return ET.fromstring(_unpack(blob)["word/styles.xml"])  # noqa: S314


def _body(blob: bytes) -> ET.Element:
    return ET.fromstring(_unpack(blob)["word/document.xml"])  # noqa: S314


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

    result = DocxColorPreProcess.preprocess(blob)

    assert result == blob


def test_preprocess_rewrites_fg_color_to_synthetic_style():
    blob = _pack(
        {
            "word/document.xml": _doc(_run('<w:color w:val="FF0000"/>', "red")),
            "word/styles.xml": EMPTY_STYLES_XML,
        }
    )

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

    body = _body(result)
    assert body.find(f".//{{{W_NS}}}highlight") is None
    assert _rstyle_vals(body) == ["PandocColor__HL_yellow"]
    assert "PandocColor__HL_yellow" in _style_ids(_styles(result))


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

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

    header_root = ET.fromstring(_unpack(result)["word/header1.xml"])  # noqa: S314
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

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)
    assert result == blob


def test_preprocess_skips_invalid_zip():
    """Not a zip — return bytes unchanged."""
    result = DocxColorPreProcess.preprocess(b"definitely not a docx")
    assert result == b"definitely not a docx"


def test_preprocess_skips_docx_without_styles_part():
    """Pandoc-style DOCXes always include styles.xml; if it's missing we
    don't even try to add styles to it."""
    blob = _pack({"word/document.xml": _doc(_run('<w:color w:val="FF0000"/>', "x"))})

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

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

    result = DocxColorPreProcess.preprocess(blob)

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

    once = DocxColorPreProcess.preprocess(blob)
    twice = DocxColorPreProcess.preprocess(once)

    # Comparing bytes is fragile across zip member ordering / mtime so we
    # compare the parsed body instead.
    assert _rstyle_vals(_body(twice)) == ["PandocColor__FG_FF0000"]
    # And there's still only one style registered (no duplication).
    assert _style_ids(_styles(twice)).count("PandocColor__FG_FF0000") == 1


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

    result = DocxColorPreProcess.preprocess(blob)
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

    result = DocxColorPreProcess.preprocess(blob)

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
