"""End-to-end integration tests for filters/inline_styles.lua.

These tests invoke a real `pandoc` binary with the lua filter loaded, then
inspect the resulting DOCX package. They skip cleanly when pandoc is not on
PATH so unit-test runs on dev machines without pandoc are unaffected.

Why an integration test (vs. mocked unit tests):
    The other tests for this filter mock subprocess.run and only verify that
    the right --lua-filter argument is added to the command line. They cannot
    catch regressions where the filter still loads but produces AST that
    pandoc's DOCX writer no longer renders correctly — for example, a future
    pandoc version changing how RawInline children of a Link node propagate
    into <w:hyperlink>, or a filter edit that drops the Link wrapper. This
    file plugs that gap with a single, focused round-trip assertion.
"""

import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

PANDOC = shutil.which("pandoc")
FILTER_PATH = Path(__file__).resolve().parents[1] / "filters" / "inline_styles.lua"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

pytestmark = pytest.mark.skipif(
    PANDOC is None or not FILTER_PATH.exists(),
    reason="pandoc binary or filters/inline_styles.lua not available",
)


def _convert_html_to_docx(html: str, output_path: Path) -> None:
    """Run pandoc directly with the local filter file. Surfaces pandoc's
    stderr verbatim if it fails, since lua errors land there."""
    src_path = output_path.with_suffix(".html")
    src_path.write_text(html, encoding="utf-8")
    result = subprocess.run(
        [
            PANDOC,
            "-f",
            "html",
            "-t",
            "docx",
            f"--lua-filter={FILTER_PATH}",
            "-o",
            str(output_path),
            str(src_path),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise AssertionError(f"pandoc failed (exit {result.returncode}):\n{result.stderr}")


def test_hyperlink_inside_styled_span_survives_filter():
    """Regression: <a> inside a styled <span> must remain clickable.

    Catches two failure modes:
      1. Filter dropping the Link wrapper -> no <w:hyperlink> in document.xml,
         no Relationship entry in word/_rels/document.xml.rels (the bug we
         just fixed in filters/inline_styles.lua walk()'s Link branch).
      2. Filter passing the Link through unchanged without re-walking its
         content -> link works but inner runs lack the surrounding span's
         color, defeating the walk-and-rewrap design.

    A future pandoc version that changes how RawInline children inside a
    Link node propagate into <w:hyperlink> would also surface here.
    """
    html = '<p><span style="color:#FF0000;"><a href="https://example.com/test">click here</a></span></p>'

    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")
            rels_xml = zf.read("word/_rels/document.xml.rels")

    doc = ET.fromstring(doc_xml)
    rels = ET.fromstring(rels_xml)

    # 1. <w:hyperlink> exists at all (proves the Link wrapper was preserved).
    hyperlink = doc.find(f".//{{{W_NS}}}hyperlink")
    assert hyperlink is not None, f"no <w:hyperlink> in document.xml — filter dropped the Link node\ndocument.xml head: {doc_xml[:1500]!r}"

    rid = hyperlink.get(f"{{{R_NS}}}id")
    assert rid, "hyperlink element has no r:id"

    # 2. The hyperlink's runs carry the surrounding span's color. This is the
    #    walk-and-rewrap part: a naive pass-through Link would emit default
    #    styling, but we want the outer span's red to apply to link text.
    color_vals = [c.get(f"{{{W_NS}}}val") for c in hyperlink.findall(f".//{{{W_NS}}}color")]
    assert "FF0000" in color_vals, f"hyperlink runs do not carry FF0000 (found colors: {color_vals}) — walk-and-rewrap regression in filters/inline_styles.lua Link branch"

    # 3. Link text survived.
    text = "".join(t.text or "" for t in hyperlink.iter(f"{{{W_NS}}}t"))
    assert "click here" in text, f"link text missing from hyperlink runs: {text!r}"

    # 4. The matching relationship targets the original href as External.
    rel = rels.find(f".//{{{PKG_NS}}}Relationship[@Id='{rid}']")
    assert rel is not None, f"no Relationship for {rid} in word/_rels/document.xml.rels — the relationship side-effect that pandoc creates from a Link node did not fire (Link AST node was probably dropped before the writer saw it)"
    assert rel.get("Target") == "https://example.com/test", f"hyperlink relationship Target mismatch: {rel.get('Target')!r}"
    assert rel.get("TargetMode") == "External", f"hyperlink relationship TargetMode mismatch: {rel.get('TargetMode')!r}"
    assert rel.get("Type", "").endswith("/hyperlink"), f"unexpected Relationship Type: {rel.get('Type')!r}"


def test_intersecting_text_decorations_are_additive():
    """Regression: child text-decoration must add to (not replace) inherited
    decorations.

    Bug: prior to fix, a nested span setting `text-decoration: line-through`
    inside a `text-decoration: underline` ancestor unconditionally overwrote
    both decoration flags from the child's token list, wiping the inherited
    underline (and vice versa for the inverse direction). CSS draws ancestor
    decorations through descendants regardless of the descendant's own
    text-decoration — only `none` clears them.

    Three scenarios are checked in one test to cover both directions plus
    the `none` escape hatch we intentionally preserved.
    """
    html = (
        "<p>"
        '<span style="text-decoration: underline">'
        "u_outer and "
        '<span style="text-decoration: line-through">u_inner</span>'
        "</span>"
        "</p>"
        "<p>"
        '<span style="text-decoration: line-through">'
        "s_outer and "
        '<span style="text-decoration: underline">s_inner</span>'
        "</span>"
        "</p>"
        "<p>"
        '<span style="text-decoration: underline">'
        "u_keep and "
        '<span style="text-decoration: none">u_clear</span>'
        "</span>"
        "</p>"
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)

    def _decorations_for(needle: str) -> set[str]:
        """Local-name tags from the <w:rPr> of the first <w:r> whose
        concatenated <w:t> text contains `needle`. Empty set when the run
        has no <w:rPr>."""
        for r in doc.iter(f"{{{W_NS}}}r"):
            text = "".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t"))
            if needle in text:
                rpr = r.find(f"{{{W_NS}}}rPr")
                if rpr is None:
                    return set()
                return {c.tag.split("}", 1)[-1] for c in rpr}
        all_run_text = ["".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t")) for r in doc.iter(f"{{{W_NS}}}r")]
        raise AssertionError(f"no <w:r> contained {needle!r}; runs were: {all_run_text!r}")

    # Case 1: outer underline + inner line-through → inner range must have BOTH.
    assert _decorations_for("u_outer") == {"u"}, "outer underline range lost its decoration"
    inner_u = _decorations_for("u_inner")
    assert "u" in inner_u and "strike" in inner_u, f"inner range lost inherited underline when it added line-through (got {inner_u!r}) — text-decoration merge regression in filters/inline_styles.lua merge_css"

    # Case 2: outer line-through + inner underline → symmetric — inner must have BOTH.
    assert _decorations_for("s_outer") == {"strike"}, "outer strike range lost its decoration"
    inner_s = _decorations_for("s_inner")
    assert "u" in inner_s and "strike" in inner_s, f"inner range lost inherited strike when it added underline (got {inner_s!r}) — text-decoration merge regression (inverse direction)"

    # Case 3: `text-decoration: none` on a descendant still clears inherited decorations.
    assert _decorations_for("u_keep") == {"u"}, "outer underline range lost its decoration"
    inner_none = _decorations_for("u_clear")
    assert "u" not in inner_none and "strike" not in inner_none, f"text-decoration: none failed to clear inherited decorations (got {inner_none!r}) — the explicit clear escape hatch is broken"


def test_superscript_wrapping_styled_span_keeps_both():
    """Regression: when <sup>/<sub> ENCLOSES a styled <span>, the inner run must
    keep BOTH the vertical alignment and the span's formatting.

    Topdown traversal reaches the Superscript/Subscript wrapper first; without a
    dedicated handler, filter.Span would consume the span subtree with no
    knowledge of the surrounding sup/sub and the DOCX writer would drop the
    vertAlign on those runs (it ignores the AST wrapper around raw OOXML runs).
    """
    html = '<p><sup>A<span style="color:#FF0000">B</span></sup> x <sub>C<span style="color:#00B050">D</span></sub></p>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc = ET.fromstring(zf.read("word/document.xml"))

    def _run(needle: str) -> ET.Element:
        for r in doc.iter(f"{{{W_NS}}}r"):
            if needle in "".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t")):
                return r
        raise AssertionError(f"no run with {needle!r}")

    def _vert_align(r: ET.Element) -> str | None:
        el = r.find(f".//{{{W_NS}}}vertAlign")
        return el.get(f"{{{W_NS}}}val") if el is not None else None

    def _color(r: ET.Element) -> str | None:
        el = r.find(f".//{{{W_NS}}}color")
        return el.get(f"{{{W_NS}}}val") if el is not None else None

    # The styled-span runs inside the wrapper keep BOTH properties.
    assert _vert_align(_run("B")) == "superscript" and _color(_run("B")) == "FF0000", "superscript lost on the colored run inside <sup>"
    assert _vert_align(_run("D")) == "subscript" and _color(_run("D")) == "00B050", "subscript lost on the colored run inside <sub>"
    # The plain part of the wrapper still gets the vertical alignment too.
    assert _vert_align(_run("A")) == "superscript"
    assert _vert_align(_run("C")) == "subscript"


# ---------------------------------------------------------------------------
# Paragraph formatting (Div handler) — see filter.Div in filters/inline_styles.lua
# ---------------------------------------------------------------------------
#
# These tests cover the contract with app/HtmlParagraphPreProcess.py: a
# <div class="pandoc-para" data-indent-twips="N" data-text-align="..."><p>...</p></div>
# wrapper becomes a raw OOXML <w:p> carrying <w:pPr><w:ind w:left="N"/>
# <w:jc w:val="..."/></w:pPr> and the original inlines rendered as <w:r> runs.


def _w_p_with_text(doc: ET.Element, needle: str) -> ET.Element:
    """Return the first <w:p> whose concatenated <w:t> text contains needle."""
    for p in doc.iter(f"{{{W_NS}}}p"):
        text = "".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t"))
        if needle in text:
            return p
    all_text = ["".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t")) for p in doc.iter(f"{{{W_NS}}}p")]
    raise AssertionError(f"no <w:p> contained {needle!r}; paragraphs were: {all_text!r}")


def _ind_left(p: ET.Element) -> str | None:
    """Read <w:ind w:left=...> off a paragraph, or None when not set."""
    ind = p.find(f".//{{{W_NS}}}ind")
    if ind is None:
        return None
    return ind.get(f"{{{W_NS}}}left")


def _jc_val(p: ET.Element) -> str | None:
    """Read <w:jc w:val=...> off a paragraph, or None when not set."""
    jc = p.find(f".//{{{W_NS}}}jc")
    if jc is None:
        return None
    return jc.get(f"{{{W_NS}}}val")


def test_indent_div_sets_w_ind_left_in_twips():
    """The canonical Polarion case: two paragraphs at different indents.

    The filter must emit <w:p> with <w:pPr><w:ind w:left="N"/></w:pPr>, where
    N matches the data-indent-twips on the wrapper. 40px = 600 twips and
    80px = 1200 twips (1 px = 15 twips at the CSS reference DPI).
    """
    html = '<div class="pandoc-para" data-indent-twips="600"><p>Indentation</p></div><div class="pandoc-para" data-indent-twips="1200"><p>2 levels</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    assert _ind_left(_w_p_with_text(doc, "Indentation")) == "600"
    assert _ind_left(_w_p_with_text(doc, "2 levels")) == "1200"


def test_indent_preserves_inline_formatting_via_walk():
    """Nested <strong>/<em>/styled <span> inside an indented paragraph must
    still produce the right run properties. This is the key reason the Div
    handler reuses the existing walk() — without it, we'd lose all inline
    styling when rewriting the Para as raw OOXML."""
    html = '<div class="pandoc-para" data-indent-twips="600"><p>plain <strong>bold</strong> and <span style="color:#FF0000">red</span> text</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "bold")
    assert _ind_left(p) == "600", "indent dropped on paragraph that contains inline formatting"

    # Find the "bold" run and verify it carries <w:b/>.
    bold_run = None
    red_run = None
    for r in p.iter(f"{{{W_NS}}}r"):
        text = "".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t"))
        if text == "bold":
            bold_run = r
        elif text == "red":
            red_run = r
    assert bold_run is not None, "bold run not found in indented paragraph"
    assert bold_run.find(f".//{{{W_NS}}}b") is not None, "bold run lost <w:b/> after Div handler rewrote the Para"
    assert red_run is not None, "red run not found in indented paragraph"
    color = red_run.find(f".//{{{W_NS}}}color")
    assert color is not None and color.get(f"{{{W_NS}}}val") == "FF0000", "color span lost <w:color val=FF0000/> after Div handler rewrote the Para"


def test_indent_div_without_twips_attribute_is_passthrough():
    """A Div with the class but no data-indent-twips must not be rewritten —
    we degrade to letting pandoc render the inner Para normally so we don't
    emit a <w:p> with malformed <w:ind w:left=""/>."""
    html = '<div class="pandoc-para"><p>no indent</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "no indent")
    assert _ind_left(p) is None, "filter emitted an <w:ind> for a Div with no data-indent-twips"


def test_indent_div_falls_back_when_para_contains_a_link():
    """A <w:hyperlink> needs a relationship registered in
    word/_rels/document.xml.rels — something the Div handler can't reproduce
    when it emits a single raw <w:p>. The handler must detect a Link inside
    the Para and fall back to leaving the Para alone (indent dropped, link
    kept) rather than corrupting the document. See the build_para_w_p
    "graceful degradation" comment in filters/inline_styles.lua.
    """
    html = '<div class="pandoc-para" data-indent-twips="600"><p>see <a href="https://example.com/t">this link</a> please</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")
            rels_xml = zf.read("word/_rels/document.xml.rels")

    doc = ET.fromstring(doc_xml)
    # The hyperlink must survive — losing it would silently corrupt the doc.
    hyperlink = doc.find(f".//{{{W_NS}}}hyperlink")
    assert hyperlink is not None, "Div handler dropped the <w:hyperlink> when falling back — graceful degradation didn't work"
    # And the relationship must still be registered.
    assert b"hyperlink" in rels_xml, "hyperlink relationship missing from document.xml.rels — graceful degradation dropped the rel side-effect"


def test_plain_div_is_left_alone():
    """A Div without the pandoc-para class must pass through unchanged —
    the handler's first guard. Otherwise unrelated <div>s in user content
    would all get rewritten as raw OOXML and lose pandoc's default styling."""
    html = '<div class="some-other-class"><p>just a div</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "just a div")
    assert _ind_left(p) is None, "filter applied an indent to a Div that doesn't have the pandoc-para class"
    # And the Para should still have whatever pStyle pandoc would normally
    # give it (not a raw <w:p> with no pStyle).
    pstyle = p.find(f".//{{{W_NS}}}pStyle")
    assert pstyle is not None, "regular Div content lost its default pStyle — filter erroneously rewrote it as raw OOXML"


# ---------------------------------------------------------------------------
# Paragraph alignment (data-text-align -> <w:jc>)
# ---------------------------------------------------------------------------


def test_align_div_sets_w_jc_val():
    """The canonical Polarion case: centered and right-aligned paragraphs.

    The filter must emit <w:p> with <w:pPr><w:jc w:val="..."/></w:pPr>, where
    CSS center -> "center" and right -> "right".
    """
    html = '<div class="pandoc-para" data-text-align="center"><p>Centered</p></div><div class="pandoc-para" data-text-align="right"><p>Right aligned</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    assert _jc_val(_w_p_with_text(doc, "Centered")) == "center"
    assert _jc_val(_w_p_with_text(doc, "Right aligned")) == "right"


def test_align_justify_maps_to_both():
    """CSS `justify` becomes OOXML `both` — the one keyword whose OOXML
    spelling differs from CSS."""
    html = '<div class="pandoc-para" data-text-align="justify"><p>Justified</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    assert _jc_val(_w_p_with_text(doc, "Justified")) == "both"


def test_indent_and_align_emit_both_in_schema_order():
    """A paragraph wrapper carrying both data attributes must produce a single
    <w:p> whose <w:pPr> has <w:ind> immediately before <w:jc> (CT_PPr schema
    order — Word reorders or drops out-of-order children otherwise)."""
    html = '<div class="pandoc-para" data-indent-twips="600" data-text-align="center"><p>Both</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "Both")
    assert _ind_left(p) == "600"
    assert _jc_val(p) == "center"
    # <w:ind> must come before <w:jc> within <w:pPr>.
    ppr = p.find(f"{{{W_NS}}}pPr")
    assert ppr is not None, "indented+aligned paragraph has no <w:pPr>"
    child_tags = [c.tag.split("}", 1)[-1] for c in ppr]
    assert "ind" in child_tags and "jc" in child_tags, f"missing ind/jc in pPr: {child_tags!r}"
    assert child_tags.index("ind") < child_tags.index("jc"), f"<w:ind> must precede <w:jc> per CT_PPr schema order, got {child_tags!r}"


def test_align_preserves_inline_formatting_via_walk():
    """Nested inline formatting inside an aligned paragraph must survive the
    raw-OOXML rewrite, same as for indent."""
    html = '<div class="pandoc-para" data-text-align="center"><p>plain <strong>bold</strong> text</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "bold")
    assert _jc_val(p) == "center", "alignment dropped on paragraph that contains inline formatting"
    bold_run = next((r for r in p.iter(f"{{{W_NS}}}r") if "".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t")) == "bold"), None)
    assert bold_run is not None and bold_run.find(f".//{{{W_NS}}}b") is not None, "bold run lost <w:b/> after Div handler rewrote the aligned Para"


def test_align_div_falls_back_when_para_contains_a_link():
    """Same graceful-degradation contract as the indent path: a Link inside the
    Para can't be reproduced in a raw <w:p>, so the handler keeps the original
    Para (alignment dropped, link + relationship kept)."""
    html = '<div class="pandoc-para" data-text-align="center"><p>see <a href="https://example.com/t">this link</a> please</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")
            rels_xml = zf.read("word/_rels/document.xml.rels")

    doc = ET.fromstring(doc_xml)
    assert doc.find(f".//{{{W_NS}}}hyperlink") is not None, "Div handler dropped the <w:hyperlink> when falling back on an aligned paragraph"
    assert b"hyperlink" in rels_xml, "hyperlink relationship missing — graceful degradation dropped the rel side-effect"


def test_unknown_align_value_is_rejected():
    """data-text-align is mapped through a fixed allowlist; an unrecognized value
    must not reach <w:jc> and must not corrupt the paragraph."""
    html = '<div class="pandoc-para" data-text-align="bogus"><p>visible text</p></div>'
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "visible text")
    assert _jc_val(p) is None, "filter emitted a <w:jc> for an unmapped data-text-align value"


# ---------------------------------------------------------------------------
# Security: data-indent-twips validation
# ---------------------------------------------------------------------------
#
# The HtmlParagraphPreProcess preprocessor only ever writes integer values into
# data-indent-twips, but an HTTP caller can submit HTML that already contains
# `<div class="pandoc-para" data-indent-twips="…">` with arbitrary content.
# Without validation that value is concatenated straight into a
# <w:ind w:left="..."/> attribute, letting the attacker close the attribute
# early and splice arbitrary OOXML into the document. The filter MUST treat
# this attribute as untrusted input and reject anything that isn't a clean
# non-negative integer.


def test_attribute_injection_in_data_indent_twips_is_rejected():
    """The reviewer's "OOXML attribute injection" finding.

    Attempt to close the <w:ind ...> attribute and inject a fake run plus
    a control character that would surface in extracted text if injection
    succeeded. The filter MUST refuse the value entirely and drop the
    indent rather than splice the attacker's payload into the document.
    """
    payload = '600"/></w:pPr><w:r><w:t>INJECTED_PAYLOAD_marker_xyz</w:t></w:r><w:pPr x="'
    html = f"<div class=\"pandoc-para\" data-indent-twips='{payload}'><p>visible text</p></div>"

    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    # The injected marker must NOT appear anywhere in the rendered document.
    # If parse_twips ever forgets to reject this shape, the marker would
    # show up here because the run we crafted would render as visible text.
    assert b"INJECTED_PAYLOAD_marker_xyz" not in doc_xml, (
        "data-indent-twips injection succeeded — the malicious payload reached document.xml. parse_twips in filters/inline_styles.lua must reject any non-numeric value before formatting it into <w:ind w:left>."
    )

    # And the original paragraph text must still be present (graceful fallback,
    # not "drop everything if anything looks suspicious").
    assert b"visible text" in doc_xml, "graceful fallback failed: malicious data-indent-twips also lost the paragraph's real content"

    # And the resulting <w:p> must not carry an <w:ind w:left> (since we
    # refused the value), so the document remains well-formed.
    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "visible text")
    assert _ind_left(p) is None, "filter emitted an <w:ind w:left> for an invalid twips value"


@pytest.mark.parametrize(
    "bad_value",
    [
        "abc",  # non-numeric
        "12abc",  # numeric prefix only
        "-1",  # negative
        "1.5",  # fractional
        "999999",  # exceeds Word's plausible-indent cap (31680)
        "1e10",  # scientific notation — parseable but vastly out of range
        "",  # empty string
        '600"',  # trailing quote — the classic injection prefix
        '0"/></w:pPr><w:r><w:t>BOOM</w:t></w:r><w:pPr x="',  # full attribute-escape payload
    ],
)
def test_invalid_twips_values_drop_indent_but_keep_content(bad_value: str):
    """Every shape that isn't a clean non-negative integer in range must
    drop the indent and pass the original Para through. Pinned individually
    so each rejection rule is failure-isolated.

    Note: shapes Lua's tonumber accepts (e.g. " 600", "0xff", "+1") are
    intentionally NOT in this list — they convert to clean integers and the
    %d formatter then emits only decimal digits, so they're safe even if
    permissive. The security boundary is "the value reaches OOXML as a
    bounded integer", not "the string looks lexically tidy".
    """
    html = f"<div class=\"pandoc-para\" data-indent-twips='{bad_value}'><p>some content here</p></div>"
    with tempfile.TemporaryDirectory() as tmpdir:
        out = Path(tmpdir) / "out.docx"
        _convert_html_to_docx(html, out)
        with zipfile.ZipFile(out) as zf:
            doc_xml = zf.read("word/document.xml")

    doc = ET.fromstring(doc_xml)
    p = _w_p_with_text(doc, "some content here")
    assert _ind_left(p) is None, f"filter accepted bad twips value {bad_value!r} and emitted an <w:ind w:left>"
    # Defensive: a payload that survived would also surface as text in the
    # document. Catches the case where some future refactor accepts the
    # value but then truncates/sanitizes — we want the marker entirely
    # absent so injection attempts leave no trace.
    assert b"BOOM" not in doc_xml, f"injection payload leaked into document.xml for {bad_value!r}"


def test_valid_integer_twips_values_still_work():
    """Sanity: the validation must not break the happy path. Includes the
    extremes the cap rules in (0 is rejected as no-op, 31680 is the upper
    bound, anything above is dropped)."""
    cases = [("1", "1"), ("600", "600"), ("31680", "31680")]
    for raw, expected in cases:
        html = f'<div class="pandoc-para" data-indent-twips="{raw}"><p>val_{raw}</p></div>'
        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "out.docx"
            _convert_html_to_docx(html, out)
            with zipfile.ZipFile(out) as zf:
                doc_xml = zf.read("word/document.xml")
        doc = ET.fromstring(doc_xml)
        p = _w_p_with_text(doc, f"val_{raw}")
        assert _ind_left(p) == expected, f"valid twips {raw!r} did not produce <w:ind w:left='{expected}'/>"
