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
