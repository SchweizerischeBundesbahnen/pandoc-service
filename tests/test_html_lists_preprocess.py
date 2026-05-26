"""Unit tests for ``app.HtmlListsPreProcess``.

Each test feeds a snippet of HTML into ``preprocess`` and asserts on the
returned bytes. We work directly on the string output (rather than re-parsing
into a tree) so the assertions document the exact contract the companion Lua
filter relies on: every orphan ``<ol>``/``<ul>`` is wrapped in a ``<li>`` that
starts with ``<span class="pandoc-suppress-marker">``.
"""

from __future__ import annotations

from app import HtmlListsPreProcess


# --- helpers --------------------------------------------------------------

SENTINEL = b'<span class="pandoc-suppress-marker"></span>'


def _count_sentinels(out: bytes) -> int:
    return out.count(SENTINEL)


# --- happy path -----------------------------------------------------------


def test_orphan_ol_inside_ol_is_wrapped():
    """The canonical Polarion case: <ol><ol><li/></ol></ol> gets a sentinel."""
    src = b"<ol><ol><li>x</li></ol></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    # The orphan inner <ol> must now be inside an <li> with the sentinel as
    # its first child.
    assert b"<li>" + SENTINEL + b"<ol>" in out
    assert _count_sentinels(out) == 1


def test_orphan_ul_inside_ol_is_wrapped():
    """Heterogeneous nesting (<ol> hosts an orphan <ul>) is handled the same."""
    src = b"<ol><ul><li>x</li></ul></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert b"<li>" + SENTINEL + b"<ul>" in out


def test_orphan_ol_inside_ul_is_wrapped():
    src = b"<ul><ol><li>x</li></ol></ul>"
    out = HtmlListsPreProcess.preprocess(src)
    assert b"<li>" + SENTINEL + b"<ol>" in out


def test_user_provided_polarion_snippet_round_trip():
    """End-to-end fidelity check on the exact shape Polarion exports."""
    src = b'<ol id="polarion_1"><li>Level 1<ol><ol><li>Level 3</li></ol><li>Level 2</li></ol></li></ol>'
    out = HtmlListsPreProcess.preprocess(src)
    # Only the inner orphan <ol> (the one wrapping Level 3) should have been
    # wrapped — not the outer well-formed <ol id="polarion_1">.
    assert _count_sentinels(out) == 1
    assert b'<ol id="polarion_1">' in out  # id preserved on outer
    assert b"Level 1" in out and b"Level 2" in out and b"Level 3" in out


# --- nested orphans -------------------------------------------------------


def test_multiple_levels_of_orphan_nesting_each_get_a_sentinel():
    """<ol><ol><ol><li>X</li></ol></ol></ol> — two orphans, two sentinels."""
    src = b"<ol><ol><ol><li>x</li></ol></ol></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert _count_sentinels(out) == 2


# --- pass-through ---------------------------------------------------------


def test_well_formed_lists_pass_through_unchanged():
    """Valid HTML must not be altered — exact byte equality."""
    src = b"<ol><li>A<ol><li>A.1</li><li>A.2</li></ol></li><li>B</li></ol><ul><li>X<ul><li>X.1</li></ul></li></ul>"
    out = HtmlListsPreProcess.preprocess(src)
    assert out == src
    assert _count_sentinels(out) == 0


def test_empty_input_is_returned_unchanged():
    assert HtmlListsPreProcess.preprocess(b"") == b""


def test_input_without_lists_is_returned_unchanged():
    src = b"<p>hello</p><p>world</p>"
    assert HtmlListsPreProcess.preprocess(src) == src


def test_orphan_li_inside_ol_is_not_treated_as_a_list():
    """Only <ol>/<ul> children trigger wrapping; <li> children are valid."""
    src = b"<ol><li>just a normal item</li></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert out == src
    assert _count_sentinels(out) == 0


# --- edge cases -----------------------------------------------------------


def test_attributes_on_orphan_list_are_preserved():
    """When we move the orphan into its sentinel <li>, its attributes stay."""
    src = b'<ol><ol id="inner" class="nested" start="5"><li>x</li></ol></ol>'
    out = HtmlListsPreProcess.preprocess(src)
    # The wrapping <li> introduces a sentinel and then the original <ol>
    # tag with its attributes intact follows.
    assert b'id="inner"' in out
    assert b'class="nested"' in out
    assert b'start="5"' in out


def test_text_between_orphan_lists_is_preserved():
    """Stray text/inline content next to an orphan list must not be dropped."""
    src = b"<ol>before<ol><li>x</li></ol>after</ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert b"before" in out
    assert b"after" in out
    assert _count_sentinels(out) == 1


def test_unparseable_input_passes_through():
    """Garbage that can't be parsed as HTML must not crash the conversion —
    the preprocessor must return the original bytes."""
    # Empty fragments_fromstring is forgiving, but a raw "<" with nothing
    # following is the canonical degenerate input.
    src = b"<"
    out = HtmlListsPreProcess.preprocess(src)
    # We don't care what it produces as long as it doesn't raise and returns
    # bytes — the public contract is "best-effort, never explode".
    assert isinstance(out, bytes)


# --- parse failure --------------------------------------------------------
#
# lxml's HTML parser is intentionally forgiving, so we patch the symbol the
# preprocessor calls to force the exception branch. The contract is the same
# as in HtmlIndentPreProcess: any caught parser exception must return the
# original source object unchanged (identity, not equality) so callers can
# detect "no work done" without re-parsing.


def test_parse_failure_passes_input_through(mocker):
    """If lxml raises, we MUST return the input bytes unchanged so the
    conversion pipeline doesn't 500 on malformed HTML."""
    mocker.patch(
        "app.HtmlListsPreProcess.html.fragments_fromstring",
        side_effect=ValueError("synthetic parse failure"),
    )
    src = b"<ol><ol><li>x</li></ol></ol>"
    assert HtmlListsPreProcess.preprocess(src) is src


def test_parse_failure_logs_warning(mocker, caplog):
    """The parse-failure path must emit a WARNING so failures are diagnosable
    in production — locked down so an operator's log alert keeps firing."""
    import logging  # noqa: PLC0415

    mocker.patch(
        "app.HtmlListsPreProcess.html.fragments_fromstring",
        side_effect=ValueError("boom"),
    )
    caplog.set_level(logging.WARNING, logger="app.HtmlListsPreProcess")
    HtmlListsPreProcess.preprocess(b"<ol><li>x</li></ol>")
    assert any("HtmlListsPreProcess" in rec.message for rec in caplog.records), f"expected a WARNING from HtmlListsPreProcess; got: {[r.message for r in caplog.records]!r}"


# --- leading text round-trip ----------------------------------------------
#
# fragments_fromstring may emit a leading str (any text before the first
# element). When at least one orphan list is wrapped, that text must be
# re-emitted verbatim ahead of the elements.


def test_leading_text_before_orphan_list_is_preserved_when_rewriting():
    """A leading text node must survive the round-trip when an orphan list
    triggers rewriting. Without this, we'd silently lose text in front of
    Polarion-style fragments that begin with prose."""
    src = b"preamble text <ol><ol><li>x</li></ol></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert out.startswith(b"preamble text"), f"leading text not preserved: {out!r}"
    assert SENTINEL in out  # rewriting did happen


def test_leading_text_alone_does_not_force_rewrite():
    """Leading text without any wrap-target must return the source unchanged
    (identity), so we don't pay the round-trip cost for inputs we have no
    work to do on."""
    src = b"just some text <ol><li>plain</li></ol>"
    assert HtmlListsPreProcess.preprocess(src) is src


# --- module-level contract ------------------------------------------------


def test_module_constants_match_documented_contract():
    """The sentinel class name is hard-coded in filters/html_lists.lua —
    pin it so a rename here doesn't silently break the Lua side."""
    assert HtmlListsPreProcess.SUPPRESS_MARKER_CLASS == "pandoc-suppress-marker"


# --- additional structural cases ------------------------------------------


def test_orphan_at_top_of_outer_list_is_wrapped():
    """Orphan as the very first child of its parent — boundary check."""
    src = b"<ol><ol><li>first orphan</li></ol><li>then real</li></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert _count_sentinels(out) == 1
    assert out.index(b"first orphan") < out.index(b"then real")


def test_orphan_at_bottom_of_outer_list_is_wrapped():
    """Orphan as the very last child — boundary check on the other side."""
    src = b"<ol><li>real first</li><ol><li>then orphan</li></ol></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert _count_sentinels(out) == 1
    assert out.index(b"real first") < out.index(b"then orphan")


def test_sibling_orphans_each_get_their_own_wrapper():
    """Two orphans side by side both get wrapped, with one sentinel each."""
    src = b"<ol><ol><li>A</li></ol><ol><li>B</li></ol></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert _count_sentinels(out) == 2
    # Each orphan must keep its content; order preserved.
    assert out.index(b">A<") < out.index(b">B<")


def test_top_level_ol_with_orphan_child_is_wrapped():
    """An <ol> that is itself the top-level fragment (no surrounding context)
    must still have its orphan child wrapped — covers the
    iter()-includes-root case in _wrap_orphan_lists."""
    src = b"<ol><ol><li>only one</li></ol></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    assert _count_sentinels(out) == 1
    assert b"only one" in out


def test_well_formed_li_wrapping_a_list_is_not_touched():
    """A valid <li> that legitimately contains a nested list (i.e. an empty
    parent item) must NOT be turned into a sentinel — the filter would then
    strip the marker the user actually wants. Documents the asymmetry: only
    orphan <ol>/<ul> children trigger wrapping."""
    src = b"<ol><li><ol><li>nested</li></ol></li></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    # No sentinel added — the existing <li> is well-formed.
    assert _count_sentinels(out) == 0


def test_deeply_nested_mixed_orphans():
    """A pathological case combining <ol> and <ul> orphans at multiple
    depths — every orphan still gets its own sentinel."""
    src = b"<ol><ul><ol><li>x</li></ol></ul></ol>"
    out = HtmlListsPreProcess.preprocess(src)
    # Two orphans: <ul> inside <ol>, and <ol> inside <ul>.
    assert _count_sentinels(out) == 2


def test_full_html_document_input_is_handled():
    """fragments_fromstring also accepts a full document — make sure we
    don't break when the input includes <html>/<body> wrappers."""
    src = b'<html><body><ol id="polarion_1"><li>Level 1<ol><ol><li>Level 3</li></ol><li>Level 2</li></ol></li></ol></body></html>'
    out = HtmlListsPreProcess.preprocess(src)
    assert _count_sentinels(out) == 1
    assert b"Level 1" in out and b"Level 2" in out and b"Level 3" in out
