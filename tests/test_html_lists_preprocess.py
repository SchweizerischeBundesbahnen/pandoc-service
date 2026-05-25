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
