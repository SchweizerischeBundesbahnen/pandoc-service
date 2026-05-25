"""Wrap orphan ``<ol>`` / ``<ul>`` nested directly inside another list.

Some HTML emitters (notably Polarion) produce non-standard list markup where
an ``<ol>`` or ``<ul>`` appears as a *direct* child of another list, with no
wrapping ``<li>``. Browsers and CSS-counter renderers (e.g. WeasyPrint when
producing PDF) handle this by creating an anonymous list item that has no
``::marker``: the deeper list keeps its intended depth and no stray marker is
drawn.

Pandoc's HTML reader, on the other hand, synthesizes an implicit list item
around the orphan list to make the AST well-formed. The DOCX writer then
emits a numbered (but empty) paragraph for that synthetic item, which Word
renders as a stray marker (e.g. ``a.``) above the deeper item.

This preprocessor finds each orphan ``<ol>`` / ``<ul>`` and wraps it in an
explicit ``<li>``, prepended with a sentinel
``<span class="pandoc-suppress-marker"></span>``. The companion Lua filter
(``filters/html_lists.lua``) detects the sentinel after pandoc has parsed
the HTML and rewrites the list item so the DOCX writer skips emitting a
numbered paragraph for it.

The transformation is only applied for ``html -> docx`` conversions; other
targets are unaffected.
"""

from __future__ import annotations

import logging

from lxml import etree, html  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)

SUPPRESS_MARKER_CLASS = "pandoc-suppress-marker"
_LIST_TAGS = frozenset({"ol", "ul"})


def preprocess(source: bytes) -> bytes:
    """Wrap orphan ``<ol>`` / ``<ul>`` children of list elements.

    Returns the (possibly transformed) HTML as bytes. Passes the input
    through unchanged if parsing fails or no orphan lists are found, so
    valid HTML is never modified.
    """
    try:
        # fragments_fromstring tolerates both full documents and bare
        # fragments. We get a list of HtmlElements (and possibly a leading
        # text string) covering everything in the input.
        fragments = html.fragments_fromstring(source)
    except etree.ParseError, etree.ParserError, ValueError:
        logger.warning("HtmlListsPreProcess: HTML parse failed; passing input through")
        return source

    rewrote = False
    for frag in fragments:
        # Plain leading text is returned as a str; skip those.
        if not hasattr(frag, "iter"):
            continue
        rewrote = _wrap_orphan_lists(frag) or rewrote

    if not rewrote:
        return source

    # Serialize each fragment and concatenate. encoding=str returns a Python
    # string we then re-encode as UTF-8 (matching what pandoc expects when
    # we hand it the temp file).
    parts: list[bytes] = []
    for frag in fragments:
        if hasattr(frag, "tag"):
            parts.append(html.tostring(frag, encoding="utf-8"))
        else:
            parts.append(frag.encode("utf-8"))
    return b"".join(parts)


def _wrap_orphan_lists(root: html.HtmlElement) -> bool:
    """Walk ``root`` and wrap every orphan ``<ol>`` / ``<ul>``.

    Returns True if any wrapping was performed.
    """
    rewrote = False
    # We must materialize the iterator into a list before mutating: the tree
    # walk is invalidated by insert/remove. iter() yields the root first if
    # it is also a list element, which is fine — we still check its children.
    for parent in list(root.iter()):
        if parent.tag not in _LIST_TAGS:
            continue
        for child in list(parent):
            if child.tag in _LIST_TAGS:
                _wrap_in_marker_li(parent, child)
                rewrote = True
    return rewrote


def _wrap_in_marker_li(parent: html.HtmlElement, orphan: html.HtmlElement) -> None:
    """Replace ``orphan`` in ``parent`` with a ``<li>`` wrapping it.

    The new ``<li>`` carries a sentinel ``<span>`` so the Lua filter can
    identify items that came from orphan lists (rather than from valid
    markup that happens to contain an empty list item).
    """
    idx = parent.index(orphan)
    li = etree.Element("li")
    sentinel = etree.SubElement(li, "span")
    sentinel.set("class", SUPPRESS_MARKER_CLASS)
    parent.remove(orphan)
    li.append(orphan)
    parent.insert(idx, li)
