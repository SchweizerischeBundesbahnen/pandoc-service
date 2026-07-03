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

# Exceptions we treat as "input isn't parseable HTML, pass it through".
# Bound to a name (rather than written inline as a tuple literal) because
# `ruff format` rewrites literal except-tuples to PEP 758's parens-free form
# under Python 3.14, which is a SyntaxError on older interpreters and trips
# every static analyzer that hasn't caught up to the new syntax yet. A name
# reference is not rewritten by the formatter, so the parens stay where the
# language requires them.
_PARSE_FAILURES = (etree.ParseError, etree.ParserError, ValueError)


def preprocess(source: bytes) -> bytes:
    """Wrap orphan ``<ol>`` / ``<ul>`` children of list elements.

    Returns the (possibly transformed) HTML as bytes. Passes the input
    through unchanged if parsing fails or no orphan lists are found, so
    valid HTML is never modified.
    """
    try:
        # document_fromstring (not fragments_fromstring) so a full HTML document
        # keeps its <head>: the exporter sends <head><title>...</title>, which
        # pandoc renders with the "Title" style. fragments_fromstring drops the
        # head, so re-serializing after wrapping an orphan list would lose the
        # title (it falls back to "First Paragraph"). A bare fragment is
        # harmlessly wrapped in <html><body> (pandoc reads it the same), and we
        # only re-serialize at all when an orphan list was actually wrapped.
        doc = html.document_fromstring(source)
    except _PARSE_FAILURES:
        logger.warning("HtmlListsPreProcess: HTML parse failed; passing input through")
        return source

    if not _wrap_orphan_lists(doc):
        return source

    return html.tostring(doc, encoding="utf-8")


def _wrap_orphan_lists(root: html.HtmlElement) -> bool:
    """Walk ``root`` and wrap every orphan ``<ol>`` / ``<ul>``.

    Returns True if any wrapping was performed.
    """
    rewrote = False
    # We must materialize the iterator into a list before mutating: the tree
    # walk is invalidated by insert/remove. iter() yields the root first if
    # it is also a list element, which is fine — we still check its children.
    for parent in root.iter():
        if parent.tag not in _LIST_TAGS:
            continue
        for child in parent:
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
