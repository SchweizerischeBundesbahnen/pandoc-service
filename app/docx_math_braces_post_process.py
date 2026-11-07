r"""Turn a brace under or over an expression into the group character of Word.

texmath writes ``\underbrace{a+b}_{n}`` as a lower limit whose limit is the brace
character alone: ``<m:limLow><m:e>a+b</m:e><m:lim>⏟</m:lim></m:limLow>``. Word draws such
a limit as a small brace under the middle of the expression. Word itself writes an
under or over brace as a group character, ``<m:groupChr>``, which it stretches over the
whole expression.

:func:`apply_group_characters` rewrites each such limit **in place**, on the lxml tree
``docx_post_process`` has already parsed, to the group character Word writes: the brace
below the expression for a lower limit, above it for an upper one.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from lxml import etree

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"  # NOSONAR OOXML math namespace identifier (ECMA-376), never dereferenced
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR OOXML namespace identifier (ECMA-376), never dereferenced

_M = f"{{{M_NS}}}"
_VAL = f"{_M}val"
# The braces and brackets Word stretches as a group character: under the expression, and over it
UNDER = frozenset("⏟⎵⏝")
OVER = frozenset("⏞⎴⏜")
# The limit, its position and the alignment of the group character: below the expression or above it
_LIMITS = {f"{_M}limLow": (UNDER, "bot", "top"), f"{_M}limUpp": (OVER, "top", "bot")}


def apply_group_characters(doc: DocumentObject) -> None:
    """Rewrite each limit which is a brace alone as a group character, in place."""
    # Collected first: the rewrite replaces elements of the tree the walk goes through
    limits = [element for tag in _LIMITS for element in doc.element.iter(tag)]
    for limit in limits:
        _rewrite(limit)


def _brace(limit: etree._Element, braces: frozenset[str]) -> tuple[str, etree._Element] | None:
    """The brace a limit holds and nothing else, with its run, or None."""
    lim = limit.find(f"{_M}lim")
    runs = list(lim) if lim is not None else []
    if len(runs) != 1 or runs[0].tag != f"{_M}r":
        return None
    text = "".join(t.text or "" for t in runs[0].iter(f"{_M}t")).strip()
    return (text, runs[0]) if text in braces else None


def _rewrite(limit: etree._Element) -> None:
    braces, position, alignment = _LIMITS[limit.tag]
    found = _brace(limit, braces)
    expression = limit.find(f"{_M}e")
    if found is None or expression is None:
        return
    brace, run = found
    group = etree.Element(f"{_M}groupChr")
    properties = etree.SubElement(group, f"{_M}groupChrPr")
    etree.SubElement(properties, f"{_M}chr").set(_VAL, brace)
    etree.SubElement(properties, f"{_M}pos").set(_VAL, position)
    etree.SubElement(properties, f"{_M}vertJc").set(_VAL, alignment)
    # The formatting of the run of the brace, its colour among it, goes to the group character
    run_properties = run.find(f"{{{W_NS}}}rPr")
    if run_properties is not None:
        etree.SubElement(properties, f"{_M}ctrlPr").append(copy.deepcopy(run_properties))
    group.append(expression)
    group.tail = limit.tail
    parent = limit.getparent()
    if parent is not None:  # pragma: no branch - a limit of the body always has a parent
        parent.replace(limit, group)
