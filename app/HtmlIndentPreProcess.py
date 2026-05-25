"""Preserve ``<p style="margin-left: ...">`` indentation for DOCX conversion.

Pandoc's HTML reader drops the ``style`` attribute from ``<p>`` entirely, so
any margin-left applied inline at the paragraph level is lost before any Lua
filter can read it. This preprocessor finds each such ``<p>``, parses the
margin-left value into twips (Word's native unit for ``<w:ind w:left=".."/>``),
and wraps the paragraph in

    <div class="pandoc-indent" data-indent-twips="N"><p ...>...</p></div>

Pandoc preserves both ``class`` and ``data-*`` attributes on ``<div>`` elements
(they survive as the AST's ``Attr`` triple), so the indent value travels
intact to the companion Lua filter (``filters/inline_styles.lua``), which
emits a raw OOXML ``<w:p>`` carrying the matching paragraph properties.

Unit conversion
---------------
1 twip = 1/1440 inch. CSS reference DPI is 96, so 1 px = 1/96 inch = 15 twips.
Absolute units (pt, in, cm, mm, pc) convert directly. em/rem are resolved
against the standard 12pt body size (1 em = 240 twips). Percentages and
unparseable values are skipped — the paragraph passes through with no indent
rather than producing wrong output. Negative or zero indents are dropped for
the same reason.
"""

from __future__ import annotations

import logging
import re

from lxml import etree, html  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)

INDENT_CLASS = "pandoc-indent"
INDENT_ATTR = "data-indent-twips"

# CSS unit -> twips conversion factor. 1 twip = 1/1440 inch.
_UNIT_TO_TWIPS: dict[str, float] = {
    "px": 1440 / 96,
    "pt": 1440 / 72,
    "in": 1440.0,
    "cm": 1440 / 2.54,
    "mm": 1440 / 25.4,
    "pc": 1440 / 6,
    # em/rem are font-relative; without a real cascade we approximate against
    # the default 12pt body size (1 em = 12pt = 240 twips).
    "em": 240.0,
    "rem": 240.0,
}

# A numeric value followed by an optional unit (letters or %). The unit is
# captured separately so percentages can be detected and rejected. No
# whitespace allowed inside — CSS forbids it between number and unit, and the
# caller already strips surrounding whitespace before handing the value over.
# Keeping the pattern free of variable-width whitespace quantifiers also
# silences the SonarCloud S5852 "regex could backtrack" warning on the
# previous `^\s*...\s*...\s*$` form (which was already linear-time but
# matched the rule's "multiple \s* quantifiers" heuristic).
_VALUE_RE = re.compile(r"^([+-]?\d+(?:\.\d+)?)([a-z%]*)$", re.IGNORECASE)

# Exceptions we treat as "input isn't parseable HTML, pass it through".
# Bound to a name (rather than written inline as a tuple literal) because
# `ruff format` rewrites literal except-tuples to PEP 758's parens-free form
# under Python 3.14, which is a SyntaxError on older interpreters and trips
# every static analyzer that hasn't caught up to the new syntax yet. A name
# reference is not rewritten by the formatter, so the parens stay where the
# language requires them.
_PARSE_FAILURES = (etree.ParseError, etree.ParserError, ValueError)


def preprocess(source: bytes) -> bytes:
    """Wrap each indented ``<p>`` in a marker ``<div>``. Idempotent on input
    that has nothing to rewrite — returns the original bytes unchanged.
    """
    try:
        fragments = html.fragments_fromstring(source)
    except _PARSE_FAILURES:
        logger.warning("HtmlIndentPreProcess: HTML parse failed; passing input through")
        return source

    # Top-level <p> fragments have no parent, so we couldn't reparent them
    # with their wrapping <div>. Hang every fragment off a synthetic root
    # while we walk so getparent()/insert() works uniformly — then drop the
    # root on the way out.
    synthetic_root = etree.Element("__pandoc_indent_root__")
    leading_text: str | None = None
    for frag in fragments:
        if hasattr(frag, "tag"):
            synthetic_root.append(frag)
        # fragments_fromstring may emit a leading text node as a plain str.
        elif leading_text is None:
            leading_text = frag
        else:
            leading_text += frag

    rewrote = _wrap_indented_paragraphs(synthetic_root)
    if not rewrote:
        return source

    parts: list[bytes] = []
    if leading_text:
        parts.append(leading_text.encode("utf-8"))
    for child in synthetic_root:
        parts.append(html.tostring(child, encoding="utf-8"))
    return b"".join(parts)


def _wrap_indented_paragraphs(root: html.HtmlElement) -> bool:
    rewrote = False
    # Materialize before mutating: parent.remove/insert invalidates the
    # iterator if we walk lazily.
    for p in list(root.iter("p")):
        style = p.get("style")
        if not style:
            continue
        twips = _extract_margin_left_twips(style)
        if twips is None:
            continue
        parent = p.getparent()
        if parent is None:
            continue
        _wrap_paragraph(parent, p, twips)
        rewrote = True
    return rewrote


def _wrap_paragraph(parent: html.HtmlElement, p: html.HtmlElement, twips: int) -> None:
    idx = parent.index(p)
    div = etree.Element("div")
    div.set("class", INDENT_CLASS)
    div.set(INDENT_ATTR, str(twips))
    parent.remove(p)
    div.append(p)
    parent.insert(idx, div)


def _extract_margin_left_twips(style: str) -> int | None:
    """Return the first ``margin-left`` value in a CSS style string, in twips.

    Implemented with plain string ops rather than a regex because the previous
    pattern (``(?:^|;)\\s*margin-left\\s*:\\s*([^;]+)``) tripped SonarCloud's
    "regex could backtrack" rule on its multiple ``\\s*`` quantifiers. The
    declarations of a CSS style string are semicolon-separated and each is a
    plain ``property: value`` pair, so ``str.split``/``str.partition`` handles
    the parsing in obviously linear time with no analyzer false-positives.

    CSS says later declarations override earlier ones; we keep the original
    regex's leftmost-wins behavior because Polarion never emits duplicates
    and the test suite locks that contract down explicitly.
    """
    for declaration in style.split(";"):
        prop, sep, value = declaration.partition(":")
        if sep and prop.strip().lower() == "margin-left":
            return _css_length_to_twips(value.strip())
    return None


def _css_length_to_twips(value: str) -> int | None:
    """Return positive twips for a CSS length; None for zero, negative,
    percentages, or unparseable input.
    """
    m = _VALUE_RE.match(value)
    if not m:
        return None
    try:
        n = float(m.group(1))
    except ValueError:
        return None
    # CSS allows bare numbers only for unitless properties (e.g. line-height);
    # for margin-left a bare number is invalid, but real-world emitters drop
    # the unit assuming px. Treat missing-unit as px to match common emitters.
    unit = (m.group(2) or "px").lower()
    factor = _UNIT_TO_TWIPS.get(unit)
    if factor is None:
        return None
    twips = round(n * factor)
    return twips if twips > 0 else None
