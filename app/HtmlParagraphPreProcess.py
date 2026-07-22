"""Preserve paragraph-level ``<p style="...">`` formatting for DOCX conversion.

Pandoc's HTML reader drops the ``style`` attribute from ``<p>`` entirely, so
any paragraph-level CSS applied inline is lost before any Lua filter can read
it. This preprocessor finds each such ``<p>``, extracts the formatting we
support — ``margin-left`` (indent) and ``text-align`` (justification) — and
wraps the paragraph in a marker ``<div>``:

    <div class="pandoc-para" data-indent-twips="N" data-text-align="center"><p ...>...</p></div>

Each ``data-*`` attribute is emitted only when the corresponding property is
present, so an indent-only, align-only, or combined paragraph all share the
same wrapper. Pandoc preserves both ``class`` and ``data-*`` attributes on
``<div>`` elements (they survive as the AST's ``Attr`` triple), so the values
travel intact to the companion Lua filter (``filters/inline_styles.lua``),
which emits a raw OOXML ``<w:p>`` carrying the matching paragraph properties
(``<w:ind w:left="N"/>`` and/or ``<w:jc w:val="..."/>``). Pandoc strips the
``data-`` prefix off these attributes when they aren't reserved HTML attribute
names, so the filter sees ``indent-twips`` and ``text-align`` — note the align
attribute is deliberately ``data-text-align`` (not ``data-align``) because
``align`` *is* a reserved HTML attribute and pandoc would keep it prefixed,
breaking the symmetry with ``indent-twips``.

Indent unit conversion
----------------------
1 twip = 1/1440 inch. CSS reference DPI is 96, so 1 px = 1/96 inch = 15 twips.
Absolute units (pt, in, cm, mm, pc) convert directly. em/rem are resolved
against the standard 12pt body size (1 em = 240 twips). Percentages and
unparseable values are skipped — the paragraph passes through with no indent
rather than producing wrong output. Negative or zero indents are dropped for
the same reason.

Alignment
---------
The CSS ``text-align`` keyword is mapped to a canonical token
(``left``/``center``/``right``/``justify``); ``start``/``end`` are folded onto
``left``/``right``. The canonical token (not the OOXML value) is written into
``data-align`` so the OOXML trust boundary stays in the Lua filter, mirroring
how the twips value is handled. Anything else (``inherit``, ``initial``,
vendor-prefixed, unknown) is skipped.
"""

from __future__ import annotations

import logging
import re

from lxml import etree, html  # type: ignore[import-untyped]

logger = logging.getLogger(__name__)

PARA_CLASS = "pandoc-para"
INDENT_ATTR = "data-indent-twips"
# Deliberately "data-text-align", not "data-align": pandoc strips the "data-"
# prefix only when the remainder isn't a reserved HTML attribute. "align" is
# reserved (pandoc would keep "data-align"), but "text-align" is not, so it
# de-prefixes to "text-align" just like "indent-twips" — keeping the Lua-side
# attribute lookups symmetric. See the module docstring.
ALIGN_ATTR = "data-text-align"

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

# CSS text-align keyword -> canonical token written into data-align. The
# logical keywords start/end fold onto left/right (we have no bidi context, so
# treat the document as left-to-right). Anything not listed here is skipped,
# so inherit/initial/unset/vendor-prefixed values never produce a wrapper.
_ALIGN_MAP: dict[str, str] = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "justify",
    "start": "left",
    "end": "right",
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
    """Wrap each formatted ``<p>`` in a marker ``<div>``. Idempotent on input
    that has nothing to rewrite — returns the original bytes unchanged.
    """
    try:
        # document_fromstring (not fragments_fromstring) so a full HTML document
        # keeps its <head>: the exporter sends <head><title>...</title>, which
        # pandoc renders with the "Title" style. fragments_fromstring drops the
        # head, so re-serializing after wrapping a paragraph would lose the title
        # (it falls back to "First Paragraph"). It also gives every <p> a real
        # parent (<body>), so getparent()/insert() work without a synthetic root.
        # A bare fragment is harmlessly wrapped in <html><body>, and we only
        # re-serialize at all when a paragraph was actually wrapped.
        doc = html.document_fromstring(source)
    except _PARSE_FAILURES:
        logger.warning("HtmlParagraphPreProcess: HTML parse failed; passing input through")
        return source

    if not _wrap_formatted_paragraphs(doc):
        return source

    return html.tostring(doc, encoding="utf-8")


def _wrap_formatted_paragraphs(root: html.HtmlElement) -> bool:
    rewrote = False
    # Materialize before mutating: parent.remove/insert invalidates the
    # iterator if we walk lazily.
    for p in root.iter("p"):
        style = p.get("style")
        if not style:
            continue
        twips = _extract_margin_left_twips(style)
        align = _extract_text_align(style)
        if twips is None and align is None:
            continue
        parent = p.getparent()
        if parent is None:
            continue
        _wrap_paragraph(parent, p, twips, align)
        rewrote = True
    return rewrote


def _wrap_paragraph(parent: html.HtmlElement, p: html.HtmlElement, twips: int | None, align: str | None) -> None:
    idx = parent.index(p)
    div = etree.Element("div")
    div.set("class", PARA_CLASS)
    if twips is not None:
        div.set(INDENT_ATTR, str(twips))
    if align is not None:
        div.set(ALIGN_ATTR, align)
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


def _extract_text_align(style: str) -> str | None:
    """Return the first ``text-align`` value in a CSS style string as a
    canonical token (``left``/``center``/``right``/``justify``), or None when
    the property is absent or its value isn't one we map.

    Parsing mirrors :func:`_extract_margin_left_twips` — plain
    ``str.split``/``str.partition`` over the semicolon-separated declarations,
    leftmost-wins — so the two extractors stay consistent.
    """
    for declaration in style.split(";"):
        prop, sep, value = declaration.partition(":")
        if sep and prop.strip().lower() == "text-align":
            return _ALIGN_MAP.get(value.strip().lower())
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
