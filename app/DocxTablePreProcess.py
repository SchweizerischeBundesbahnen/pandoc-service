"""Rewrite table properties in a DOCX so tables survive into LaTeX/PDF.

Pandoc's DOCX reader has two table-related problems this preprocessor fixes:

1. **Missing grid-column widths**: ``<w:gridCol/>`` elements without a
   ``w:w`` attribute cause pandoc to produce an empty ``ColSpec`` list,
   which in turn drops every cell in the table.  This preprocessor adds
   equal-width ``w:w`` attributes when they are missing so pandoc can
   compute column proportions and preserve cell content.

2. **Dropped cell formatting**: pandoc strips cell-level properties
   (background shading, borders, vertical alignment) before building the
   AST.  This preprocessor prepends a PUA-delimited sentinel to the first
   paragraph of each coloured cell so the companion Lua filter
   (``filters/docx_tables_to_latex.lua``) can inject ``\\cellcolor`` into
   the LaTeX output.

Sentinel format (PUA delimiters U+E010 / U+E011)::

    <U+E010>bg=RRGGBB<U+E011>

This is the table companion to :mod:`app.DocxColorPreProcess` /
:mod:`app.DocxParagraphPreProcess`; it runs on the same docx->latex path
and is independent of them (table cells carry ``<w:tcPr>``, not
run/paragraph properties).
"""

from __future__ import annotations

import logging
from xml.etree import ElementTree as ET

from .docx_ooxml import W_NS, enumerate_body_parts, parse_xml, read_entries, repack, serialize_tree

logger = logging.getLogger(__name__)

# PUA delimiters — different from list-level (\uE000/\uE001) to avoid
# ambiguity when a list paragraph happens to sit inside a table cell.
SENTINEL_OPEN = "\ue010"
SENTINEL_CLOSE = "\ue011"

_TBL_TAG = f"{{{W_NS}}}tbl"
_TBLGRID_TAG = f"{{{W_NS}}}tblGrid"
_GRIDCOL_TAG = f"{{{W_NS}}}gridCol"
_TC_TAG = f"{{{W_NS}}}tc"
_TCPR_TAG = f"{{{W_NS}}}tcPr"
_SHD_TAG = f"{{{W_NS}}}shd"
_P_TAG = f"{{{W_NS}}}p"
_PPR_TAG = f"{{{W_NS}}}pPr"
_R_TAG = f"{{{W_NS}}}r"
_T_TAG = f"{{{W_NS}}}t"
_W_ATTR = f"{{{W_NS}}}w"
_FILL_ATTR = f"{{{W_NS}}}fill"
_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

# Default column width (twips) assigned when <w:gridCol> lacks a w:w
# attribute.  The exact value is unimportant — pandoc only uses the
# *ratios* between columns — so we pick a round number that gives a
# plausible-looking layout (~3.3 inches per column).
_DEFAULT_GRIDCOL_WIDTH = "4800"

HEX_COLOR_LENGTH = 6
_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX with table grid-column widths filled in and each
    styled table cell's background prepended as a sentinel run.

    Returns the input unchanged when there are no tables to fix or the
    package is not a recognisable DOCX.
    """
    entries = read_entries(docx_bytes)
    if entries is None:
        logger.warning("Input is not a valid zip / DOCX; skipping table preprocess")
        return docx_bytes

    changed = False
    for part in enumerate_body_parts(entries.keys()):
        rewritten, part_changed = _rewrite_part(entries[part])
        if part_changed:
            entries[part] = rewritten
            changed = True

    if not changed:
        return docx_bytes

    return repack(entries)


def _normalize_hex(value: str | None) -> str | None:
    """Uppercase a 6-digit hex string, or None for missing / unusable values.

    White (``FFFFFF``) is treated as "no background" because it is the page
    colour and wrapping every cell in ``\\cellcolor`` would be wasteful.
    """
    if not value:
        return None
    stripped = value.strip()
    if stripped.lower() == "auto":
        return None
    if stripped.startswith("#"):
        stripped = stripped[1:]
    if len(stripped) == HEX_COLOR_LENGTH and all(c in _HEX_DIGITS for c in stripped):
        upper = stripped.upper()
        if upper == "FFFFFF":
            return None
        return upper
    return None


def _extract_cell_bg(tcpr: ET.Element) -> str | None:
    """Return the hex fill colour from ``<w:tcPr>/<w:shd>``, or None."""
    shd = tcpr.find(_SHD_TAG)
    if shd is None:
        return None
    return _normalize_hex(shd.get(_FILL_ATTR))


def _build_sentinel_text(bg: str) -> str:
    """Encode cell properties as a sentinel string."""
    return f"{SENTINEL_OPEN}bg={bg}{SENTINEL_CLOSE}"


def _already_tagged(para: ET.Element) -> bool:
    """True when the paragraph's first run already carries a table-cell
    sentinel (idempotency guard for repeated preprocessing passes).
    """
    for child in para:
        if child.tag == _R_TAG:
            text_el = child.find(_T_TAG)
            text = text_el.text if text_el is not None else None
            return bool(text and text.startswith(SENTINEL_OPEN))
        if child.tag != _PPR_TAG:
            return False
    return False


def _make_sentinel_run(sentinel_text: str) -> ET.Element:
    """Build a ``<w:r><w:t>sentinel</w:t></w:r>`` element."""
    run = ET.Element(_R_TAG)
    text = ET.SubElement(run, _T_TAG, {_SPACE_ATTR: "preserve"})
    text.text = sentinel_text
    return run


def _fix_grid_col_widths(tbl: ET.Element) -> bool:
    """Add ``w:w`` to ``<w:gridCol>`` elements that lack it.

    Pandoc's DOCX reader needs ``w:w`` to compute ``ColSpec`` proportions;
    without it the column list is empty and every cell is silently dropped.

    Returns True when any ``<w:gridCol>`` was modified.
    """
    tblgrid = tbl.find(_TBLGRID_TAG)
    if tblgrid is None:
        return False

    grid_cols = tblgrid.findall(_GRIDCOL_TAG)
    if not grid_cols:
        return False

    changed = False
    for col in grid_cols:
        if col.get(_W_ATTR) is None:
            col.set(_W_ATTR, _DEFAULT_GRIDCOL_WIDTH)
            changed = True

    return changed


def _find_or_create_first_para(tc: ET.Element, tcpr: ET.Element) -> ET.Element:
    """Return the first ``<w:p>`` before any nested ``<w:tbl>`` in *tc*.

    A cell containing a nested table has structure ``[tcPr, tbl, p]`` where the
    trailing ``<w:p/>`` is the mandatory cell-mark paragraph.  Injecting the
    sentinel there would place it after the ``Table`` AST node, so we look for
    a ``<w:p>`` that *precedes* any ``<w:tbl>``.  If none exists, a new empty
    paragraph is inserted right after ``<w:tcPr>``.
    """
    for child in tc:
        if child.tag == _TBL_TAG:
            break
        if child.tag == _P_TAG:
            return child
    para = ET.Element(_P_TAG)
    insert_pos = 1 if next(iter(tc), None) is tcpr else 0
    tc.insert(insert_pos, para)
    return para


def _inject_sentinel(para: ET.Element, bg: str) -> None:
    """Insert a sentinel run encoding *bg* into *para*."""
    ppr = para.find(_PPR_TAG)
    insert_at = 1 if ppr is not None and next(iter(para), None) is ppr else 0
    para.insert(insert_at, _make_sentinel_run(_build_sentinel_text(bg)))


def _tag_cell_backgrounds(tbl: ET.Element) -> bool:
    """Prepend sentinels encoding background colour to styled cells.

    Returns True when any cell was modified.
    """
    changed = False

    for tc in tbl.iter(_TC_TAG):
        tcpr = tc.find(_TCPR_TAG)
        if tcpr is None:
            continue

        bg = _extract_cell_bg(tcpr)
        if not bg:
            continue

        first_para = _find_or_create_first_para(tc, tcpr)
        if _already_tagged(first_para):
            continue

        _inject_sentinel(first_para, bg)
        changed = True

    return changed


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Fix grid-column widths and tag styled table cells in one body part.

    Returns ``(new_bytes, changed)``.
    """
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping table preprocess")
        return xml_bytes, False

    changed = False

    for tbl in tree.iter(_TBL_TAG):
        if _fix_grid_col_widths(tbl):
            changed = True
        if _tag_cell_backgrounds(tbl):
            changed = True

    if not changed:
        return xml_bytes, False

    return serialize_tree(tree), True
