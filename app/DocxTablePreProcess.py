"""Rewrite table properties in a DOCX so tables survive into LaTeX/PDF.

Pandoc's DOCX reader has three table-related problems this preprocessor fixes:

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

3. **Dropped table width & alignment**: pandoc normalises the column widths
   to sum to 1.0 (losing the table's ``<w:tblW>`` share of the page) and
   discards the table's ``<w:jc>`` alignment entirely, so every table renders
   full-width and centered in the PDF.  This preprocessor records the table's
   width fraction and alignment on the first cell's sentinel; the Lua filter
   scales the ``ColSpec`` widths back down and emits longtable ``\\LTleft``/
   ``\\LTright`` glue.  Every table with an explicit ``<w:tblW>`` is tagged,
   including 100% ones (pandoc may otherwise render them content-width and
   centered when the DOCX carries no usable column widths).

Sentinel format (PUA delimiters U+E010 / U+E011), a ``;``-separated key map::

    <U+E010>bg=RRGGBB;tw=0.4000;ta=left<U+E011>

``bg`` is per coloured cell; ``tw`` (0..1 line fraction) and ``ta``
(left/center/right) are table-level and live on the first cell (merged into
its sentinel if it also carries ``bg``).

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
_TBLPR_TAG = f"{{{W_NS}}}tblPr"
_TBLGRID_TAG = f"{{{W_NS}}}tblGrid"
_GRIDCOL_TAG = f"{{{W_NS}}}gridCol"
_TBLW_TAG = f"{{{W_NS}}}tblW"
_JC_TAG = f"{{{W_NS}}}jc"
_TR_TAG = f"{{{W_NS}}}tr"
_TC_TAG = f"{{{W_NS}}}tc"
_TCPR_TAG = f"{{{W_NS}}}tcPr"
_GRIDSPAN_TAG = f"{{{W_NS}}}gridSpan"
_SHD_TAG = f"{{{W_NS}}}shd"
_P_TAG = f"{{{W_NS}}}p"
_PPR_TAG = f"{{{W_NS}}}pPr"
_PSTYLE_TAG = f"{{{W_NS}}}pStyle"
_R_TAG = f"{{{W_NS}}}r"
_T_TAG = f"{{{W_NS}}}t"
_FLDSIMPLE_TAG = f"{{{W_NS}}}fldSimple"
_INSTRTEXT_TAG = f"{{{W_NS}}}instrText"
_INSTR_ATTR = f"{{{W_NS}}}instr"
_W_ATTR = f"{{{W_NS}}}w"
_TYPE_ATTR = f"{{{W_NS}}}type"
_VAL_ATTR = f"{{{W_NS}}}val"
_FILL_ATTR = f"{{{W_NS}}}fill"
_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

# OOXML table width of 100% == 5000 fiftieths of a percent.
_MAX_PCT = 5000.0
# Reference text width in twips for turning an absolute (dxa) table width into a
# line fraction: Letter (8.5in) minus 1in margins each side = 6.5in = 9360 twips.
# Matches the Letter assumption in app/DocxPostProcess.py and the 468pt used by
# filters/html_tables_to_latex.lua. Best-effort for absolute widths (same caveat).
_REFERENCE_WIDTH_TWIPS = 9360.0
# Map an OOXML <w:jc w:val> to a canonical alignment token.
_JC_TO_ALIGN = {"left": "left", "center": "center", "right": "right", "start": "left", "end": "right"}

# Paragraph style id that HTML->DOCX assigns to Polarion captions (see
# filters/html_captions.lua). Their text already carries the sequence number
# ("Table 1 ..."), so if pandoc's DOCX reader turns them into a LaTeX
# \caption it prepends its OWN "Table N:" counter, printing the number twice.
_CAPTION_STYLE_VAL = "Caption"

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


def _build_sentinel_text(kv: dict[str, str]) -> str:
    """Encode a property map as a ``<OPEN>k1=v1;k2=v2<CLOSE>`` sentinel string."""
    payload = ";".join(f"{key}={value}" for key, value in kv.items())
    return f"{SENTINEL_OPEN}{payload}{SENTINEL_CLOSE}"


def _parse_sentinel_text(text: str) -> tuple[dict[str, str], str]:
    """Split a leading sentinel into ``({k: v}, rest)``.

    Returns ``({}, text)`` when *text* does not start with a complete sentinel.
    """
    if not text.startswith(SENTINEL_OPEN):
        return {}, text
    close = text.find(SENTINEL_CLOSE)
    if close == -1:
        return {}, text
    payload = text[len(SENTINEL_OPEN) : close]
    rest = text[close + len(SENTINEL_CLOSE) :]
    kv: dict[str, str] = {}
    for segment in payload.split(";"):
        key, sep, value = segment.partition("=")
        if sep:
            kv[key] = value
    return kv, rest


def _first_run_text_element(para: ET.Element) -> ET.Element | None:
    """Return the ``<w:t>`` of the paragraph's first run (skipping ``<w:pPr>``),
    or None when the paragraph does not begin with a run.
    """
    for child in para:
        if child.tag == _R_TAG:
            return child.find(_T_TAG)
        if child.tag != _PPR_TAG:
            return None
    return None


def _make_sentinel_run(sentinel_text: str) -> ET.Element:
    """Build a ``<w:r><w:t>sentinel</w:t></w:r>`` element."""
    run = ET.Element(_R_TAG)
    text = ET.SubElement(run, _T_TAG, {_SPACE_ATTR: "preserve"})
    text.text = sentinel_text
    return run


def _is_positive_int(value: str | None) -> bool:
    """True when *value* is a positive integer string (e.g. a ``w:w`` width)."""
    if value is None:
        return False
    stripped = value.strip()
    return stripped.isdigit() and int(stripped) > 0


def _row_column_count(tr: ET.Element) -> int:
    """Number of grid columns a row spans (summing each cell's ``w:gridSpan``)."""
    total = 0
    for tc in tr.findall(_TC_TAG):
        span = 1
        tcpr = tc.find(_TCPR_TAG)
        if tcpr is not None:
            gridspan = tcpr.find(_GRIDSPAN_TAG)
            if gridspan is not None and _is_positive_int(gridspan.get(_VAL_ATTR)):
                span = int(gridspan.get(_VAL_ATTR))  # type: ignore[arg-type]
        total += span
    return total


def _table_column_count(tbl: ET.Element) -> int:
    """Widest row's column count — how many ``<w:gridCol>`` the table needs."""
    return max((_row_column_count(tr) for tr in tbl.findall(_TR_TAG)), default=0)


def _fix_grid_col_widths(tbl: ET.Element) -> bool:
    """Ensure the table has a ``<w:tblGrid>`` with one positive-width
    ``<w:gridCol>`` per column.

    Pandoc's DOCX reader derives ``ColSpec`` proportions from the grid columns.
    If the grid is missing, has too few columns, or carries missing/zero/invalid
    widths (some editors emit ``w:w="0"``), pandoc ends up with an empty or
    partial column list and the LaTeX writer renders the table content-width
    and centered — ignoring its ``<w:tblW>``. This normalises the grid so every
    column has a positive width (the ratios are all pandoc uses), which keeps
    the cells and lets the width/alignment recovery place the table correctly.

    Returns True when the grid was created or modified.
    """
    num_cols = _table_column_count(tbl)
    if num_cols == 0:
        return False

    changed = False
    tblgrid = tbl.find(_TBLGRID_TAG)
    if tblgrid is None:
        tblgrid = ET.Element(_TBLGRID_TAG)
        tblpr = tbl.find(_TBLPR_TAG)
        insert_at = list(tbl).index(tblpr) + 1 if tblpr is not None else 0
        tbl.insert(insert_at, tblgrid)
        changed = True

    grid_cols = tblgrid.findall(_GRIDCOL_TAG)
    for col in grid_cols:
        if not _is_positive_int(col.get(_W_ATTR)):
            col.set(_W_ATTR, _DEFAULT_GRIDCOL_WIDTH)
            changed = True

    # Add any missing columns so pandoc sees the full width, not a subset.
    for _ in range(len(grid_cols), num_cols):
        col = ET.SubElement(tblgrid, _GRIDCOL_TAG)
        col.set(_W_ATTR, _DEFAULT_GRIDCOL_WIDTH)
        changed = True

    return changed


def _find_or_create_first_para(tc: ET.Element, tcpr: ET.Element | None) -> ET.Element:
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


def _ensure_sentinel(para: ET.Element, kv: dict[str, str]) -> None:
    """Merge *kv* into the paragraph's leading sentinel, creating one if absent.

    Merging (rather than prepending a second sentinel) keeps a single parseable
    ``<OPEN>...<CLOSE>`` at the very start, so table-level keys (``tw``/``ta``)
    and a cell background (``bg``) can share one sentinel on the first cell.
    Idempotent: re-applying the same keys leaves the text unchanged.
    """
    text_el = _first_run_text_element(para)
    if text_el is not None and text_el.text and text_el.text.startswith(SENTINEL_OPEN):
        existing, rest = _parse_sentinel_text(text_el.text)
        existing.update(kv)
        text_el.text = _build_sentinel_text(existing) + rest
        return

    ppr = para.find(_PPR_TAG)
    insert_at = 1 if ppr is not None and next(iter(para), None) is ppr else 0
    para.insert(insert_at, _make_sentinel_run(_build_sentinel_text(kv)))


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
        text_el = _first_run_text_element(first_para)
        if text_el is not None and text_el.text:
            existing, _ = _parse_sentinel_text(text_el.text)
            if existing.get("bg") == bg:
                continue  # already tagged with this background (idempotency)

        _ensure_sentinel(first_para, {"bg": bg})
        changed = True

    return changed


def _extract_table_layout(tbl: ET.Element) -> tuple[float | None, str | None, bool]:
    """Return ``(width_fraction, alignment, is_absolute)`` from ``<w:tblPr>``.

    ``width_fraction`` is the table's share of the line width: ``pct`` -> value
    / 5000; ``dxa`` -> twips / reference text width; both clamped to 1.0. It is
    None when there is no usable ``<w:tblW>``. ``alignment`` is left/center/right
    from ``<w:jc>``, or None. ``is_absolute`` is True when the width came from a
    ``dxa`` (absolute) value — such tables are usually deliberately narrow, so
    the filter tightens their inter-column padding to keep them close to the
    requested size. These are the properties pandoc's DOCX reader discards (it
    keeps only column-relative widths and no table jc).
    """
    tblpr = tbl.find(_TBLPR_TAG)
    if tblpr is None:
        return None, None, False

    fraction: float | None = None
    is_absolute = False
    tblw = tblpr.find(_TBLW_TAG)
    if tblw is not None:
        raw = tblw.get(_W_ATTR)
        try:
            value = float(raw) if raw is not None else 0.0
        except ValueError:
            value = 0.0
        if value > 0:
            width_type = tblw.get(_TYPE_ATTR)
            if width_type == "pct":
                fraction = min(value / _MAX_PCT, 1.0)
            elif width_type == "dxa":
                fraction = min(value / _REFERENCE_WIDTH_TWIPS, 1.0)
                is_absolute = True

    align: str | None = None
    jc = tblpr.find(_JC_TAG)
    if jc is not None:
        align = _JC_TO_ALIGN.get((jc.get(_VAL_ATTR) or "").lower())

    return fraction, align, is_absolute


def _first_own_cell(tbl: ET.Element) -> ET.Element | None:
    """Return this table's first cell (first ``<w:tc>`` of its first ``<w:tr>``),
    ignoring cells of nested tables.
    """
    for child in tbl:
        if child.tag == _TR_TAG:
            for cell in child:
                if cell.tag == _TC_TAG:
                    return cell
            return None
    return None


def _tag_table_layout(tbl: ET.Element) -> bool:
    """Tag the first cell with the table's width fraction and/or alignment.

    Every table with an explicit ``<w:tblW>`` is tagged, **including 100 %**:
    pandoc's DOCX reader may still emit a table with no usable column widths,
    which the LaTeX writer then renders content-width (and, with no alignment
    glue, centered) instead of filling the line. The companion filter forces
    the recovered fraction onto the columns and re-applies the alignment, so a
    100 % table fills the text width flush-left rather than floating centered.
    (This matches how the HTML path in filters/html_tables_to_latex.lua already
    handles full-width tables.) Returns True when a sentinel was written.
    """
    fraction, align, is_absolute = _extract_table_layout(tbl)
    kv: dict[str, str] = {}
    if fraction is not None:
        kv["tw"] = f"{min(fraction, 1.0):.4f}"
        if is_absolute:
            # Absolute (px/pt) widths are usually deliberately narrow; flag them
            # so the filter tightens inter-column padding, which otherwise adds
            # a fixed minimum that leaves a small table wider than requested.
            kv["aw"] = "1"
    if align is not None:
        kv["ta"] = align
    if not kv:
        return False

    first_tc = _first_own_cell(tbl)
    if first_tc is None:
        return False

    first_para = _find_or_create_first_para(first_tc, first_tc.find(_TCPR_TAG))
    _ensure_sentinel(first_para, kv)
    return True


def _has_sequence_field(para: ET.Element) -> bool:
    """True when the paragraph numbers itself with a Word ``SEQ`` field.

    A genuine Word caption carries its number as a ``SEQ`` field, so pandoc can
    strip that label and re-number it (keeping it in a List of Tables/Figures).
    A Polarion caption instead has the number as literal text, so there is no
    field to find.
    """
    if any("SEQ" in (fld.get(_INSTR_ATTR) or "").upper() for fld in para.iter(_FLDSIMPLE_TAG)):
        return True
    return any(instr.text and "SEQ" in instr.text.upper() for instr in para.iter(_INSTRTEXT_TAG))


def _neutralize_caption_paragraphs(tree: ET.Element) -> bool:
    """Strip the ``Caption`` style from *Polarion* captions so pandoc does not
    turn them into auto-numbered LaTeX ``\\caption`` blocks.

    A Polarion caption's text already contains its number as literal text
    ("Table 1 ..."); left as a ``Caption``-styled paragraph, pandoc's DOCX
    reader attaches it to the adjacent table and LaTeX prints
    "Table N: Table N ..." — the counter twice. Dropping the style makes it a
    normal paragraph that renders its text once (matching how the DOCX shows it).

    A genuine Word caption is left untouched: it numbers itself with a ``SEQ``
    field, so pandoc re-numbers it correctly and can still list it in a
    generated List of Tables/Figures. We therefore only neutralise captions
    that have no ``SEQ`` field (i.e. the literal-numbered Polarion ones), and
    never pandoc's own ``TableCaption``/``ImageCaption`` styles.

    Returns True when any paragraph was changed.
    """
    changed = False
    for para in tree.iter(_P_TAG):
        ppr = para.find(_PPR_TAG)
        if ppr is None:
            continue
        pstyle = ppr.find(_PSTYLE_TAG)
        if pstyle is not None and pstyle.get(_VAL_ATTR) == _CAPTION_STYLE_VAL and not _has_sequence_field(para):
            ppr.remove(pstyle)
            changed = True
    return changed


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Fix grid-column widths, tag styled table cells, carry table width/
    alignment, and neutralise caption styles in one body part.

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
        if _tag_table_layout(tbl):
            changed = True
        if _tag_cell_backgrounds(tbl):
            changed = True

    if _neutralize_caption_paragraphs(tree):
        changed = True

    if not changed:
        return xml_bytes, False

    return serialize_tree(tree), True
