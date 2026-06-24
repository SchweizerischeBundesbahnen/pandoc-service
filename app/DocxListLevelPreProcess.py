"""Tag list-paragraph indent levels in a DOCX so they survive into LaTeX/PDF.

Polarion allows malformed lists where a deeper level is nested directly inside
a shallower one with no intermediate item (e.g. a level-3 ``<ol>`` straight
inside a level-1 list). The DOCX is fine — each paragraph carries its absolute
``<w:numPr>/<w:ilvl>`` and Word indents by level — but pandoc's DOCX reader
**flattens skipped levels** when it reconstructs nested lists, so a level-3
item with no level-2 item before it collapses to the same nesting depth (and
therefore the same indentation) as level 2 in the PDF output.

The true level lives only in ``<w:ilvl>`` and is gone once pandoc has read the
file, so this preprocessor runs *before* pandoc and prepends a sentinel run to
every list paragraph encoding its level::

    <ilvl>

The sentinel survives into the AST as a leading ``Str`` on the list item. The
companion Lua filter (``filters/docx_lists_to_latex.lua``) reads the level,
strips the sentinel, and pushes any under-nested sublist to its intended depth
with marker-less wrapper levels. The Private-Use-Area delimiters never collide
with real document text.

This is the list-level companion to :mod:`app.DocxColorPreProcess` /
:mod:`app.DocxParagraphPreProcess`; it runs on the same docx→latex path and is
independent of them (list paragraphs carry ``<w:numPr>``, not colour/jc/ind).
"""

from __future__ import annotations

import logging
from xml.etree import ElementTree as ET

from defusedxml import ElementTree as DefusedET

from .docx_ooxml import W_NS, enumerate_body_parts, read_entries, repack

logger = logging.getLogger(__name__)

# Private-Use-Area delimiters wrapping the decimal level, e.g. "2".
# The Lua filter matches the same shape. PUA code points never appear in real
# document text, so the sentinel is unambiguous and safe to strip.
SENTINEL_OPEN = ""
SENTINEL_CLOSE = ""

_P_TAG = f"{{{W_NS}}}p"
_PPR_TAG = f"{{{W_NS}}}pPr"
_NUMPR_TAG = f"{{{W_NS}}}numPr"
_ILVL_TAG = f"{{{W_NS}}}ilvl"
_R_TAG = f"{{{W_NS}}}r"
_T_TAG = f"{{{W_NS}}}t"
_VAL_ATTR = f"{{{W_NS}}}val"
_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX with each list paragraph's ``<w:ilvl>`` prepended as a
    sentinel run. Returns the input unchanged when there are no list paragraphs
    or the package is not a recognizable DOCX.
    """
    entries = read_entries(docx_bytes)
    if entries is None:
        logger.warning("Input is not a valid zip / DOCX; skipping list-level preprocess")
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


def _list_level(ppr: ET.Element) -> int | None:
    """Return the ``<w:ilvl>`` value of a list paragraph, or None when the
    paragraph is not a numbered/bulleted list item.
    """
    numpr = ppr.find(_NUMPR_TAG)
    if numpr is None:
        return None
    ilvl = numpr.find(_ILVL_TAG)
    # A list paragraph without an explicit <w:ilvl> defaults to level 0.
    if ilvl is None:
        return 0
    raw = ilvl.get(_VAL_ATTR)
    if raw is None:
        return 0
    try:
        value = int(raw)
    except ValueError:
        return None
    return value if value >= 0 else None


def _already_tagged(para: ET.Element) -> bool:
    """True if the paragraph's first run already carries a sentinel (so a second
    preprocess pass is a no-op)."""
    for child in para:
        if child.tag == _R_TAG:
            text_el = child.find(_T_TAG)
            text = text_el.text if text_el is not None else None
            return bool(text and text.startswith(SENTINEL_OPEN))
        if child.tag != _PPR_TAG:
            # Some other leading element (e.g. bookmark) — not our sentinel run.
            return False
    return False


def _make_sentinel_run(level: int) -> ET.Element:
    run = ET.Element(_R_TAG)
    text = ET.SubElement(run, _T_TAG, {_SPACE_ATTR: "preserve"})
    text.text = f"{SENTINEL_OPEN}{level}{SENTINEL_CLOSE}"
    return run


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Tag every list paragraph in one body part. Returns (new_bytes, changed)."""
    try:
        tree = DefusedET.fromstring(xml_bytes)
    except ET.ParseError:
        logger.warning("Unparseable XML in DOCX part; skipping")
        return xml_bytes, False

    changed = False
    for para in tree.iter(_P_TAG):
        ppr = para.find(_PPR_TAG)
        if ppr is None:
            continue
        level = _list_level(ppr)
        if level is None:
            continue
        if _already_tagged(para):
            # Idempotent: a previous pass already prepended a sentinel run.
            continue
        # Insert the sentinel run as the first run of the paragraph, after the
        # <w:pPr> (which must stay the first child of <w:p>). list(para) keeps
        # document order; pPr is index 0 when present.
        insert_at = 1 if list(para)[0] is ppr else 0
        para.insert(insert_at, _make_sentinel_run(level))
        changed = True

    if not changed:
        return xml_bytes, False

    return ET.tostring(tree, xml_declaration=True, encoding="UTF-8"), True
