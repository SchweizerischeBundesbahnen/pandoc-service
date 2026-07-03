r"""Preserve OMML math-run color across the DOCX -> LaTeX/PDF path.

Pandoc reads Office math (``<m:oMath>``) through the ``texmath`` library
(``readOMML`` -> ``[Exp]`` -> ``writeTeX``). ``texmath``'s expression AST has no
color constructor and its OMML reader ignores ``<w:color>`` on math runs
entirely, so a colored equation reaches the LaTeX/PDF writer black — even though
the color is right there in the DOCX (``<m:r><w:rPr><w:color w:val="RRGGBB"/>``,
which is exactly what ``app/DocxMathColorPostProcess.py`` writes on the
HTML -> DOCX path, and what Word renders).

This preprocessor is the *encode* half of a shim that is the mirror image of the
HTML -> DOCX one (:mod:`app.HtmlMathColorPreProcess` + :mod:`app.DocxMathColorPostProcess`).
For every math run carrying a direct ``<w:color>`` it wraps the run's ``<m:t>``
text in plain-text markers::

    <m:t>x</m:t>   with color RRGGBB   ->   <m:t>PMCzzzRRGGBBzzzxzzzPMCENDzzz</m:t>

The markers are pure alphanumerics, which ``texmath`` emits contiguously into the
TeX string (only operators get spacing), so they survive ``readOMML`` ->
``writeTeX`` intact inside the produced ``Math`` inline. ``filters/docx_math_colors_to_latex.lua``
then rewrites each marker pair in the math string into ``{\color[HTML]{RRGGBB} ...}``,
which ``xcolor`` (already loaded by the pipeline) renders in the PDF.

The direct ``<w:color>`` is stripped once its color is encoded: ``texmath``
ignores it anyway, and removing it keeps the run from being re-detected if the
preprocessor is ever run twice.

Scope / limitations
-------------------
* Only a direct ``<w:color w:val="RRGGBB"/>`` on the run is handled. Theme colors
  (``w:themeColor``) and the literal ``auto`` are skipped (left uncolored), and a
  non 6-hex value is ignored — matching what ``\color[HTML]{...}`` can accept.
* Background shading (``<w:shd>``) and highlight (``<w:highlight>``) inside math
  are not handled — math runs carry only ``<w:color>`` in our pipeline.
* Only meaningful for the DOCX -> LaTeX/PDF targets; wire it in there
  (see :mod:`app.DocxLatexPreProcess`), never for DOCX -> DOCX/HTML/etc.
"""

from __future__ import annotations

import logging

from .docx_ooxml import W_NS, enumerate_body_parts, parse_xml, read_entries, repack, serialize_tree

logger = logging.getLogger(__name__)

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"  # NOSONAR OOXML math namespace identifier (ECMA-376), never dereferenced

_M_R = f"{{{M_NS}}}r"
_M_T = f"{{{M_NS}}}t"
_W_RPR = f"{{{W_NS}}}rPr"
_W_COLOR = f"{{{W_NS}}}color"
_W_VAL = f"{{{W_NS}}}val"

HEX_COLOR_LENGTH = 6
_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")

# Marker sentinels wrapping a colored run's text. Shared, by construction, with
# the regex in filters/docx_math_colors_to_latex.lua:
#     PMCzzz(%x%x%x%x%x%x)zzz(.-)zzzPMCENDzzz  ->  {\color[HTML]{%1} %2}
# Pure alphanumerics on purpose: texmath concatenates letters/digits without
# inserting spacing (only operators get spaced), so "PMCzzzRRGGBBzzz...zzzPMCENDzzz"
# reaches the TeX output byte-for-byte. The "zzz" delimiters plus the PMC/PMCEND
# tokens make an accidental collision with real formula text vanishingly unlikely.
MARKER_PREFIX = "PMCzzz"
MARKER_INFIX = "zzz"
MARKER_END = "zzzPMCENDzzz"


def preprocess(docx_bytes: bytes) -> bytes:
    """Return a DOCX byte-string with colored math runs' text wrapped in color
    markers. Returns the input unchanged when there is no colored math run or the
    package is not a recognizable DOCX.
    """
    entries = read_entries(docx_bytes)
    if entries is None:
        logger.warning("Input is not a valid zip / DOCX; skipping math-color preprocess")
        return docx_bytes

    changed = False
    for part in enumerate_body_parts(entries.keys()):
        rewritten, part_changed = _rewrite_part(entries[part])
        if part_changed:
            entries[part] = rewritten
            changed = True

    # Fast path: no colored math anywhere -> return the original bytes so the
    # zip layout is preserved and no needless re-zip happens.
    if not changed:
        return docx_bytes
    return repack(entries)


def _rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Wrap colored math runs in one body part. Returns (new_bytes, changed)."""
    tree = parse_xml(xml_bytes)
    if tree is None:
        logger.warning("Unparseable XML in DOCX part; skipping math-color preprocess")
        return xml_bytes, False

    changed = False
    # iter() reaches every <m:r> regardless of nesting (fractions, scripts,
    # matrices, ...); a math run is <m:r>, distinct from a text run <w:r>, so
    # this never touches the runs DocxColorPreProcess handles.
    for run in tree.iter(_M_R):
        w_rpr = run.find(_W_RPR)
        if w_rpr is None:
            continue
        color_el = w_rpr.find(_W_COLOR)
        if color_el is None:
            continue
        hex_color = _normalize_hex(color_el.get(_W_VAL))
        if hex_color is None:
            continue
        t_el = run.find(_M_T)
        if t_el is None:
            # No <m:t> to wrap (e.g. a run holding only properties); leave it.
            continue
        t_el.text = f"{MARKER_PREFIX}{hex_color}{MARKER_INFIX}{t_el.text or ''}{MARKER_END}"
        # Strip the now-encoded color so the run isn't re-detected on a re-run.
        w_rpr.remove(color_el)
        changed = True

    if not changed:
        return xml_bytes, False
    return serialize_tree(tree), True


def _normalize_hex(value: str | None) -> str | None:
    """Uppercase a 6-digit hex color, accepting an optional leading '#'. Returns
    None for absent/``auto``/theme/non 6-hex values (the run then stays uncolored,
    which is safer than emitting a wrong ``\\color``)."""
    if not value:
        return None
    stripped = value.strip()
    if stripped.startswith("#"):
        stripped = stripped[1:]
    if len(stripped) == HEX_COLOR_LENGTH and all(c in _HEX_DIGITS for c in stripped):
        return stripped.upper()
    return None
