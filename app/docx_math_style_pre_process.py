r"""Keep the style of upright and bold italic letters in Office math on the DOCX -> LaTeX/PDF path.

texmath reads ``<m:sty m:val="p"/>`` on a letter as the default, so the d of
"dx" comes out italic, and reads ``bi`` as ``\mathbf``, which is upright. Word
sets them upright and bold italic. This wraps the text of such a run in
markers, as :mod:`app.docx_math_color_pre_process` does for a color::

    <m:t>d</m:t>  with m:sty p  ->  <m:t>PMSzzzpzzzdzzzPMSENDzzz</m:t>

and ``filters/docx/docx_math_styles_to_latex.lua`` turns them into ``\symup``
and ``\symbfit``. The ``m:sty`` is removed, so texmath sets the markers as plain
letters.

Only runs of Latin letters are wrapped. A run in another script (``m:scr``),
of normal text (``m:nor``), or with the name of a function texmath already sets
upright (``sin``, ``max``) stays as it is.
"""

from __future__ import annotations

import re

from .docx_ooxml import parse_xml, serialize_tree

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"  # NOSONAR OOXML math namespace identifier (ECMA-376), never dereferenced

_M_R = f"{{{M_NS}}}r"
_M_T = f"{{{M_NS}}}t"
_M_RPR = f"{{{M_NS}}}rPr"
_M_STY = f"{{{M_NS}}}sty"
_M_SCR = f"{{{M_NS}}}scr"
_M_NOR = f"{{{M_NS}}}nor"
_M_VAL = f"{{{M_NS}}}val"

# Shared with the pattern in filters/docx/docx_math_styles_to_latex.lua
MARKER_PREFIX = "PMSzzz"
MARKER_INFIX = "zzz"
MARKER_END = "zzzPMSENDzzz"

# The styles Word sets otherwise than texmath reads them
STYLES = frozenset({"p", "bi"})

_LETTERS = re.compile(r"[A-Za-z]+")

# The names texmath sets as operators of their own
FUNCTIONS = frozenset(
    {
        *("arccos", "arcsin", "arctan", "arg", "cos", "cosh", "cot", "coth", "csc", "deg", "det", "dim", "exp", "gcd", "hom", "inf", "ker"),
        *("lg", "lim", "liminf", "limsup", "ln", "log", "max", "min", "mod", "Pr", "sec", "sin", "sinh", "sup", "tan", "tanh"),
    },
)


def rewrite_part(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Wrap the upright and bold italic letter runs of one body part. Returns (new_bytes, changed)."""
    if b"m:sty" not in xml_bytes:
        return xml_bytes, False
    tree = parse_xml(xml_bytes)
    if tree is None:
        return xml_bytes, False

    changed = False
    for run in tree.iter(_M_R):
        properties = run.find(_M_RPR)
        text = run.find(_M_T)
        if properties is None or text is None or properties.find(_M_SCR) is not None or properties.find(_M_NOR) is not None:
            continue
        style = properties.find(_M_STY)
        value = style.get(_M_VAL) if style is not None else None
        if style is None or value not in STYLES or not _LETTERS.fullmatch(text.text or "") or text.text in FUNCTIONS:
            continue
        text.text = f"{MARKER_PREFIX}{value}{MARKER_INFIX}{text.text}{MARKER_END}"
        properties.remove(style)
        changed = True

    if not changed:
        return xml_bytes, False
    return serialize_tree(tree), True
