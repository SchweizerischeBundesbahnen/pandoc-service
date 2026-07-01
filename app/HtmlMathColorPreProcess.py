r"""Preserve LaTeX math color (``\color`` / ``\textcolor``) across pandoc.

Pandoc converts the LaTeX inside ``<script type="math/tex">`` to native Word
equations (OMML) through its ``texmath`` library. ``texmath`` discards ``\color``
(it parses the formula but drops the color), cannot parse ``\textcolor`` at all
(the whole formula then leaks as ``$$...$$`` text), and its OMML writer has no
color output — so a colored math run reaches Word black, or not at all.

This preprocessor is the *encode* half of a two-step shim (the *decode* half is
``app/DocxMathColorPostProcess.py``). For every math script it rewrites

    \color{NAME}{X}   \textcolor[HTML]{RRGGBB}{X}

into

    \text{@@PMC:RRGGBB@@}X\text{@@PMCEND@@}

The ``\text{...}`` markers are plain text that ``texmath`` keeps as distinct OMML
runs (verified: each ``\text{}`` becomes one ``<m:r><m:t>...</m:t></m:r>``), with
the formerly-colored content sitting as separate runs between the start and end
markers in document order. ``DocxMathColorPostProcess`` then walks the OMML runs,
turns each start/end marker pair into a ``<w:color>`` on the runs between them,
and deletes the marker runs.

Why encode here and not just leave the color to ``texmath``: ``texmath`` has no
path to colored OMML (reader drops it, writer can't emit it), so the intent must
be carried across the conversion as text that survives, then re-applied to the
OOXML afterwards.

Scope: the two color macros MathJax's color extension actually defines are
``\color`` and ``\textcolor`` (not ``\mathcolor``, which is a KaTeX command MathJax
never emits). Only their two-argument form ``\cmd{color}{content}`` is transformed
— in Polarion's MathJax 2.7.9 ``\color`` is the two-argument macro that colors its
argument. The one-argument switch form ``\color{name}`` (color applies to the rest
of the group, the standard-LaTeX/MathJax-v3 behavior) is left untouched — that is
``texmath``'s existing behavior for ``\color`` (content survives, color dropped). A
color we cannot resolve to a hex value unwraps to its bare content (dropping the
color) rather than being left in place, so ``\textcolor`` — which would otherwise
leak the whole formula — still renders.

Not handled: ``\colorbox`` / ``\fcolorbox`` (background/border shading, a different
OMML feature than run ``<w:color>``) leak or lose their box.
"""

from __future__ import annotations

import logging
import re

logger = logging.getLogger(__name__)

# Marker sentinels. Shared with app/DocxMathColorPostProcess.py, which parses them
# back out. Chosen to (a) survive texmath intact inside \text{} as a single run and
# (b) be vanishingly unlikely to collide with real formula text. The start marker
# carries the resolved 6-hex color; the end marker is generic (the decoder uses a
# color stack, so nested colors pop correctly without the end marker naming a color).
MARKER_PREFIX = "@@PMC:"
MARKER_SUFFIX = "@@"
MARKER_END = "@@PMCEND@@"

# The math scripts Polarion emits: <script type="math/tex; mode=display">LATEX</script>.
# We locate them with a regex and rewrite only their body, leaving the rest of the
# HTML byte-for-byte unchanged (safer than a full HTML round-trip, which could
# normalize unrelated markup or escape the "<"/">" that appears in math like a<b).
# The body is captured non-greedily up to the closing tag; LaTeX never contains
# "</script>". Both quote styles and a trailing "; mode=display" are tolerated.
_MATH_SCRIPT_RE = re.compile(
    r"""(<script\b[^>]*\btype\s*=\s*["']math/tex[^"']*["'][^>]*>)(.*?)(</script>)""",
    re.IGNORECASE | re.DOTALL,
)

# The color macros MathJax's color extension defines that take a (color, content)
# argument pair. \color's one-argument switch form (standard LaTeX / MathJax v3) is
# handled by the "no second brace group -> leave untouched" fallback. \mathcolor is
# deliberately absent: it is a KaTeX command, not a MathJax one, so Polarion never
# emits it.
_COLOR_COMMANDS = frozenset({"color", "textcolor"})

# CSS/xcolor color name -> 6-hex (uppercase). Keys are lowercased on lookup, so
# "Red"/"RED"/"red" all resolve. Covers the CSS basic/extended names and the
# dvips-style capitalized names Polarion authors commonly type; an unlisted name
# resolves to None and the content renders uncolored (never leaks).
_COLOR_NAMES: dict[str, str] = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "008000",
    "lime": "00FF00",
    "blue": "0000FF",
    "yellow": "FFFF00",
    "cyan": "00FFFF",
    "aqua": "00FFFF",
    "magenta": "FF00FF",
    "fuchsia": "FF00FF",
    "gray": "808080",
    "grey": "808080",
    "silver": "C0C0C0",
    "lightgray": "D3D3D3",
    "lightgrey": "D3D3D3",
    "darkgray": "A9A9A9",
    "darkgrey": "A9A9A9",
    "maroon": "800000",
    "olive": "808000",
    "purple": "800080",
    "teal": "008080",
    "navy": "000080",
    "orange": "FFA500",
    "pink": "FFC0CB",
    "brown": "A52A2A",
    "violet": "EE82EE",
    "gold": "FFD700",
    "indigo": "4B0082",
    "darkred": "8B0000",
    "darkgreen": "006400",
    "darkblue": "00008B",
    "lightblue": "ADD8E6",
    "darkorange": "FF8C00",
}

_HEX6_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
_HEX3_RE = re.compile(r"^[0-9A-Fa-f]{3}$")


def preprocess(source: bytes) -> bytes:
    """Rewrite ``\\color``/``\\textcolor`` inside math scripts into color markers.
    Idempotent on input with no color to rewrite — returns the original bytes
    unchanged (so the byte-preserving guarantee holds for the common case).
    """
    try:
        text = source.decode("utf-8")
    except UnicodeDecodeError:
        logger.warning("HtmlMathColorPreProcess: input is not valid UTF-8; passing through unchanged")
        return source

    # Cheap guard: nothing to do unless a color command appears somewhere. ("\\color"
    # is not a substring of "\\textcolor", so both spellings must be checked.)
    if "\\color" not in text and "\\textcolor" not in text:
        return source

    changed = False

    def replace_script(match: re.Match[str]) -> str:
        nonlocal changed
        open_tag, body, close_tag = match.group(1), match.group(2), match.group(3)
        new_body = _rewrite_math_colors(body)
        if new_body != body:
            changed = True
        return open_tag + new_body + close_tag

    result = _MATH_SCRIPT_RE.sub(replace_script, text)
    if not changed:
        return source
    return result.encode("utf-8")


def _rewrite_math_colors(latex: str) -> str:
    """Rewrite the color commands in one LaTeX string, recursing into their content."""
    out: list[str] = []
    i = 0
    n = len(latex)
    while i < n:
        char = latex[i]
        if char != "\\":
            out.append(char)
            i += 1
            continue

        name, name_end = _read_control_word(latex, i)
        if name in _COLOR_COMMANDS:
            parsed = _parse_color_command(latex, name_end)
            if parsed is not None:
                hex_color, content, end = parsed
                inner = _rewrite_math_colors(content)  # nested colors convert too
                if hex_color is not None:
                    out.append("\\text{" + MARKER_PREFIX + hex_color + MARKER_SUFFIX + "}")
                    out.append(inner)
                    out.append("\\text{" + MARKER_END + "}")
                else:
                    # Color we cannot resolve: keep the content (uncolored) rather than
                    # leave the command in place, so \textcolor does not leak.
                    out.append(inner)
                i = end
                continue

        # Not a color command (or its arguments were malformed): copy the control
        # sequence verbatim. A control word is backslash + letters; a control symbol
        # (backslash + single non-letter, e.g. "\{") is copied as its two characters.
        if name_end > i + 1:
            out.append(latex[i:name_end])
            i = name_end
        else:
            out.append(latex[i : i + 2])
            i += 2
    return "".join(out)


def _read_control_word(latex: str, backslash_index: int) -> tuple[str, int]:
    """Return the control-word name after the backslash at ``backslash_index`` and the
    index just past it. An empty name means a control symbol (backslash + non-letter).
    """
    j = backslash_index + 1
    while j < len(latex) and latex[j].isascii() and latex[j].isalpha():
        j += 1
    return latex[backslash_index + 1 : j], j


def _parse_color_command(latex: str, pos: int) -> tuple[str | None, str, int] | None:
    """Parse the ``[model]{color}{content}`` arguments starting at ``pos`` (just past the
    command name). Returns ``(hex_or_None, content, end_index)``, or ``None`` when the
    two-argument brace form is absent (switch form or malformed) so the caller leaves
    the command untouched. ``hex_or_None`` is ``None`` when the color cannot be resolved.
    """
    i = _skip_ws(latex, pos)
    model: str | None = None
    if i < len(latex) and latex[i] == "[":
        close = latex.find("]", i)
        if close == -1:
            return None
        model = latex[i + 1 : close].strip()
        i = _skip_ws(latex, close + 1)

    color_value = _read_brace_group(latex, i)
    if color_value is None:
        return None
    i = _skip_ws(latex, color_value[1])

    content = _read_brace_group(latex, i)
    if content is None:
        return None

    return _resolve_color(model, color_value[0].strip()), content[0], content[1]


def _read_brace_group(latex: str, i: int) -> tuple[str, int] | None:
    """If ``latex[i]`` opens a brace group, return ``(inner_text, index_past_close)``;
    otherwise ``None``.
    """
    if i >= len(latex) or latex[i] != "{":
        return None
    close = _find_matching_brace(latex, i)
    if close == -1:
        return None
    return latex[i + 1 : close], close + 1


def _find_matching_brace(latex: str, open_index: int) -> int:
    """Index of the ``}`` matching the ``{`` at ``open_index`` (honoring nesting and
    skipping escaped characters), or ``-1`` if unbalanced.
    """
    depth = 0
    i = open_index
    n = len(latex)
    while i < n:
        char = latex[i]
        if char == "\\":
            i += 2  # skip the escaped character
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def _skip_ws(latex: str, i: int) -> int:
    while i < len(latex) and latex[i].isspace():
        i += 1
    return i


def _resolve_color(model: str | None, value: str) -> str | None:
    """Resolve a color argument to an uppercase 6-hex string, or ``None`` if unknown.

    Supported: the default model (a named color, ``#RGB``/``#RRGGBB`` or bare
    ``RRGGBB`` hex) and ``[HTML]{RRGGBB}``. Other xcolor models (rgb, cmyk, gray)
    are not resolved — the content then renders uncolored.
    """
    if model is None or model == "":
        return _resolve_hex(value) or _COLOR_NAMES.get(value.lower())
    if model.lower() == "html":
        return _resolve_hex(value)
    return None


def _resolve_hex(value: str) -> str | None:
    """Normalize ``#RRGGBB``, ``#RGB``, ``RRGGBB`` or ``RGB`` to uppercase ``RRGGBB``."""
    candidate = value[1:] if value.startswith("#") else value
    if _HEX6_RE.match(candidate):
        return candidate.upper()
    if _HEX3_RE.match(candidate):
        return "".join(component * 2 for component in candidate).upper()
    return None
