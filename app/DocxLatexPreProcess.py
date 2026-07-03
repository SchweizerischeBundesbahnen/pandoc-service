"""Single-pass docx→latex preprocessing.

For LaTeX/PDF targets three independent rewrites run on the source DOCX before
pandoc reads it: colour/size runs (:mod:`app.DocxColorPreProcess`), paragraph
alignment/indent (:mod:`app.DocxParagraphPreProcess`) and list-level tagging
(:mod:`app.DocxListLevelPreProcess`). Run separately they each unzip the whole
package, rewrite their body parts and re-zip — so an image-heavy document gets
its media decompressed and recompressed three times, tripling the peak memory
and CPU of the step.

This module orchestrates the same per-part transforms over a single unzip /
re-zip: the media is held once and the body XML flows colour → paragraph →
list through the existing ``_rewrite_part`` helpers, so the produced DOCX is
byte-for-byte identical to chaining the three ``preprocess`` calls — only much
lighter on memory.
"""

from __future__ import annotations

from . import DocxColorPreProcess, DocxListLevelPreProcess, DocxMathColorPreProcess, DocxParagraphPreProcess, DocxTablePreProcess
from .docx_ooxml import STYLES_PART, augment_styles, enumerate_body_parts, read_entries, repack


def _rewrite_body_part(
    xml: bytes,
    has_styles: bool,
    color_styles: dict[str, DocxColorPreProcess._StyleSpec],
    para_styles: dict[str, DocxParagraphPreProcess._StyleSpec],
) -> tuple[bytes, bool]:
    """Run the colour → paragraph → list → table rewrites over one body part,
    collecting any synthetic styles into the shared dicts.  Returns
    (new_xml, changed)."""
    changed = False
    if has_styles:
        rewritten, color_used = DocxColorPreProcess._rewrite_part(xml)
        if color_used:
            xml, changed = rewritten, True
            color_styles.update(color_used)
        rewritten, para_used = DocxParagraphPreProcess._rewrite_part(xml)
        if para_used:
            xml, changed = rewritten, True
            para_styles.update(para_used)
    rewritten, list_changed = DocxListLevelPreProcess._rewrite_part(xml)
    if list_changed:
        xml, changed = rewritten, True
    rewritten, table_changed = DocxTablePreProcess._rewrite_part(xml)
    if table_changed:
        xml, changed = rewritten, True
    rewritten, math_color_changed = DocxMathColorPreProcess._rewrite_part(xml)
    if math_color_changed:
        xml, changed = rewritten, True
    return xml, changed


def preprocess(docx_bytes: bytes) -> bytes:
    """Apply the colour, paragraph, list-level and table-cell rewrites in one
    unzip/re-zip.

    Equivalent to chaining ``DocxColorPreProcess.preprocess``,
    ``DocxParagraphPreProcess.preprocess``, ``DocxListLevelPreProcess
    .preprocess`` and ``DocxTablePreProcess.preprocess`` but without
    re-zipping the package (and its media) between each step.
    """
    entries = read_entries(docx_bytes)
    if entries is None:
        return docx_bytes

    body_parts = enumerate_body_parts(entries.keys())
    if not body_parts:
        return docx_bytes

    # Colour and paragraph rewrites both register synthetic styles in
    # styles.xml; without it they bail out entirely (matching their standalone
    # preprocess()). List-level tagging needs no styles and always runs.
    has_styles = STYLES_PART in entries
    color_styles: dict[str, DocxColorPreProcess._StyleSpec] = {}
    para_styles: dict[str, DocxParagraphPreProcess._StyleSpec] = {}
    changed = False

    for part in body_parts:
        entries[part], part_changed = _rewrite_body_part(entries[part], has_styles, color_styles, para_styles)
        changed = changed or part_changed

    if not changed:
        return docx_bytes

    if color_styles:
        entries[STYLES_PART] = augment_styles(entries[STYLES_PART], color_styles, DocxColorPreProcess._build_style_element)
    if para_styles:
        entries[STYLES_PART] = augment_styles(entries[STYLES_PART], para_styles, DocxParagraphPreProcess._build_style_element)

    return repack(entries)
