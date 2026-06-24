"""Shared OOXML/zip plumbing for the DOCX preprocessors.

``DocxColorPreProcess``, ``DocxParagraphPreProcess`` and
``DocxListLevelPreProcess`` all rewrite body parts of a DOCX package before
pandoc reads it. This module factors out the boilerplate they share: the
WordprocessingML namespace, the canonical-prefix registration, the set of parts
that may contain runs/paragraphs, and reading/repacking the zip.
"""

from __future__ import annotations

import io
import logging
import zipfile
from typing import TYPE_CHECKING, Protocol
from xml.etree import ElementTree as ET

from defusedxml import ElementTree as DefusedET

if TYPE_CHECKING:
    from collections.abc import Callable, Iterable

logger = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
STYLES_PART = "word/styles.xml"

# Canonical OOXML prefixes. ElementTree mints synthetic prefixes (ns0, ns1, …)
# for namespaces it wasn't told about, which makes pandoc's docx reader drop
# every <w:drawing> (images vanish). Registering the canonical prefixes keeps
# them on serialize.
_OOXML_NAMESPACES = {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",  # NOSONAR False positive - URI is OOXML namespace identifier (ECMA-376), it's never dereferenced
}
for _prefix, _uri in _OOXML_NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)

# Parts that can contain <w:r> runs / <w:p> paragraphs to rewrite. Headers and
# footers live in numbered parts (word/header1.xml, word/footer2.xml, …) whose
# count depends on the section layout, so they're matched by prefix + suffix.
_FIXED_BODY_PARTS = frozenset(
    {
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    }
)


def enumerate_body_parts(names: Iterable[str]) -> list[str]:
    """Return zip entry names that may contain runs/paragraphs to rewrite."""
    return [name for name in names if name in _FIXED_BODY_PARTS or (name.startswith(("word/header", "word/footer")) and name.endswith(".xml"))]


def read_entries(docx_bytes: bytes) -> dict[str, bytes] | None:
    """Read every zip entry into memory, or None if the bytes aren't a zip."""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
            return {name: zin.read(name) for name in zin.namelist()}
    except zipfile.BadZipFile:
        return None


def repack(entries: dict[str, bytes]) -> bytes:
    """Re-zip the (possibly mutated) entry dict back into a DOCX byte string."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    return buf.getvalue()


def parse_xml(xml_bytes: bytes) -> ET.Element | None:
    """Parse a body/styles part with defusedxml (blocks XXE / billion-laughs),
    returning None on malformed XML so callers can skip it. We serialize with
    stdlib ET (:func:`serialize_tree`) since defusedxml has no tostring.
    """
    try:
        return DefusedET.fromstring(xml_bytes)
    except ET.ParseError:
        return None


def serialize_tree(tree: ET.Element) -> bytes:
    return ET.tostring(tree, xml_declaration=True, encoding="UTF-8")


class _StyleLike(Protocol):
    style_id: str


def augment_styles[S: _StyleLike](styles_xml: bytes, specs: dict[str, S], build_style_element: Callable[[S], ET.Element]) -> bytes:
    """Append a ``<w:style>`` element (built by ``build_style_element``) for each
    spec to ``word/styles.xml``, idempotently. Never appends a second style with
    an existing ``w:styleId`` (Word rejects duplicate styleIds); style order in
    styles.xml has no semantic meaning, so appending to the root is sufficient.
    Returns the input unchanged when styles.xml is unparseable.
    """
    tree = parse_xml(styles_xml)
    if tree is None:
        logger.warning("Unparseable %s; skipping style augmentation", STYLES_PART)
        return styles_xml

    existing_ids = {el.get(f"{{{W_NS}}}styleId") for el in tree.findall(f"{{{W_NS}}}style")}
    for spec in specs.values():
        if spec.style_id in existing_ids:
            continue
        tree.append(build_style_element(spec))

    return serialize_tree(tree)
