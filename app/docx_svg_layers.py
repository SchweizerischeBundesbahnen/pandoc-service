"""
Put the vector SVG of a diagram on top of its PNG in a DOCX.

In the weasysvg mode the DOCX gets each diagram as a PNG, as in the chromium mode, so the writer of
pandoc and :mod:`app.docx_post_process` size it as before. :func:`attach` then finds each picture whose
PNG weasysvg made, adds the flat SVG of that PNG as a part of its own, and names it in an ``asvg:svgBlip``
extension of the picture. This is how Word itself saves an SVG: Word from 2016 draws the SVG, an older
Word and other readers draw the PNG.

pandoc could write the SVG itself, but it makes the PNG for an older Word only without ``--sandbox``,
which the service keeps on.
"""

from __future__ import annotations

import hashlib
import io
import logging
import posixpath
import re
import zipfile

from lxml import etree

logger = logging.getLogger(__name__)

_CONTENT_TYPES = "[Content_Types].xml"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_OFFICE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
# The extension that carries the SVG of a picture, as Word writes it
_SVG_EXTENSION_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
_SVG_CONTENT_TYPE = "image/svg+xml"

# The relationships of a part: word/_rels/document.xml.rels belongs to word/document.xml
_RELS_RE = re.compile(r"(?P<folder>.*?)_rels/(?P<part>[^/]+)\.rels")
# lxml resolves no entity and reads no external DTD of an untrusted part
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)


def attach(docx: bytes, svg_layers: dict[bytes, bytes]) -> bytes:
    """
    Add the flat SVG of each PNG of `svg_layers` to the pictures of `docx` that show that PNG.

    `svg_layers` maps the SHA-256 of a PNG to its flat SVG. A DOCX without such a picture comes back as it
    is; a DOCX that does not read comes back as it is too, with a warning.
    """
    if not svg_layers:
        return docx
    try:
        return _attach(docx, svg_layers)
    # The SVG is an addition: the document with its PNG is still whole without it.
    except (zipfile.BadZipFile, etree.XMLSyntaxError, KeyError) as e:
        logger.warning("Could not add the SVG of the diagrams to the DOCX: %s", e)
        return docx


def _attach(docx: bytes, svg_layers: dict[bytes, bytes]) -> bytes:
    with zipfile.ZipFile(io.BytesIO(docx)) as source:
        names = set(source.namelist())
        changed: dict[str, bytes] = {}
        added: dict[str, bytes] = {}
        # One SVG part per PNG, shared by every part that shows it
        svg_parts: dict[bytes, str] = {}

        for match in sorted(filter(None, map(_RELS_RE.fullmatch, names)), key=lambda found: found.string):
            rels_name = match.string
            part_name = match["folder"] + match["part"]
            if part_name not in names:
                continue
            rels = etree.fromstring(source.read(rels_name), _PARSER)
            svg_ids = _add_svg_relationships(source, names, match["folder"], rels, svg_layers, svg_parts, added)
            if not svg_ids:
                continue
            part = etree.fromstring(source.read(part_name), _PARSER)
            if _add_svg_blips(part, svg_ids):
                changed[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
                changed[part_name] = etree.tostring(part, xml_declaration=True, encoding="UTF-8", standalone=True)

        if not changed:
            return docx
        changed[_CONTENT_TYPES] = _with_svg_content_type(source.read(_CONTENT_TYPES))

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                target.writestr(info, changed.get(info.filename, source.read(info)))
            for name, data in added.items():
                target.writestr(name, data)
    return out.getvalue()


def _add_svg_relationships(
    source: zipfile.ZipFile,
    names: set[str],
    folder: str,
    rels: etree._Element,
    svg_layers: dict[bytes, bytes],
    svg_parts: dict[bytes, str],
    added: dict[str, bytes],
) -> dict[str, str]:
    """Relate each PNG of `svg_layers` this part shows to its SVG. Returns the id of the SVG by the id of the PNG."""
    ids = {relationship.get("Id") for relationship in rels}
    svg_ids: dict[str, str] = {}
    # A copy: the loop adds relationships to rels
    for relationship in rels.findall(f"{{{_RELATIONSHIPS_NS}}}Relationship"):
        if relationship.get("Type") != _IMAGE_RELATIONSHIP or relationship.get("TargetMode") == "External":
            continue
        png_name = posixpath.normpath(posixpath.join(folder, relationship.get("Target", "")))
        if png_name not in names:
            continue
        key = hashlib.sha256(source.read(png_name)).digest()
        if key not in svg_layers:
            continue
        if key not in svg_parts:
            svg_parts[key] = _free_name(f"{posixpath.splitext(png_name)[0]}.svg", names)
            names.add(svg_parts[key])
            added[svg_parts[key]] = svg_layers[key]
        svg_id = _free_id(ids)
        ids.add(svg_id)
        etree.SubElement(rels, f"{{{_RELATIONSHIPS_NS}}}Relationship", Id=svg_id, Type=_IMAGE_RELATIONSHIP, Target=posixpath.relpath(svg_parts[key], folder))
        svg_ids[relationship.get("Id", "")] = svg_id
    return svg_ids


def _add_svg_blips(part: etree._Element, svg_ids: dict[str, str]) -> bool:
    """Name the SVG in each blip of a related PNG that has none yet. Returns whether one changed."""
    embed = f"{{{_OFFICE_RELATIONSHIPS_NS}}}embed"
    changed = False
    for blip in part.iter(f"{{{_DRAWING_NS}}}blip"):
        svg_id = svg_ids.get(blip.get(embed, ""))
        if svg_id is None or blip.find(f".//{{{_SVG_NS}}}svgBlip") is not None:
            continue
        extensions = blip.find(f"{{{_DRAWING_NS}}}extLst")
        if extensions is None:
            extensions = etree.SubElement(blip, f"{{{_DRAWING_NS}}}extLst")
        extension = etree.SubElement(extensions, f"{{{_DRAWING_NS}}}ext", uri=_SVG_EXTENSION_URI)
        etree.SubElement(extension, f"{{{_SVG_NS}}}svgBlip", {embed: svg_id}, nsmap={"asvg": _SVG_NS})
        changed = True
    return changed


def _with_svg_content_type(content_types: bytes) -> bytes:
    """The content types of the package, with the default for the .svg extension."""
    root = etree.fromstring(content_types, _PARSER)
    for default in root.iter(f"{{{_CONTENT_TYPES_NS}}}Default"):
        if (default.get("Extension") or "").lower() == "svg":
            return content_types
    etree.SubElement(root, f"{{{_CONTENT_TYPES_NS}}}Default", Extension="svg", ContentType=_SVG_CONTENT_TYPE)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _free_name(name: str, names: set[str]) -> str:
    """name, or name with a number before its extension, whichever the package does not hold yet."""
    stem, extension = posixpath.splitext(name)
    candidate, number = name, 1
    while candidate in names:
        candidate = f"{stem}-{number}{extension}"
        number += 1
    return candidate


def _free_id(ids: set[str | None]) -> str:
    """A relationship id the part does not use yet."""
    number = 1
    while f"rIdSvg{number}" in ids:
        number += 1
    return f"rIdSvg{number}"
