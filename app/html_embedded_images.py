"""Hand the images an HTML document carries as ``data:`` URIs to pandoc as files.

pandoc keeps a ``data:`` URI as a string in its document tree and passes it through every Lua
filter. A document of 20 MB with some 4000 small images took pandoc 5.8 GiB of memory that way, and
218 MiB with the same images as files (SchweizerischeBundesbahnen/pandoc-service#271).

pandoc runs with ``--sandbox``, which reads no file a document names, so the images cannot simply
be pointed at. :func:`extract` writes each image into a directory of the conversion and leaves its
name in the ``src``. ``filters/html_embedded_images.lua`` reads those files into the media bag of
pandoc, which the DOCX writer takes them from, and reads nothing else. The directory reaches the
filter through an environment variable, which a document cannot set.

``filters/inline_styles.lua`` writes some images into the text as ``{{IMG:...}}`` placeholders for
:mod:`app.docx_post_process`. :func:`restore_placeholders` puts their ``data:`` URIs back from the
files, before the directory is removed.
"""

from __future__ import annotations

import base64
import hashlib
import io
import logging
import re
import zipfile
from typing import TYPE_CHECKING

from lxml import etree, html  # type: ignore[import-untyped]

from app.html_document import parse_document

if TYPE_CHECKING:
    from pathlib import Path

logger = logging.getLogger(__name__)

# The variable filters/html_embedded_images.lua reads the directory of the images from.
DIRECTORY_VARIABLE = "PANDOC_SERVICE_IMAGES_DIR"

# The raster formats the DOCX writer embeds as they are, by the subtype of their media type. An SVG
# stays a data: URI: pandoc converts it, and app.svg_processor has made most of them PNG anyway.
_EXTENSIONS = {
    "png": "png",
    "jpeg": "jpg",
    "jpg": "jpg",
    "gif": "gif",
    "bmp": "bmp",
    "tiff": "tiff",
}

_DATA_IMAGE_PREFIX = "data:image/"

# See html_lists_pre_process for why these are named tuples rather than inline except-literals.
_PARSE_FAILURES = (etree.ParseError, etree.ParserError, ValueError)

# The media subtype of each extension, for a data: URI rebuilt from a file.
_SUBTYPES = {"png": "png", "jpg": "jpeg", "gif": "gif", "bmp": "bmp", "tiff": "tiff"}

# The name an extracted image is given. filters/html_embedded_images.lua accepts this form only.
_NAME_PREFIX = "pandoc-service-image-"

# {{IMG:<width>|<height>|<src>}}: see image_placeholder in filters/inline_styles.lua.
_PLACEHOLDER_RE = re.compile(rf"(\{{\{{IMG:[^|]*\|[^|]*\|)({_NAME_PREFIX}\d+\.([a-z]+))(\}}\}})".encode())

# The parts of a DOCX whose text can hold a placeholder.
_TEXT_PART_RE = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml")


def extract(source: bytes, directory: Path) -> bytes:
    """Write the ``data:`` images of `source` into `directory` and name them in their ``src``.

    Equal images share one file, as pandoc keeps one media part for equal ``data:`` URIs. Returns
    the input unchanged when it does not parse or holds no image to extract.
    """
    try:
        doc = parse_document(source)
    except _PARSE_FAILURES:
        logger.warning("html_embedded_images: HTML parse failed; passing input through")
        return source

    names: dict[bytes, str] = {}
    for img in doc.iter("img"):
        name = _extract_image(img.get("src"), directory, names)
        if name is not None:
            img.set("src", name)
    if not names:
        return source
    return html.tostring(doc, encoding="utf-8")


def _extract_image(src: str | None, directory: Path, names: dict[bytes, str]) -> str | None:
    """The file name of the image of a ``data:`` URI, written into `directory` once; None for any other src."""
    header, _, payload = (src or "").partition(",")
    media_type, _, encoding = header.lower().rpartition(";")
    if not media_type.startswith(_DATA_IMAGE_PREFIX) or encoding != "base64":
        return None
    extension = _EXTENSIONS.get(media_type.removeprefix(_DATA_IMAGE_PREFIX).split(";")[0])
    if extension is None:
        return None
    try:
        image = base64.b64decode(payload, validate=False)
    # binascii.Error, what an undecodable payload raises, is a ValueError
    except ValueError:
        return None
    if not image:
        return None
    key = hashlib.sha256(image).digest()
    if key not in names:
        names[key] = f"{_NAME_PREFIX}{len(names) + 1}.{extension}"
        (directory / names[key]).write_bytes(image)
    return names[key]


def restore_placeholders(docx: bytes, directory: Path) -> bytes:
    """Put the ``data:`` URI of an extracted image back into each ``{{IMG:...}}`` placeholder of `docx`.

    A conversion which extracted no image returns its output as it is, whatever its format.
    """
    if not any(directory.iterdir()):
        return docx
    with zipfile.ZipFile(io.BytesIO(docx)) as source:
        parts = {info.filename: source.read(info) for info in source.infolist() if _TEXT_PART_RE.fullmatch(info.filename)}
        changed = {name: _restore(text, directory) for name, text in parts.items() if _PLACEHOLDER_RE.search(text)}
        if not changed:
            return docx
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                target.writestr(info, changed.get(info.filename, source.read(info)))
    return out.getvalue()


def _restore(text: bytes, directory: Path) -> bytes:
    def replace(match: re.Match[bytes]) -> bytes:
        path = directory / match.group(2).decode()
        subtype = _SUBTYPES.get(match.group(3).decode())
        if subtype is None or not path.is_file():
            return match.group(0)
        uri = f"data:image/{subtype};base64,".encode() + base64.b64encode(path.read_bytes())
        return match.group(1) + uri + match.group(4)

    return _PLACEHOLDER_RE.sub(replace, text)
