"""Record the size each HTML image asks for, before pandoc fits it to the template's page.

pandoc brings an image wider than the text of the template's page back to that width, and keeps
the size it asked for nowhere in the DOCX. The page the image ends up on can have more room: a
landscape section of a portrait document, or a larger paper size. To give an image that room, up
to the size it asked for, the post-processor needs that size.

The sizes are read the way pandoc reads them: the ``width`` and ``height`` attributes of an
``<img>`` first, else its CSS ``width`` and ``height`` (see ``filters/inline_styles.lua``), and for an
image stating neither, the size :mod:`app.html_image_pre_process` gives it. pandoc keeps the bytes
of an image as they are, so the post-processor finds the size of each picture by its bytes, taking
images with the same bytes in document order.
"""

import base64
import binascii
import hashlib
import logging
from dataclasses import dataclass

from lxml import etree  # type: ignore[import-untyped]

from app import html_image_pre_process
from app.html_document import parse_document

logger = logging.getLogger(__name__)

# Exceptions that mean the input is not parseable HTML: no sizes, and the images keep pandoc's.
_PARSE_FAILURES = (etree.ParseError, etree.ParserError, ValueError)
# A data: URI whose payload is not base64: no size to record. Named for the reason app/html_table_layout.py gives.
_DECODE_FAILURES = (binascii.Error, ValueError)


@dataclass(frozen=True)
class RequestedSize:
    """The width and height an image asks for, each a CSS length or empty where it states none."""

    width: str
    height: str


def extract(source: bytes | str) -> dict[str, list[RequestedSize]]:
    """The requested size of each ``data:`` image, by the digest of its bytes, in document order."""
    data = source if isinstance(source, bytes) else source.encode("utf-8")
    try:
        doc = parse_document(data)
    except _PARSE_FAILURES:
        logger.warning("html_image_sizes: HTML parse failed; no image sizes extracted")
        return {}
    sizes: dict[str, list[RequestedSize]] = {}
    for img in doc.iter("img"):
        image = _decode(img.get("src"))
        if image is None:
            continue
        # The size pandoc is given for an image stating none; the tree is this module's own copy
        html_image_pre_process.size_image(img)
        style = _style(img.get("style"))
        width = img.get("width") or style.get("width", "")
        height = img.get("height") or style.get("height", "")
        sizes.setdefault(digest(image), []).append(RequestedSize(width.strip(), height.strip()))
    return sizes


def digest(image: bytes) -> str:
    """The key an image's size is found by: a digest of its bytes."""
    return hashlib.sha256(image).hexdigest()


def _decode(src: str | None) -> bytes | None:
    """The bytes of a base64 ``data:`` URI, else None."""
    if not src or not src.startswith("data:"):
        return None
    header, _, payload = src.partition(",")
    if "base64" not in header:
        return None
    try:
        return base64.b64decode(payload, validate=False)
    except _DECODE_FAILURES:
        return None


def _style(style: str | None) -> dict[str, str]:
    """Split a CSS declaration list into ``{property: value}``; a later declaration wins."""
    declarations: dict[str, str] = {}
    for declaration in (style or "").split(";"):
        name, separator, value = declaration.partition(":")
        if separator:
            declarations[name.strip().lower()] = value.strip()
    return declarations
