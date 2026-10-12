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
import binascii
import hashlib
import io
import re
import zipfile
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pathlib import Path

# The variable filters/html_embedded_images.lua reads the directory of the images from.
DIRECTORY_VARIABLE = "PANDOC_SERVICE_IMAGES_DIR"

# The raster formats the DOCX writer embeds as they are, by the subtype of their media type. An SVG
# stays a data: URI: pandoc converts it, and app.svg_processor has made most of them PNG anyway.
_EXTENSIONS = {
    b"png": "png",
    b"jpeg": "jpg",
    b"jpg": "jpg",
    b"gif": "gif",
    b"bmp": "bmp",
    b"tiff": "tiff",
}

# The src of an <img> holding a base64 data: URI of one of those formats. The attributes before it
# are skipped a quoted value at a time, so a ">" or a "src=" inside a value is not taken for the end
# of the tag or for its src. No search crosses a "<", quoted or not, so each character is scanned
# from one tag start at most and malformed HTML takes linear time. A tag the expression does not
# match, a "<" in a value before the src among them, keeps its data: URI, which pandoc reads as before.
_IMG_DATA_URI_RE = re.compile(
    rb"""(<img\b(?:[^<>"']|"[^"<]*"|'[^'<]*')*?\ssrc\s*=\s*)(["'])data:image/([a-z]+);base64,([A-Za-z0-9+/=\s]*)\2""",
    re.IGNORECASE,
)

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

    Equal images share one file, as pandoc keeps one media part for equal ``data:`` URIs.
    """
    names: dict[bytes, str] = {}

    def replace(match: re.Match[bytes]) -> bytes:
        extension = _EXTENSIONS.get(match.group(3).lower())
        if extension is None:
            return match.group(0)
        try:
            image = base64.b64decode(match.group(4), validate=False)
        except binascii.Error, ValueError:
            return match.group(0)
        if not image:
            return match.group(0)
        key = hashlib.sha256(image).digest()
        name = names.get(key)
        if name is None:
            name = f"{_NAME_PREFIX}{len(names) + 1}.{extension}"
            names[key] = name
            (directory / name).write_bytes(image)
        quote = match.group(2)
        return match.group(1) + quote + name.encode() + quote

    return _IMG_DATA_URI_RE.sub(replace, source)


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
