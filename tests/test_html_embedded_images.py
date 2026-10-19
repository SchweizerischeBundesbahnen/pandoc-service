"""The data: images of an HTML document, handed to pandoc as files; #271."""

from __future__ import annotations

import base64
import io
import re
import time
import zipfile
from typing import TYPE_CHECKING

from app import html_embedded_images
from app.html_document import parse_document

if TYPE_CHECKING:
    from pathlib import Path

PNG = b"\x89PNG\r\n\x1a\n" + b"png bytes"
GIF = b"GIF89a" + b"gif bytes"


def _uri(subtype: str, image: bytes) -> str:
    return f"data:image/{subtype};base64,{base64.b64encode(image).decode()}"


# A name extract gives an image: the prefix, the token of its conversion, a number.
NAME = re.compile(r"pandoc-service-image-[0-9a-f]{32}-(\d+)\.([a-z]+)")


def _numbered(srcs: list[str | None]) -> list[str]:
    """Each src as "<number>.<extension>", the token of the conversion left out, after checking its form."""
    numbered = []
    for src in srcs:
        match = NAME.fullmatch(src or "")
        assert match, src
        numbered.append(f"{match.group(1)}.{match.group(2)}")
    return numbered


def _srcs(source: bytes) -> list[str | None]:
    return [img.get("src") for img in parse_document(source).iter("img")]


def _files(directory: Path) -> dict[str, bytes]:
    return {_numbered([path.name])[0]: path.read_bytes() for path in sorted(directory.iterdir())}


def test_an_image_is_written_into_the_directory_and_named_in_its_src(tmp_path: Path):
    source = f'<p><img alt="a" src="{_uri("png", PNG)}" style="width: 10px"/></p>'.encode()

    result = html_embedded_images.extract(source, tmp_path)

    img = next(parse_document(result).iter("img"))
    assert _numbered([img.get("src")]) == ["1.png"]
    assert (img.get("alt"), img.get("style")) == ("a", "width: 10px")
    assert _files(tmp_path) == {"1.png": PNG}
    assert (tmp_path / img.get("src")).is_file()


def test_equal_images_share_one_file(tmp_path: Path):
    source = f'<img src="{_uri("png", PNG)}"><img src=\'{_uri("png", PNG)}\'><img src="{_uri("gif", GIF)}">'.encode()

    result = html_embedded_images.extract(source, tmp_path)

    assert _numbered(_srcs(result)) == ["1.png", "1.png", "2.gif"]
    assert _files(tmp_path) == {"1.png": PNG, "2.gif": GIF}


def test_the_tag_and_the_media_type_are_read_in_any_case(tmp_path: Path):
    source = f'<IMG SRC="{_uri("JPEG", b"jpeg bytes").replace("data:image", "DATA:IMAGE").replace("base64", "BASE64")}">'.encode()

    assert _numbered(_srcs(html_embedded_images.extract(source, tmp_path))) == ["1.jpg"]
    assert _files(tmp_path) == {"1.jpg": b"jpeg bytes"}


def test_base64_broken_over_lines_is_read(tmp_path: Path):
    encoded = base64.encodebytes(PNG * 20).decode()

    html_embedded_images.extract(f'<img src="data:image/png;base64,{encoded}">'.encode(), tmp_path)

    assert _files(tmp_path) == {"1.png": PNG * 20}


def test_the_src_is_read_as_html_reads_it(tmp_path: Path):
    """Quoted values may hold ">", "<" and "src=", a src may go unquoted and carry an entity."""
    source = (f'<img title="a > b" alt=\'c < d\' data-x=\' src="{_uri("gif", GIF)}"\' src="{_uri("png", PNG)}"><img src={_uri("png", PNG)}><img src="data:image/png;base64,{base64.b64encode(GIF * 3).decode()}&#61;">').encode()

    result = html_embedded_images.extract(source, tmp_path)

    assert _numbered(_srcs(result)) == ["1.png", "1.png", "2.png"]
    assert next(parse_document(result).iter("img")).get("data-x") == f' src="{_uri("gif", GIF)}"'
    assert _files(tmp_path)["2.png"] == GIF * 3


def test_what_is_no_raster_data_image_of_an_img_is_left_alone(tmp_path: Path):
    source = (
        f'<img src="{_uri("svg+xml", b"<svg/>")}">'  # pandoc converts an SVG itself
        f'<img src="data:image/png,raw">'  # not base64
        f'<img src="data:image/png;base64,">'  # empty
        f'<img src="https://example.com/a.png">'
        f"<img>"
        f'<a href="{_uri("png", PNG)}">link</a>'  # not an image of the document
        f'<img data-src="{_uri("png", PNG)}">'  # not its src
    ).encode()

    assert html_embedded_images.extract(source, tmp_path) is source
    assert list(tmp_path.iterdir()) == []


def test_each_conversion_names_its_images_with_a_token_of_its_own(tmp_path: Path):
    source = f'<img src="pandoc-service-image-1.png"><img src="{_uri("png", PNG)}">'.encode()
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()

    first = _srcs(html_embedded_images.extract(source, tmp_path / "a"))
    second = _srcs(html_embedded_images.extract(source, tmp_path / "b"))

    assert first[0] == second[0] == "pandoc-service-image-1.png"
    assert first[1] != second[1]
    assert _numbered([first[1], second[1]]) == ["1.png", "1.png"]


def test_malformed_html_takes_linear_time(tmp_path: Path):
    source = b'<img title=">" ' * 200_000

    start = time.perf_counter()
    result = html_embedded_images.extract(source, tmp_path)

    assert result is source
    assert time.perf_counter() - start < 5


def _docx(parts: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as docx:
        for name, data in parts.items():
            docx.writestr(name, data)
    return out.getvalue()


def _parts(docx: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_a_placeholder_gets_the_data_uri_of_its_image_back(tmp_path: Path):
    (tmp_path / "pandoc-service-image-0123abcd-1.png").write_bytes(PNG)
    (tmp_path / "pandoc-service-image-0123abcd-2.jpg").write_bytes(b"jpeg bytes")
    document = b"<w:t>{{IMG:10px||pandoc-service-image-0123abcd-1.png}} and {{IMG:|5px|pandoc-service-image-0123abcd-2.jpg}}</w:t>"
    footer = b"<w:t>{{IMG:||pandoc-service-image-0123abcd-1.png}}</w:t>"
    media = b"\x89PNG media"

    result = _parts(html_embedded_images.finish_docx(_docx({"word/document.xml": document, "word/footer1.xml": footer, "word/media/a.png": media}), tmp_path))

    assert result["word/document.xml"] == f"<w:t>{{{{IMG:10px||{_uri('png', PNG)}}}}} and {{{{IMG:|5px|{_uri('jpeg', b'jpeg bytes')}}}}}</w:t>".encode()
    assert result["word/footer1.xml"] == f"<w:t>{{{{IMG:||{_uri('png', PNG)}}}}}</w:t>".encode()
    assert result["word/media/a.png"] == media


def test_a_placeholder_naming_no_extracted_image_stays_as_it_is(tmp_path: Path):
    (tmp_path / "pandoc-service-image-0123abcd-1.png").write_bytes(PNG)
    document = b"<w:t>{{IMG:||pandoc-service-image-0123abcd-7.png}} {{IMG:||../secret.png}}</w:t>"

    result = _parts(html_embedded_images.finish_docx(_docx({"word/document.xml": document}), tmp_path))

    assert result["word/document.xml"] == document


def test_a_picture_of_an_extracted_image_gets_an_empty_description(tmp_path: Path):
    """pandoc writes the src of a picture, the name of the image, into descr of pic:cNvPr; #289."""
    (tmp_path / "pandoc-service-image-0123abcd-1.png").write_bytes(PNG)
    document = (
        b'<wp:docPr descr="Alt" title="" id="10" name="Picture"/><pic:cNvPr descr="pandoc-service-image-0123abcd-1.png" id="11" name="Picture"/>'
        b'<pic:cNvPr descr="https://example.com/a.png" id="13" name="Picture"/>'
        b'<w:t>descr="pandoc-service-image-1.png" descr="pandoc-service-image-0123abcd-1.png"</w:t>'
    )
    header = b'<pic:cNvPr descr="pandoc-service-image-0123abcd-1.png" id="1" name="Picture"/>'

    result = _parts(html_embedded_images.finish_docx(_docx({"word/document.xml": document, "word/header1.xml": header}), tmp_path))

    assert result["word/document.xml"] == (
        b'<wp:docPr descr="Alt" title="" id="10" name="Picture"/><pic:cNvPr descr="" id="11" name="Picture"/>'
        b'<pic:cNvPr descr="https://example.com/a.png" id="13" name="Picture"/>'
        b'<w:t>descr="pandoc-service-image-1.png" descr="pandoc-service-image-0123abcd-1.png"</w:t>'
    )
    assert result["word/header1.xml"] == b'<pic:cNvPr descr="" id="1" name="Picture"/>'


def test_a_document_without_placeholders_is_returned_unchanged(tmp_path: Path):
    (tmp_path / "pandoc-service-image-0123abcd-1.png").write_bytes(PNG)
    docx = _docx({"word/document.xml": b"<w:t>text</w:t>"})

    assert html_embedded_images.finish_docx(docx, tmp_path) is docx


def test_the_output_of_a_conversion_which_extracted_nothing_is_not_read(tmp_path: Path):
    assert html_embedded_images.finish_docx(b"%PDF-1.7 not a zip", tmp_path) == b"%PDF-1.7 not a zip"


def test_text_is_read_as_utf_8_as_pandoc_reads_it(tmp_path: Path):
    """lxml guesses Latin-1 for HTML declaring no charset; pandoc reads UTF-8 whatever a document declares."""
    image = f'<img src="{_uri("png", PNG)}">'
    for source in (
        f"<p>café Grüße 日本</p>{image}".encode(),
        f'<html><head><meta charset="iso-8859-1"></head><body><p>café Grüße 日本</p>{image}</body></html>'.encode(),
    ):
        result = html_embedded_images.extract(source, tmp_path)

        assert "café Grüße 日本" in result.decode("utf-8")
