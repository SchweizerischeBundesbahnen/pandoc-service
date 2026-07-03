"""Unit tests for ``app.DocxLatexPreProcess`` (the single-pass orchestrator).

It must produce the same package as chaining the three standalone docx→latex
preprocessors, but in one unzip/re-zip so an image-heavy document's media is
recompressed once rather than three times.
"""

from __future__ import annotations

import io
import zipfile

from app import DocxColorPreProcess, DocxLatexPreProcess, DocxListLevelPreProcess, DocxMathColorPreProcess, DocxParagraphPreProcess, DocxTablePreProcess

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
STYLES = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'


def _pack(parts: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return buf.getvalue()


def _entries(blob: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _doc(*body: str) -> bytes:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + "".join(body) + "</w:body></w:document>").encode()


# A paragraph that is aligned (jc), inside a numbered list (numPr), with a
# coloured + sized run — i.e. work for all three preprocessors at once.
_BODY = _doc(
    '<w:p><w:pPr><w:jc w:val="center"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="32"/></w:rPr><w:t>x</w:t></w:r></w:p>',
)


def test_single_pass_matches_sequential_preprocessors():
    blob = _pack({"word/document.xml": _BODY, "word/styles.xml": STYLES, "word/media/img.png": b"\x89PNG" + b"\x00" * 500})

    sequential = DocxMathColorPreProcess.preprocess(DocxTablePreProcess.preprocess(DocxListLevelPreProcess.preprocess(DocxParagraphPreProcess.preprocess(DocxColorPreProcess.preprocess(blob)))))
    single = DocxLatexPreProcess.preprocess(blob)

    seq_entries, single_entries = _entries(sequential), _entries(single)
    assert set(seq_entries) == set(single_entries)
    # Every part (document.xml, styles.xml, media) is byte-identical in content.
    for name, data in seq_entries.items():
        assert single_entries[name] == data, f"{name} differs between single-pass and sequential"


def test_media_is_preserved_and_changes_applied():
    blob = _pack({"word/document.xml": _BODY, "word/styles.xml": STYLES, "word/media/img.png": b"PNGDATA"})
    out = _entries(DocxLatexPreProcess.preprocess(blob))
    # media untouched
    assert out["word/media/img.png"] == b"PNGDATA"
    # the run now references a synthetic style (colour + size captured)
    assert b"PandocColor__FG_FF0000__SZ_32" in out["word/document.xml"]
    # the paragraph carries the alignment style and a list-level sentinel
    assert b"PandocPara__ALIGN_center" in out["word/document.xml"]
    assert "".encode() in out["word/document.xml"]


def test_math_colour_applied_without_styles_xml():
    """A colored equation with no styles.xml: colour/paragraph rewrites skip
    (has_styles False), but the math-colour rewrite still runs in the single pass."""
    math_body = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:body><w:p><m:oMath><m:r><w:rPr><w:color w:val="FF0000"/></w:rPr><m:t>E</m:t></m:r></m:oMath></w:p></w:body></w:document>'
    ).encode()
    blob = _pack({"word/document.xml": math_body})
    out = _entries(DocxLatexPreProcess.preprocess(blob))["word/document.xml"]
    # The colour is encoded as markers and the direct <w:color> stripped.
    assert b"PMCzzzFF0000zzzEzzzPMCENDzzz" in out
    assert b"<w:color" not in out


def test_unchanged_document_returns_original_bytes():
    """A document nothing rewrites (plain run, no colour/list/table/math) returns
    the original bytes rather than a re-zipped copy."""
    plain = _doc("<w:p><w:r><w:t>plain</w:t></w:r></w:p>")
    blob = _pack({"word/document.xml": plain, "word/styles.xml": STYLES})
    assert DocxLatexPreProcess.preprocess(blob) == blob


def test_non_docx_returned_unchanged():
    assert DocxLatexPreProcess.preprocess(b"not a zip") == b"not a zip"


def test_no_body_parts_returned_unchanged():
    blob = _pack({"word/styles.xml": STYLES})
    assert DocxLatexPreProcess.preprocess(blob) == blob


def test_single_pass_includes_table_cell_preprocessing():
    """Table cell backgrounds are tagged by the single-pass orchestrator."""
    body_with_table = _doc(
        '<w:tbl><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    )
    blob = _pack({"word/document.xml": body_with_table, "word/styles.xml": STYLES})

    sequential = DocxTablePreProcess.preprocess(blob)
    single = DocxLatexPreProcess.preprocess(blob)

    seq_entries, single_entries = _entries(sequential), _entries(single)
    assert set(seq_entries) == set(single_entries)
    for name, data in seq_entries.items():
        assert single_entries[name] == data, f"{name} differs between single-pass and sequential"

    # Verify the sentinel is present
    assert "\ue010bg=D9EAF7\ue011".encode() in single_entries["word/document.xml"]
