"""Unit tests for ``app.docx_latex_pre_process`` (the single-pass orchestrator).

It must produce the same package as chaining the three standalone docx→latex
preprocessors, but in one unzip/re-zip so an image-heavy document's media is
recompressed once rather than three times.
"""

from __future__ import annotations

import io
import zipfile

from app import docx_color_pre_process, docx_latex_pre_process, docx_list_level_pre_process, docx_math_color_pre_process, docx_paragraph_pre_process, docx_table_pre_process

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

    sequential = docx_math_color_pre_process.preprocess(docx_table_pre_process.preprocess(docx_list_level_pre_process.preprocess(docx_paragraph_pre_process.preprocess(docx_color_pre_process.preprocess(blob)))))
    single = docx_latex_pre_process.preprocess(blob)

    seq_entries, single_entries = _entries(sequential), _entries(single)
    assert set(seq_entries) == set(single_entries)
    # Every part (document.xml, styles.xml, media) is byte-identical in content.
    for name, data in seq_entries.items():
        assert single_entries[name] == data, f"{name} differs between single-pass and sequential"


def test_media_is_preserved_and_changes_applied():
    blob = _pack({"word/document.xml": _BODY, "word/styles.xml": STYLES, "word/media/img.png": b"PNGDATA"})
    out = _entries(docx_latex_pre_process.preprocess(blob))
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
    out = _entries(docx_latex_pre_process.preprocess(blob))["word/document.xml"]
    # The colour is encoded as markers and the direct <w:color> stripped.
    assert b"PMCzzzFF0000zzzEzzzPMCENDzzz" in out
    assert b"<w:color" not in out


def test_unchanged_document_returns_original_bytes():
    """A document nothing rewrites (plain run, no colour/list/table/math) returns
    the original bytes rather than a re-zipped copy."""
    plain = _doc("<w:p><w:r><w:t>plain</w:t></w:r></w:p>")
    blob = _pack({"word/document.xml": plain, "word/styles.xml": STYLES})
    assert docx_latex_pre_process.preprocess(blob) == blob


def test_non_docx_returned_unchanged():
    assert docx_latex_pre_process.preprocess(b"not a zip") == b"not a zip"


def test_no_body_parts_returned_unchanged():
    blob = _pack({"word/styles.xml": STYLES})
    assert docx_latex_pre_process.preprocess(blob) == blob


def test_single_pass_includes_table_cell_preprocessing():
    """Table cell backgrounds are tagged by the single-pass orchestrator."""
    body_with_table = _doc(
        '<w:tbl><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    )
    blob = _pack({"word/document.xml": body_with_table, "word/styles.xml": STYLES})

    sequential = docx_table_pre_process.preprocess(blob)
    single = docx_latex_pre_process.preprocess(blob)

    seq_entries, single_entries = _entries(sequential), _entries(single)
    assert set(seq_entries) == set(single_entries)
    for name, data in seq_entries.items():
        assert single_entries[name] == data, f"{name} differs between single-pass and sequential"

    # Verify the sentinel is present
    assert "\ue010bg=D9EAF7\ue011".encode() in single_entries["word/document.xml"]


def test_a_caption_gets_its_style_defined_for_pandoc():
    """pandoc keeps the name of a paragraph style only when styles.xml defines it."""
    import io as _io
    import zipfile as _zipfile

    from app import docx_latex_pre_process

    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    buffer = _io.BytesIO()
    with _zipfile.ZipFile(buffer, "w") as package:
        package.writestr("word/document.xml", f'<w:document {w}><w:body><w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Table 1 x</w:t></w:r></w:p></w:body></w:document>')
        package.writestr("word/styles.xml", f'<w:styles {w}><w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/></w:style></w:styles>')
    result = _zipfile.ZipFile(_io.BytesIO(docx_latex_pre_process.preprocess(buffer.getvalue())))
    assert 'w:val="PandocCaption"' in result.read("word/document.xml").decode()
    assert 'w:styleId="PandocCaption"' in result.read("word/styles.xml").decode()


def test_preprocess_takes_the_color_of_the_hyperlink_style():
    """The single-pass orchestrator hands the character style colors to the color rewrite too."""
    styles = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="156082" w:themeColor="accent1"/></w:rPr></w:style>'
        b"</w:styles>"
    )
    document = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>'
        b'<w:hyperlink w:anchor="_a"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink>'
        b"</w:p></w:body></w:document>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        zf.writestr("word/document.xml", document)
        zf.writestr("word/styles.xml", styles)

    result = docx_latex_pre_process.preprocess(buffer.getvalue())

    with zipfile.ZipFile(io.BytesIO(result)) as zf:
        assert b'w:val="PandocColor__FG_156082"' in zf.read("word/document.xml")
        assert b'w:styleId="PandocColor__FG_156082"' in zf.read("word/styles.xml")


def test_preprocess_keeps_translated_table_of_contents_entries_uncolored():
    """The single-pass orchestrator hands the table of contents styles to the color rewrite too."""
    styles = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="156082"/></w:rPr></w:style>'
        b'<w:style w:type="paragraph" w:styleId="Verzeichnis1"><w:name w:val="toc 1"/></w:style>'
        b"</w:styles>"
    )
    document = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        b'<w:p><w:pPr><w:pStyle w:val="Verzeichnis1"/></w:pPr><w:hyperlink w:anchor="_a"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>entry</w:t></w:r></w:hyperlink></w:p>'
        b"</w:body></w:document>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        zf.writestr("word/document.xml", document)
        zf.writestr("word/styles.xml", styles)

    result = docx_latex_pre_process.preprocess(buffer.getvalue())

    with zipfile.ZipFile(io.BytesIO(result)) as zf:
        assert b"PandocColor__FG_156082" not in zf.read("word/document.xml")


def test_the_styles_which_break_the_page_before_their_paragraphs():
    from app.docx_ooxml import parse_xml
    from app.docx_page_break_pre_process import page_break_styles

    styles = parse_xml(
        b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:style w:styleId="Chapter"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>'
        b'<w:style w:styleId="Part"><w:basedOn w:val="Chapter"/></w:style>'
        b'<w:style w:styleId="Plain"><w:basedOn w:val="Chapter"/><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr></w:style>'
        b'<w:style w:styleId="Other"/></w:styles>'
    )
    assert page_break_styles(styles) == frozenset({"Chapter", "Part"})


def test_a_long_chain_of_styles_based_on_each_other():
    from app.docx_ooxml import parse_xml
    from app.docx_page_break_pre_process import page_break_styles

    chain = "".join(f'<w:style w:styleId="S{index}"><w:basedOn w:val="S{index - 1}"/></w:style>' for index in range(1, 5000))
    styles = parse_xml(b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:styleId="S0"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>' + chain.encode() + b"</w:styles>")
    assert len(page_break_styles(styles)) == 5000
