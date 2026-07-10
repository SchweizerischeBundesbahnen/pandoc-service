import logging
import re
from typing import TYPE_CHECKING, Any

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject

# XML namespace identifier, not an actual HTTP URL
SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR

# Common XPath expressions
XPATH_P_PR = ".//w:pPr"
XPATH_P_STYLE = ".//w:pStyle"

# Paragraph style id that marks a genuine caption. filters/html_captions.lua
# sets it on paragraphs that carry Polarion's <span data-sequence=...> caption
# counter. Non-HTML inputs bring their own caption styles: pandoc emits
# "TableCaption"/"ImageCaption" for markdown/docx table & figure captions, and
# Word uses "Caption". Keying off this set of caption STYLES — rather than "does
# the text start with Table/Figure" — captures real captions from every source
# while still excluding headings ("Table test III"), cross-references
# ("Table 1 shows ...") and labels ("Table 50px"), which never carry a caption
# style.
CAPTION_STYLE_IDS = frozenset({"Caption", "TableCaption", "ImageCaption"})

logger = logging.getLogger(__name__)


def add_table_of_contents_entries(doc: DocumentObject) -> None:
    """
    Add Table of Contents, Table of Figures and/or Table of Tables by:
    1. Finding figure/table captions, ensuring SEQ fields, adding TC fields with bookmarks
    2. Replacing TOC_PLACEHOLDER with Word TOC field
    3. Replacing TOF_PLACEHOLDER with Word TOF field (with cached entries)
    4. Replacing TOT_PLACEHOLDER with Word TOT field (with cached entries)
    """
    body = doc.element.body

    # Step 1: Find and process captions
    figure_paragraphs, table_paragraphs = _find_and_process_captions(body)

    # Step 2: Find elements to replace
    elements_to_replace = _find_elements_to_replace(body)

    # Step 3: Replace elements with Word fields
    _replace_elements_with_fields(body, elements_to_replace, figure_paragraphs, table_paragraphs)


def enable_auto_update_fields(doc: DocumentObject) -> None:
    """
    Enable automatic update of fields when document is opened.
    This sets the updateFields flag in settings.xml
    """
    try:
        settings = doc.settings
        settings_element = settings.element

        update_fields = settings_element.find(".//w:updateFields", namespaces={"w": SCHEMA})

        if update_fields is None:
            update_fields = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
            settings_element.append(update_fields)
            logger.info("Enabled auto-update of fields on document open")
        else:
            update_fields.set(f"{{{SCHEMA}}}val", "true")
            logger.info("Updated existing auto-update fields setting")
    except Exception as e:
        logger.warning(f"Could not enable auto-update fields: {e}")


def _escape_xml(text: str) -> str:
    """Escape characters that aren't safe inside XML element body or attributes."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _get_max_bookmark_id(body: Any) -> int:
    """Find the highest bookmark ID in the document to avoid ID conflicts."""
    max_id = 0
    for bm in body.findall(".//w:bookmarkStart", namespaces={"w": SCHEMA}):
        try:
            bm_id = int(bm.get(f"{{{SCHEMA}}}id", "0"))
            max_id = max(max_id, bm_id)
        except ValueError:
            pass
    return max_id


def _find_and_process_captions(body: Any) -> tuple[list, list]:
    """Find all figure and table captions, ensure SEQ fields, add TC fields with bookmarks.

    Returns two lists (figure_entries, table_entries) where each entry is
    a tuple ``(paragraph_element, caption_text, bookmark_name)``.
    """
    figure_paragraphs: list[tuple[Any, str, str]] = []
    table_paragraphs: list[tuple[Any, str, str]] = []

    bookmark_id = _get_max_bookmark_id(body)

    for para in body.findall(".//w:p", namespaces={"w": SCHEMA}):
        style = _get_paragraph_style(para)
        if style not in CAPTION_STYLE_IDS:
            continue

        text_stripped = _get_paragraph_text(para).strip()

        # Classify Figure vs Table. Language-independent strategy:
        # 1. Pandoc's own styles are unambiguous (ImageCaption / TableCaption).
        # 2. For the generic "Caption" style, check whether the paragraph is
        #    adjacent to a <w:tbl> element — if so it's a table caption.
        is_figure = style == "ImageCaption" or (style != "TableCaption" and not _is_adjacent_to_table(para))
        seq_name = "Figure" if is_figure else "Table"

        # Ensure the caption number is a SEQ field (not plain text).
        _ensure_seq_field(para, seq_name)
        # Re-read text after potential SEQ insertion (content may have changed)
        text_stripped = _get_paragraph_text(para).strip()

        # Ensure Caption style is set
        _ensure_caption_style(para)

        # Assign a unique bookmark for this caption's TC field
        bookmark_id += 1
        bookmark_name = f"_Toc{bookmark_id:09d}"

        # Add TC field with bookmark at the end of the caption paragraph
        field_flag = "F" if is_figure else "T"
        _add_tc_field(para, text_stripped, field_flag, bookmark_id, bookmark_name)

        if is_figure:
            figure_paragraphs.append((para, text_stripped, bookmark_name))
            logger.debug(f"Processed figure caption: {text_stripped}")
        else:
            table_paragraphs.append((para, text_stripped, bookmark_name))
            logger.debug(f"Processed table caption: {text_stripped}")

    logger.info(f"Found {len(figure_paragraphs)} figure captions and {len(table_paragraphs)} table captions")
    return figure_paragraphs, table_paragraphs


def _is_adjacent_to_table(para: Any) -> bool:
    """Return True if the paragraph is followed by a w:tbl element.

    This is a language-independent way to classify captions: table captions
    sit **before** their table in the document body. Only the forward
    direction is checked — a caption after a table (e.g. a figure caption
    that happens to follow a table) must not be misclassified as a table
    caption.

    Skips non-content elements (bookmarkStart, bookmarkEnd, sectPr, empty
    paragraphs) that pandoc/Word may insert between a caption and its table.
    """
    tbl_tag = f"{{{SCHEMA}}}tbl"
    p_tag = f"{{{SCHEMA}}}p"
    skip_tags = {
        f"{{{SCHEMA}}}bookmarkStart",
        f"{{{SCHEMA}}}bookmarkEnd",
        f"{{{SCHEMA}}}sectPr",
    }

    def _is_empty_para(el: Any) -> bool:
        return el.tag == p_tag and not _get_paragraph_text(el).strip()

    def _should_skip(el: Any) -> bool:
        return el.tag in skip_tags or _is_empty_para(el)

    # Look forward only — table caption always precedes its table
    el = para.getnext()
    while el is not None and _should_skip(el):
        el = el.getnext()
    return el is not None and el.tag == tbl_tag


def _has_seq_field(para: Any) -> bool:
    """Return True if the paragraph already contains a SEQ field."""
    return any(instr.text and "SEQ" in instr.text for instr in para.findall(".//w:instrText", namespaces={"w": SCHEMA}))


def _ensure_seq_field(para: Any, seq_name: str) -> None:
    """Replace a plain-text caption number with a Word SEQ field.

    For a caption paragraph like ``Figure 1 My picture`` this replaces the
    ``1`` with ``{SEQ Figure \\* ARABIC}`` so Word can auto-number captions.
    Does nothing if the paragraph already contains a SEQ field (e.g. one
    inserted by the Lua filter).
    """
    if _has_seq_field(para):
        return

    text = _get_paragraph_text(para).strip()
    match = re.match(r"^(\D+?)(\d+)(.*)", text, re.DOTALL)
    if not match:
        return

    prefix = match.group(1)
    number = match.group(2)
    suffix = match.group(3)

    # Remove all existing runs (keep pPr and non-run children like bookmarks)
    for run in list(para.findall("w:r", namespaces={"w": SCHEMA})):
        para.remove(run)

    ns = f'xmlns:w="{SCHEMA}"'

    para.append(parse_xml(f'<w:r {ns}><w:t xml:space="preserve">{_escape_xml(prefix)}</w:t></w:r>'))
    para.append(parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="begin"/></w:r>'))
    para.append(parse_xml(f'<w:r {ns}><w:instrText xml:space="preserve"> SEQ {_escape_xml(seq_name)} \\* ARABIC </w:instrText></w:r>'))
    para.append(parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="separate"/></w:r>'))
    para.append(parse_xml(f"<w:r {ns}><w:t>{_escape_xml(number)}</w:t></w:r>"))
    para.append(parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="end"/></w:r>'))

    if suffix:
        para.append(parse_xml(f'<w:r {ns}><w:t xml:space="preserve">{_escape_xml(suffix)}</w:t></w:r>'))

    logger.debug(f"Replaced plain-text number with SEQ {seq_name} field in caption: {text}")


def _ensure_caption_style(para: Any) -> None:
    """Ensure the paragraph has the Caption style set."""
    p_pr = para.find(XPATH_P_PR, namespaces={"w": SCHEMA})
    if p_pr is None:
        p_pr = parse_xml(f'<w:pPr {nsdecls("w")}><w:pStyle w:val="Caption"/></w:pPr>')
        para.insert(0, p_pr)
    else:
        p_style = p_pr.find(XPATH_P_STYLE, namespaces={"w": SCHEMA})
        if p_style is None:
            p_style = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="Caption"/>')
            p_pr.insert(0, p_style)
        else:
            p_style.set(f"{{{SCHEMA}}}val", "Caption")


def _add_tc_field(para: Any, caption_text: str, field_flag: str, bookmark_id: int, bookmark_name: str) -> None:
    """Add a TC field with embedded bookmark at the end of a caption paragraph.

    The TC field marks this paragraph for collection by {TOC \\f T/F}.
    A bookmark inside the TC instruction text lets cached TOF/TOT
    entries hyperlink directly to this caption.
    """
    ns = f'xmlns:w="{SCHEMA}"'
    escaped = _escape_xml(caption_text)

    para.append(parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="begin"/></w:r>'))
    para.append(parse_xml(f'<w:r {ns}><w:instrText xml:space="preserve"> TC "</w:instrText></w:r>'))
    para.append(parse_xml(f'<w:bookmarkStart {ns} w:id="{bookmark_id}" w:name="{bookmark_name}"/>'))
    para.append(parse_xml(f"<w:r {ns}><w:instrText>{escaped}</w:instrText></w:r>"))
    para.append(parse_xml(f'<w:bookmarkEnd {ns} w:id="{bookmark_id}"/>'))
    para.append(parse_xml(f'<w:r {ns}><w:instrText xml:space="preserve">" \\f {field_flag} \\l "1" </w:instrText></w:r>'))
    para.append(parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="end"/></w:r>'))


def _find_elements_to_replace(body: Any) -> list:
    """Find all elements that need to be replaced with Word fields."""
    elements_to_replace: list = []
    _find_placeholder_paragraphs(body, elements_to_replace)
    return elements_to_replace


def _find_placeholder_paragraphs(body: Any, elements_to_replace: list) -> None:
    """Find TOC_PLACEHOLDER, TOF_PLACEHOLDER, and TOT_PLACEHOLDER paragraphs."""
    for idx, element in enumerate(body):
        if element.tag.endswith("}p"):
            text = _get_paragraph_text(element).strip()
            style = _get_paragraph_style(element)

            if style not in ["BodyText", "FirstParagraph", None]:
                continue

            if text == "TOC_PLACEHOLDER":
                elements_to_replace.append((idx, element, True, False, False))
                logger.info(f"Found TOC_PLACEHOLDER at index {idx}, will replace with TOC field")
            elif text == "TOF_PLACEHOLDER":
                elements_to_replace.append((idx, element, False, True, False))
                logger.info(f"Found TOF_PLACEHOLDER at index {idx}, will replace with TOF field")
            elif text == "TOT_PLACEHOLDER":
                elements_to_replace.append((idx, element, False, False, True))
                logger.info(f"Found TOT_PLACEHOLDER at index {idx}, will replace with TOT field")


def _replace_elements_with_fields(body: Any, elements_to_replace: list, figure_paragraphs: list, table_paragraphs: list) -> None:
    """Replace found elements with Word field codes."""
    for idx, element, has_toc, has_figure_links, has_table_links in sorted(elements_to_replace, key=lambda x: x[0], reverse=True):
        if element is not None:
            body.remove(element)
            logger.debug(f"Removed element at index {idx}")

        _insert_field_at_position(body, idx, has_toc, has_figure_links, has_table_links, figure_paragraphs, table_paragraphs)


def _insert_field_at_position(body: Any, idx: int, has_toc: bool, has_figure_links: bool, has_table_links: bool, figure_paragraphs: list, table_paragraphs: list) -> None:  # noqa: PLR0913
    """Insert appropriate Word fields at the specified position."""
    if has_toc:
        toc_paragraphs = _create_toc_field()
        for toc_para in reversed(toc_paragraphs):
            body.insert(idx, toc_para)
        logger.info(f"Inserted Table of Contents at index {idx}")

    if has_figure_links and figure_paragraphs:
        figure_entries = [(text, bm) for _, text, bm in figure_paragraphs]
        tof_paragraphs = _create_tof_field(figure_entries)
        for tof_para in reversed(tof_paragraphs):
            body.insert(idx, tof_para)
        logger.info(f"Inserted Table of Figures at index {idx}")

    if has_table_links and table_paragraphs:
        table_entries = [(text, bm) for _, text, bm in table_paragraphs]
        tot_paragraphs = _create_tot_field(table_entries)
        for tot_para in reversed(tot_paragraphs):
            body.insert(idx, tot_para)
        logger.info(f"Inserted Table of Tables at index {idx}")


def _create_field(field_code: str) -> list[Any]:
    """Create a Word field with specified field code (no cached entries)."""
    field_para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> {field_code} </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
    ''')
    empty_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    return [field_para, empty_para]


def _create_field_with_entries(field_code: str, entries: list[tuple[str, str]]) -> list[Any]:
    """Create a Word field with pre-filled TOC entries (cached content)."""
    if not entries:
        return _create_field(field_code)

    paragraphs: list[Any] = []

    first_text, first_bm = entries[0]
    first_para = parse_xml(
        f'<w:p xmlns:w="{SCHEMA}">'
        f'<w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
        f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {field_code} </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + _hyperlink_xml(first_text, first_bm) + "</w:p>"
    )
    paragraphs.append(first_para)

    for text, bm in entries[1:]:
        para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' + _hyperlink_xml(text, bm) + "</w:p>")
        paragraphs.append(para)

    end_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    paragraphs.append(end_para)

    empty_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    paragraphs.append(empty_para)

    return paragraphs


def _hyperlink_xml(caption_text: str, bookmark_name: str) -> str:
    """Build the raw OOXML for a TOC entry hyperlink with a PAGEREF field."""
    escaped = _escape_xml(caption_text)
    return (
        f'<w:hyperlink w:anchor="{bookmark_name}" w:history="1">'
        f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
        f"<w:t>{escaped}</w:t></w:r>"
        f"<w:r><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r>"
        f'<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        f"<w:r><w:rPr><w:webHidden/></w:rPr>"
        f'<w:instrText xml:space="preserve"> PAGEREF {bookmark_name} \\h </w:instrText></w:r>'
        f'<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f"<w:r><w:rPr><w:webHidden/></w:rPr><w:t>1</w:t></w:r>"
        f'<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        f"</w:hyperlink>"
    )


def _create_toc_field() -> list[Any]:
    """Create Table of Contents field paragraphs."""
    return _create_field('TOC \\o "1-9" \\h \\z \\u')


def _create_tof_field(entries: list[tuple[str, str]] | None = None) -> list[Any]:
    """Create Table of Figures field paragraphs using TOC \\f F."""
    field_code = "TOC \\h \\z \\f F"
    if entries:
        return _create_field_with_entries(field_code, entries)
    return _create_field(field_code)


def _create_tot_field(entries: list[tuple[str, str]] | None = None) -> list[Any]:
    """Create Table of Tables field paragraphs using TOC \\f T."""
    field_code = "TOC \\h \\z \\f T"
    if entries:
        return _create_field_with_entries(field_code, entries)
    return _create_field(field_code)


def _get_paragraph_text(para: Any) -> str:
    """Extract text content from a paragraph element."""
    texts = []
    for t in para.findall(".//w:t", namespaces={"w": SCHEMA}):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _get_paragraph_style(element: Any) -> str | None:
    """Extract the style name from a paragraph element."""
    p_pr = element.find(XPATH_P_PR, namespaces={"w": SCHEMA})
    if p_pr is None:
        return None

    p_style = p_pr.find(XPATH_P_STYLE, namespaces={"w": SCHEMA})
    if p_style is None:
        return None

    return p_style.get(f"{{{SCHEMA}}}val")
