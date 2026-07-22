import logging
import re
from copy import deepcopy
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

# Word paragraph style applied to (and marking) caption paragraphs.
CAPTION_STYLE = "Caption"

# Caption styles pandoc's own writers emit for markdown/docx table & figure captions.
TABLE_CAPTION_STYLE = "TableCaption"
IMAGE_CAPTION_STYLE = "ImageCaption"

# Paragraph style ids that mark a genuine caption. filters/html_captions.lua
# sets CAPTION_STYLE on paragraphs that carry Polarion's <span data-sequence=...>
# caption counter. Non-HTML inputs bring their own caption styles: pandoc emits
# TableCaption/ImageCaption for markdown/docx table & figure captions, and
# Word uses "Caption". Keying off this set of caption STYLES — rather than "does
# the text start with Table/Figure" — captures real captions from every source
# while still excluding headings ("Table test III"), cross-references
# ("Table 1 shows ...") and labels ("Table 50px"), which never carry a caption
# style.
CAPTION_STYLE_IDS = frozenset({CAPTION_STYLE, TABLE_CAPTION_STYLE, IMAGE_CAPTION_STYLE})

# Placeholder paragraph texts the HTML producer (docx-exporter) emits for the
# ToC (table of contents), ToF (table of figures) and ToT (table of tables)
# macros.
TOC_PLACEHOLDER = "TOC_PLACEHOLDER"
TOF_PLACEHOLDER = "TOF_PLACEHOLDER"
TOT_PLACEHOLDER = "TOT_PLACEHOLDER"

# Placeholder kinds (parsed from the placeholder texts above)
KIND_TOC = "toc"
KIND_TOF = "tof"
KIND_TOT = "tot"

# Paragraph styles a placeholder paragraph may carry; placeholders inside
# titles/headings are ignored
PLACEHOLDER_PARAGRAPH_STYLES = ("BodyText", "FirstParagraph", None)

# Caption sequence identifiers assigned when a caption does not already carry
# one (standard Polarion sequences)
FIGURE_SEQUENCE = "Figure"
TABLE_SEQUENCE = "Table"

# Word field instruction codes for all generated fields
TOC_FIELD_CODE = 'TOC \\o "1-9" \\h \\z \\u'
TOF_FIELD_CODE = "TOC \\h \\z \\f F"
TOT_FIELD_CODE = "TOC \\h \\z \\f T"

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
            logger.debug("Skipping bookmark with non-integer id: %r", bm.get(f"{{{SCHEMA}}}id"))
    return max_id


def _find_and_process_captions(body: Any) -> tuple[list, list]:  # noqa: C901  # NOSONAR
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

        # Classify Figure vs Table. Strategy (in priority order):
        # 1. Pandoc's own styles are unambiguous (ImageCaption / TableCaption).
        # 2. If a SEQ field already exists (from Lua filter), its identifier
        #    comes from Polarion's data-sequence attribute. Table captions
        #    always have a <w:tbl> nearby; if adjacency fails but a SEQ name
        #    exists that is NOT "Figure", treat it as a table caption anyway
        #    (Polarion would not set a table sequence on a figure).
        # 3. Fall back to structural adjacency: if the next content element
        #    is a <w:tbl>, it's a table caption; otherwise figure.
        if style == IMAGE_CAPTION_STYLE:
            is_figure = True
        elif style == TABLE_CAPTION_STYLE or _is_adjacent_to_table(para):
            is_figure = False
        else:
            # No table nearby — check if an existing SEQ name hints at table
            existing_seq = _get_seq_name(para)
            is_figure = existing_seq is None or existing_seq == FIGURE_SEQUENCE
        seq_name = FIGURE_SEQUENCE if is_figure else TABLE_SEQUENCE

        # Ensure the caption number is a SEQ field (not plain text).
        _ensure_seq_field(para, seq_name)
        # Re-read text after potential SEQ insertion (content may have changed)
        text_stripped = _get_paragraph_text(para).strip()

        # Ensure Caption style is set
        _set_paragraph_style(para, CAPTION_STYLE)

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


def _get_seq_name(para: Any) -> str | None:
    """Extract the SEQ field identifier from a paragraph (e.g. "Table", "Tabela")."""
    for instr in para.findall(".//w:instrText", namespaces={"w": SCHEMA}):
        if instr.text and "SEQ" in instr.text:
            match = re.search(r"SEQ\s+(\S+)", instr.text)
            if match:
                return match.group(1)
    return None


def _clone_run_with_text(run: Any, text: str) -> Any:
    """Create a new run with the given text, copying rPr from the source run."""
    ns = f'xmlns:w="{SCHEMA}"'
    new_run = parse_xml(f'<w:r {ns}><w:t xml:space="preserve">{_escape_xml(text)}</w:t></w:r>')
    rpr = run.find("w:rPr", namespaces={"w": SCHEMA})
    if rpr is not None:
        new_run.insert(0, deepcopy(rpr))
    return new_run


def _build_seq_field_runs(seq_name: str, number: str) -> list[Any]:
    """Build the run elements for a SEQ field: begin, instrText, separate, value, end."""
    ns = f'xmlns:w="{SCHEMA}"'
    return [
        parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="begin"/></w:r>'),
        parse_xml(f'<w:r {ns}><w:instrText xml:space="preserve"> SEQ {_escape_xml(seq_name)} \\* ARABIC </w:instrText></w:r>'),
        parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="separate"/></w:r>'),
        parse_xml(f"<w:r {ns}><w:t>{_escape_xml(number)}</w:t></w:r>"),
        parse_xml(f'<w:r {ns}><w:fldChar w:fldCharType="end"/></w:r>'),
    ]


def _ensure_seq_field(para: Any, seq_name: str) -> None:
    """Replace a plain-text caption number with a Word SEQ field.

    For a caption paragraph like ``Figure 1 My picture`` this replaces the
    ``1`` with ``{SEQ Figure \\* ARABIC}`` so Word can auto-number captions.
    Does nothing if the paragraph already contains a SEQ field (e.g. one
    inserted by the Lua filter).

    Only the run containing the first number is replaced; all other runs
    (with formatting, links, etc.) are preserved.
    """
    if _has_seq_field(para):
        return

    for run in list(para.findall("w:r", namespaces={"w": SCHEMA})):  # NOSONAR — list() needed: loop body mutates para
        t_el = run.find("w:t", namespaces={"w": SCHEMA})
        if t_el is None or not t_el.text:
            continue
        match = re.search(r"\d+", t_el.text)
        if not match:
            continue

        number = match.group(0)
        before = t_el.text[: match.start()]
        after = t_el.text[match.end() :]

        replacements: list[Any] = []
        if before:
            replacements.append(_clone_run_with_text(run, before))
        replacements.extend(_build_seq_field_runs(seq_name, number))
        if after:
            replacements.append(_clone_run_with_text(run, after))

        for repl in replacements:
            run.addprevious(repl)
        para.remove(run)

        logger.debug(f"Replaced number {number} with SEQ {seq_name} field in caption run")
        return


def _set_paragraph_style(para: Any, style_name: str) -> None:
    """Set (or replace) the paragraph style, creating pPr when missing."""
    p_pr = para.find(XPATH_P_PR, namespaces={"w": SCHEMA})
    if p_pr is None:
        p_pr = parse_xml(f'<w:pPr {nsdecls("w")}><w:pStyle w:val="{style_name}"/></w:pPr>')
        para.insert(0, p_pr)
        return
    p_style = p_pr.find(XPATH_P_STYLE, namespaces={"w": SCHEMA})
    if p_style is None:
        p_style = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{style_name}"/>')
        p_pr.insert(0, p_style)
    else:
        p_style.set(f"{{{SCHEMA}}}val", style_name)


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


def _parse_placeholder(text: str) -> str | None:
    """Map a placeholder paragraph text to its kind (KIND_TOC / KIND_TOF /
    KIND_TOT), or None when the text is not a placeholder."""
    return {
        TOC_PLACEHOLDER: KIND_TOC,
        TOF_PLACEHOLDER: KIND_TOF,
        TOT_PLACEHOLDER: KIND_TOT,
    }.get(text)


def _find_placeholder_paragraphs(body: Any, elements_to_replace: list) -> None:
    """Find TOC/TOF/TOT placeholder paragraphs."""
    for idx, element in enumerate(body):
        if element.tag.endswith("}p"):
            text = _get_paragraph_text(element).strip()
            style = _get_paragraph_style(element)

            # Only process placeholders in body text paragraphs, not in titles or headings
            if style not in PLACEHOLDER_PARAGRAPH_STYLES:
                continue

            kind = _parse_placeholder(text)
            if kind is not None:
                elements_to_replace.append((idx, element, kind))
                logger.info(f"Found {text} at index {idx}, will replace with a {kind} field")


def _replace_elements_with_fields(body: Any, elements_to_replace: list, figure_paragraphs: list, table_paragraphs: list) -> None:
    """Replace found elements with Word field codes."""
    # Sort by index in reverse order to maintain correct positions during removal
    for idx, element, kind in sorted(elements_to_replace, key=lambda x: x[0], reverse=True):
        body.remove(element)
        logger.debug(f"Removed element at index {idx}")

        _insert_field_at_position(body, idx, kind, figure_paragraphs, table_paragraphs)


def _insert_field_at_position(body: Any, idx: int, kind: str, figure_paragraphs: list, table_paragraphs: list) -> None:
    """Insert the Word field for one placeholder at the specified position.

    A table of figures/tables with no entries is skipped entirely (the
    placeholder is still removed by the caller)."""
    if kind == KIND_TOC:
        _insert_paragraphs(body, idx, _create_toc_field())
        logger.info(f"Inserted Table of Contents at index {idx}")
    elif kind == KIND_TOF and figure_paragraphs:
        figure_entries = [(text, bm) for _, text, bm in figure_paragraphs]
        _insert_paragraphs(body, idx, _create_tof_field(figure_entries))
        logger.info(f"Inserted Table of Figures at index {idx}")
    elif kind == KIND_TOT and table_paragraphs:
        table_entries = [(text, bm) for _, text, bm in table_paragraphs]
        _insert_paragraphs(body, idx, _create_tot_field(table_entries))
        logger.info(f"Inserted Table of Tables at index {idx}")


def _insert_paragraphs(body: Any, idx: int, paragraphs: list[Any]) -> None:
    """Insert paragraphs at the given body position, preserving their order."""
    for para in reversed(paragraphs):
        body.insert(idx, para)


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
    return _create_field(TOC_FIELD_CODE)


def _create_listing_field(field_code: str, entries: list[tuple[str, str]] | None) -> list[Any]:
    """Create a table-of listing field, pre-filled when entries are given."""
    if entries:
        return _create_field_with_entries(field_code, entries)
    return _create_field(field_code)


def _create_tof_field(entries: list[tuple[str, str]] | None = None) -> list[Any]:
    """Create Table of Figures field paragraphs using TOC \\f F."""
    return _create_listing_field(TOF_FIELD_CODE, entries)


def _create_tot_field(entries: list[tuple[str, str]] | None = None) -> list[Any]:
    """Create Table of Tables field paragraphs using TOC \\f T."""
    return _create_listing_field(TOT_FIELD_CODE, entries)


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
