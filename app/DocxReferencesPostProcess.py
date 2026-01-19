import logging
from typing import Any

from docx.document import Document as DocumentObject
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# XML namespace identifier, not an actual HTTP URL
SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR

# Common XPath expressions
XPATH_P_PR = ".//w:pPr"
XPATH_P_STYLE = ".//w:pStyle"


def add_table_of_contents_entries(doc: DocumentObject) -> None:
    """
    Add Table of Contents, Table of Figures and/or Table of Tables by:
    1. Finding figure/table captions and adding TC fields
    2. Replacing TOC_PLACEHOLDER with Word TOC field
    3. Replacing TOF_PLACEHOLDER with Word TOF field
    4. Replacing TOT_PLACEHOLDER with Word TOT field
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
        # Access the settings part
        settings = doc.settings
        settings_element = settings.element

        # Check if updateFields already exists
        update_fields = settings_element.find(".//w:updateFields", namespaces={"w": SCHEMA})

        if update_fields is None:
            # Create the updateFields element
            update_fields = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
            settings_element.append(update_fields)
            logging.info("Enabled auto-update of fields on document open")
        else:
            # Update existing value
            update_fields.set(f"{{{SCHEMA}}}val", "true")
            logging.info("Updated existing auto-update fields setting")
    except Exception as e:
        logging.warning(f"Could not enable auto-update fields: {e}")


def _find_and_process_captions(body: Any) -> tuple[list, list]:
    """Find all figure and table captions and add TC fields to them."""
    figure_paragraphs = []
    table_paragraphs = []

    for para in body.findall(".//w:p", namespaces={"w": SCHEMA}):
        text = _get_paragraph_text(para)
        text_stripped = text.strip()

        if text_stripped.startswith("Figure"):
            figure_paragraphs.append((para, text_stripped))
            _add_caption_style_and_tc_field(para, text_stripped, field_flag="F")
            logging.debug(f"Added Caption style and TC field to figure: {text_stripped}")
        elif text_stripped.startswith("Table"):
            table_paragraphs.append((para, text_stripped))
            _add_caption_style_and_tc_field(para, text_stripped, field_flag="T")
            logging.debug(f"Added Caption style and TC field to table: {text_stripped}")

    logging.info(f"Found {len(figure_paragraphs)} figure captions and {len(table_paragraphs)} table captions")
    return figure_paragraphs, table_paragraphs


def _add_caption_style_and_tc_field(para: Any, caption_text: str, field_flag: str) -> None:
    """Add Caption style and TC field to a paragraph."""
    # Add or update paragraph style to Caption
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

    # Add TC field runs at the end of paragraph
    tc_runs = _create_tc_field_runs(caption_text, field_flag=field_flag)
    for run in tc_runs:
        para.append(run)


def _find_elements_to_replace(body: Any) -> list:
    """Find all elements that need to be replaced with Word fields."""
    elements_to_replace: list = []

    # Find TOC/TOF/TOT placeholder paragraphs
    _find_placeholder_paragraphs(body, elements_to_replace)

    return elements_to_replace


def _find_placeholder_paragraphs(body: Any, elements_to_replace: list) -> None:
    """Find TOC_PLACEHOLDER, TOF_PLACEHOLDER, and TOT_PLACEHOLDER paragraphs."""
    for idx, element in enumerate(body):
        # Check for placeholders in paragraphs
        if element.tag.endswith("}p"):
            text = _get_paragraph_text(element).strip()
            style = _get_paragraph_style(element)

            # Only process placeholders in body text paragraphs, not in titles or headings
            if style not in ["BodyText", "FirstParagraph", None]:
                continue

            if text == "TOC_PLACEHOLDER":
                elements_to_replace.append((idx, element, True, False, False))
                logging.info(f"Found TOC_PLACEHOLDER at index {idx}, will replace with TOC field")
            elif text == "TOF_PLACEHOLDER":
                elements_to_replace.append((idx, element, False, True, False))
                logging.info(f"Found TOF_PLACEHOLDER at index {idx}, will replace with TOF field")
            elif text == "TOT_PLACEHOLDER":
                elements_to_replace.append((idx, element, False, False, True))
                logging.info(f"Found TOT_PLACEHOLDER at index {idx}, will replace with TOT field")


def _replace_elements_with_fields(body: Any, elements_to_replace: list, figure_paragraphs: list, table_paragraphs: list) -> None:
    """Replace found elements with Word field codes."""
    # Sort by index in reverse order to maintain correct positions during removal
    for idx, element, has_toc, has_figure_links, has_table_links in sorted(elements_to_replace, key=lambda x: x[0], reverse=True):
        # Remove element
        if element is not None:
            body.remove(element)
            logging.debug(f"Removed element at index {idx}")

        # Insert replacement fields
        _insert_field_at_position(body, idx, has_toc, has_figure_links, has_table_links, figure_paragraphs, table_paragraphs)


def _insert_field_at_position(body: Any, idx: int, has_toc: bool, has_figure_links: bool, has_table_links: bool, figure_paragraphs: list, table_paragraphs: list) -> None:  # noqa: PLR0913
    """Insert appropriate Word fields at the specified position."""
    if has_toc:
        toc_paragraphs = _create_toc_field()
        for toc_para in reversed(toc_paragraphs):
            body.insert(idx, toc_para)
        logging.info(f"Inserted Table of Contents at index {idx}")

    if has_figure_links and figure_paragraphs:
        tof_paragraphs = _create_tof_field()
        for tof_para in reversed(tof_paragraphs):
            body.insert(idx, tof_para)
        logging.info(f"Inserted Table of Figures at index {idx}")

    if has_table_links and table_paragraphs:
        tot_paragraphs = _create_tot_field()
        for tot_para in reversed(tot_paragraphs):
            body.insert(idx, tot_para)
        logging.info(f"Inserted Table of Tables at index {idx}")


def _create_field(field_code: str) -> list[Any]:
    """Create a Word field with specified field code.

    Args:
        field_code: The field instruction text (e.g., 'TOC \\o "1-3" \\h \\z \\u')

    Returns:
        List of paragraph elements containing the field
    """
    paragraphs = []

    # Field paragraph
    field_para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> {field_code} </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
    </w:p>
    ''')
    paragraphs.append(field_para)

    # Add empty paragraph for spacing
    empty_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    paragraphs.append(empty_para)

    return paragraphs


def _create_toc_field() -> list[Any]:
    """Create Table of Contents field paragraphs."""
    return _create_field('TOC \\o "1-3" \\h \\z \\u')


def _create_table_listing_field(field_type: str) -> list[Any]:
    """Create TOC-based field paragraphs (TOF or TOT).

    Args:
        field_type: "F" for figures, "T" for tables

    Returns:
        List of paragraph elements
    """
    return _create_field(f"TOC \\h \\z \\f {field_type}")


def _create_tof_field() -> list[Any]:
    """Create Table of Figures field paragraphs."""
    return _create_table_listing_field("F")


def _create_tot_field() -> list[Any]:
    """Create Table of Tables field paragraphs."""
    return _create_table_listing_field("T")


def _create_tc_field_runs(caption_text: str, field_flag: str = "F") -> list[Any]:
    """Create TC (Table of Contents Entry) field runs for a caption.

    Args:
        caption_text: The caption text (e.g., "Figure 1", "Table 1")
        field_flag: The field flag - "F" for figures, "T" for tables
    """
    runs = []

    # Begin run
    begin_run = parse_xml(f'<w:r xmlns:w="{SCHEMA}"><w:fldChar w:fldCharType="begin"/></w:r>')
    runs.append(begin_run)

    # Instruction run
    instr_run = parse_xml(f'''<w:r xmlns:w="{SCHEMA}">
        <w:instrText xml:space="preserve"> TC "{caption_text}" \\f {field_flag} \\l "1" </w:instrText>
    </w:r>''')
    runs.append(instr_run)

    # End run
    end_run = parse_xml(f'<w:r xmlns:w="{SCHEMA}"><w:fldChar w:fldCharType="end"/></w:r>')
    runs.append(end_run)

    return runs


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
