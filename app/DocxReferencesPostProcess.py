import logging
from typing import Any

from docx.document import Document as DocumentObject
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# XML namespace identifier, not an actual HTTP URL
SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main" # NOSONAR

# Common XPath expressions
XPATH_P_PR = ".//w:pPr"
XPATH_P_STYLE = ".//w:pStyle"


def add_table_of_contents_entries(doc: DocumentObject) -> None:
    """
    Add Table of Contents, Table of Figures and/or Table of Tables by:
    1. Finding headings and ensuring they have proper heading styles
    2. Finding figure/table captions and adding TC fields
    3. Replacing TOC_PLACEHOLDER with Word TOC field
    4. Replacing reference link divs/lists with Word field codes
    """
    body = doc.element.body

    # Step 1: Process headings
    heading_count = _process_headings(body)
    logging.info(f"Processed {heading_count} headings")

    # Step 2: Find and process captions
    figure_paragraphs, table_paragraphs = _find_and_process_captions(body)

    # Step 3: Find elements to replace
    elements_to_replace = _find_elements_to_replace(body)

    # Step 4: Replace elements with Word fields
    _replace_elements_with_fields(body, elements_to_replace, figure_paragraphs, table_paragraphs)


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


def _find_elements_to_replace(body: Any) -> list:
    """Find all elements that need to be replaced with Word fields."""
    elements_to_replace: list = []

    # Find reference link paragraphs
    _find_reference_links(body, elements_to_replace)

    # Find TOC placeholder
    toc_index = _find_toc_placeholder(body)
    if toc_index is not None:
        element = body[toc_index]
        elements_to_replace.append((toc_index, element, True, False, False))
        logging.info(f"Found TOC_PLACEHOLDER at index {toc_index}, will replace with TOC field")
    else:
        logging.info("No TOC_PLACEHOLDER found, skipping TOC insertion")

    return elements_to_replace


def _find_reference_links(body: Any, elements_to_replace: list) -> None:
    """Find paragraphs containing reference links to figures/tables."""
    for idx, element in enumerate(body):
        if not element.tag.endswith("}p"):
            continue

        hyperlinks = element.findall(".//w:hyperlink", namespaces={"w": SCHEMA})
        has_figure_links, has_table_links = _check_hyperlinks(hyperlinks)

        if has_figure_links or has_table_links:
            elements_to_replace.append((idx, element, False, has_figure_links, has_table_links))
            logging.debug(f"Found reference link paragraph at index {idx}: figures={has_figure_links}, tables={has_table_links}")


def _check_hyperlinks(hyperlinks: list) -> tuple[bool, bool]:
    """Check if hyperlinks contain references to figures or tables."""
    has_figure_links = False
    has_table_links = False

    for hyperlink in hyperlinks:
        anchor = hyperlink.get(f"{{{SCHEMA}}}anchor", "")
        if anchor.startswith("dlecaption_"):
            link_text = _get_paragraph_text(hyperlink)
            if "Figure" in link_text:
                has_figure_links = True
            elif "Table" in link_text:
                has_table_links = True

    return has_figure_links, has_table_links


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
        toc_paragraphs = create_toc_field()
        for toc_para in reversed(toc_paragraphs):
            body.insert(idx, toc_para)
        logging.info(f"Inserted Table of Contents at index {idx}")

    if has_figure_links and figure_paragraphs:
        tof_paragraphs = create_tof_field()
        for tof_para in reversed(tof_paragraphs):
            body.insert(idx, tof_para)
        logging.info(f"Inserted Table of Figures at index {idx}")

    if has_table_links and table_paragraphs:
        tot_paragraphs = create_tot_field()
        for tot_para in reversed(tot_paragraphs):
            body.insert(idx, tot_para)
        logging.info(f"Inserted Table of Tables at index {idx}")


def _process_headings(body: Any) -> int:
    """
    Process all headings in the document to ensure they have proper Word styles.
    Pandoc converts <h1>, <h2>, <h3> to paragraphs with specific styles.
    We need to ensure these have the correct Heading1, Heading2, Heading3 styles.

    Returns: Number of headings processed
    """
    heading_count = 0

    for para in body.findall(".//w:p", namespaces={"w": SCHEMA}):
        p_pr = para.find(XPATH_P_PR, namespaces={"w": SCHEMA})
        if p_pr is not None:
            p_style = p_pr.find(XPATH_P_STYLE, namespaces={"w": SCHEMA})
            if p_style is not None:
                style_val = p_style.get(f"{{{SCHEMA}}}val", "")

                # Pandoc uses these style names for headings
                # Make sure they're set to proper Word heading styles
                if style_val in ["Heading1", "Heading2", "Heading3", "Title"]:
                    # Already has correct style
                    heading_count += 1
                    logging.debug(f"Found heading with style: {style_val}")
                elif style_val.startswith("Heading"):
                    # Has a heading style, count it
                    heading_count += 1

    return heading_count


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


def _find_toc_placeholder(body: Any) -> int | None:
    """
    Find the TOC_PLACEHOLDER marker in the document.

    Returns the index of the placeholder paragraph, or None if not found.
    """
    for idx, element in enumerate(body):
        if not element.tag.endswith("}p"):
            continue

        if _is_toc_placeholder_paragraph(element, idx):
            return idx

    return None


def _is_toc_placeholder_paragraph(element: Any, idx: int) -> bool:
    """Check if a paragraph element is a TOC_PLACEHOLDER."""
    text = _get_paragraph_text(element)
    text_stripped = text.strip()

    if text_stripped != "TOC_PLACEHOLDER":
        return False

    # Only accept placeholder in body text paragraphs, not in titles or headings
    style = _get_paragraph_style(element)
    if style not in ["BodyText", "FirstParagraph", None]:
        return False

    logging.info(f"Found TOC_PLACEHOLDER marker at index {idx}")
    return True


def _get_paragraph_style(element: Any) -> str | None:
    """Extract the style name from a paragraph element."""
    p_pr = element.find(XPATH_P_PR, namespaces={"w": SCHEMA})
    if p_pr is None:
        return None

    p_style = p_pr.find(XPATH_P_STYLE, namespaces={"w": SCHEMA})
    if p_style is None:
        return None

    return p_style.get(f"{{{SCHEMA}}}val")


def create_toc_field() -> list[Any]:
    """Create Table of Contents field paragraphs."""
    paragraphs = []
    # TOC field paragraph (standard TOC with heading levels 1-3)
    toc_para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
    </w:p>
    ''')
    paragraphs.append(toc_para)

    empty_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    paragraphs.append(empty_para)

    return paragraphs


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


def _get_paragraph_text(para: Any) -> str:
    """Extract text content from a paragraph element."""
    texts = []
    for t in para.findall(".//w:t", namespaces={"w": SCHEMA}):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


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


def create_tof_field() -> list[Any]:
    """Create Table of Figures field paragraphs."""
    paragraphs = []
    # TOC field paragraph
    toc_para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> TOC \\h \\z \\f F </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
    </w:p>
    ''')
    paragraphs.append(toc_para)

    empty_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    paragraphs.append(empty_para)

    return paragraphs


def create_tot_field() -> list[Any]:
    """Create Table of Tables field paragraphs."""
    paragraphs = []

    # TOC field paragraph (using \f T flag for tables)
    toc_para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
            <w:instrText xml:space="preserve"> TOC \\h \\z \\f T </w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end"/>
        </w:r>
    </w:p>
    ''')
    paragraphs.append(toc_para)

    empty_para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    paragraphs.append(empty_para)

    return paragraphs
