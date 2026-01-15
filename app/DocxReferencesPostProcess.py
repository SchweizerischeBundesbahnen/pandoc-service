import logging
from typing import Any

from docx.document import Document as DocumentObject
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def add_table_of_contents_entries(doc: DocumentObject) -> None:
    """
    Add Table of Contents, Table of Figures and/or Table of Tables by:
    1. Finding headings/figure/table captions
    2. Finding and replacing div(s)/ul(s) with reference links
    """
    body = doc.element.body

    # Find all figure and table captions
    figure_paragraphs = []
    table_paragraphs = []

    for para in body.findall(".//w:p", namespaces={"w": SCHEMA}):
        text = _get_paragraph_text(para)
        if text.strip().startswith("Figure"):
            figure_paragraphs.append((para, text.strip()))
        elif text.strip().startswith("Table"):
            table_paragraphs.append((para, text.strip()))

    logging.info(f"Found {len(figure_paragraphs)} figure captions and {len(table_paragraphs)} table captions")

    # Add Caption style and TC fields
    for idx, (para, caption_text) in enumerate(figure_paragraphs):
        _add_caption_style_and_tc_field(para, caption_text, field_flag="F")
        logging.debug(f"Added Caption style and TC field to figure: {caption_text}")

    for idx, (para, caption_text) in enumerate(table_paragraphs):
        logging.debug(f"Added Caption style and TC field to table: {caption_text}")
        _add_caption_style_and_tc_field(para, caption_text, field_flag="T")

    # Find paragraphs/lists containing reference links
    elements_to_replace = []  # List of (index, element, has_toc, has_figure_links, has_table_links)

    for idx, element in enumerate(body):
        # Check paragraphs
        if element.tag.endswith("}p"):
            text = _get_paragraph_text(element)
            hyperlinks = element.findall(".//w:hyperlink", namespaces={"w": SCHEMA})

            has_figure_links = False
            has_table_links = False

            for hyperlink in hyperlinks:
                anchor = hyperlink.get(f"{{{SCHEMA}}}anchor", "")
                if anchor.startswith("dlecaption_"):
                    # Check link text to determine if it's figure or table
                    link_text = _get_paragraph_text(hyperlink)
                    if "Figure" in link_text:
                        has_figure_links = True
                    elif "Table" in link_text:
                        has_table_links = True

            if has_figure_links or has_table_links:
                elements_to_replace.append((idx, element, False, has_figure_links, has_table_links))
                logging.debug(
                    f"Found reference link paragraph at index {idx}: figures={has_figure_links}, tables={has_table_links}")

    # Also search for TOC by finding <ul> elements (Pandoc may keep some structure)
    # or by finding multiple consecutive paragraphs with heading links
    toc_indices = _find_toc_structure(body)
    if toc_indices:
        # Mark ALL TOC elements for removal
        for toc_idx in toc_indices:
            # Check if not already marked
            already_marked = any(idx == toc_idx for idx, _, _, _, _ in elements_to_replace)
            if not already_marked and toc_idx < len(body):
                element = body[toc_idx]
                elements_to_replace.append((toc_idx, element, True, False, False))
        logging.info(f"Marked {len(toc_indices)} TOC elements for replacement")

    # Process elements and insert TOC/TOF/TOT
    # We need to process from end to beginning to maintain indices
    # Group consecutive TOC elements together
    toc_groups = []
    current_group = []

    for item in sorted(elements_to_replace, key=lambda x: x[0]):
        idx, element, has_toc, has_figure_links, has_table_links = item

        if has_toc:
            if not current_group or idx == current_group[-1][0] + 1:
                current_group.append(item)
            else:
                if current_group:
                    toc_groups.append(current_group)
                current_group = [item]
        else:
            if current_group:
                toc_groups.append(current_group)
                current_group = []

    if current_group:
        toc_groups.append(current_group)

    # Add non-TOC items to separate list
    non_toc_items = [item for item in elements_to_replace if not item[2]]

    # Process TOC groups (remove all elements, insert TOC at first position)
    for toc_group in reversed(toc_groups):
        first_idx = toc_group[0][0]

        # Remove all TOC elements in this group
        for idx, element, _, _, _ in reversed(toc_group):
            if element is not None:
                body.remove(element)
                logging.debug(f"Removed TOC element at index {idx}")

        toc_paragraphs = create_toc_field()
        for toc_para in reversed(toc_paragraphs):
            body.insert(first_idx, toc_para)
        logging.info(f"Inserted Table of Contents at index {first_idx}")

    # Process non-TOC elements (TOF/TOT)
    for idx, element, has_toc, has_figure_links, has_table_links in reversed(non_toc_items):
        # Remove this element if it exists
        if element is not None:
            body.remove(element)
            logging.debug(f"Removed reference element at index {idx}")

        # Insert TOF and/or TOT at this location
        insert_offset = 0

        if has_figure_links and figure_paragraphs:
            tof_paragraphs = create_tof_field()
            for tof_para in reversed(tof_paragraphs):
                body.insert(idx + insert_offset, tof_para)
            insert_offset += len(tof_paragraphs)
            logging.info(f"Inserted Table of Figures at index {idx + insert_offset - len(tof_paragraphs)}")

        if has_table_links and table_paragraphs:
            tot_paragraphs = create_tot_field()
            for tot_para in reversed(tot_paragraphs):
                body.insert(idx + insert_offset, tot_para)
            logging.info(f"Inserted Table of Tables at index {idx + insert_offset}")


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
        update_fields = settings_element.find('.//w:updateFields', namespaces={'w': SCHEMA})

        if update_fields is None:
            # Create the updateFields element
            update_fields = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
            settings_element.append(update_fields)
            logging.info("Enabled auto-update of fields on document open")
        else:
            # Update existing value
            update_fields.set(f'{{{SCHEMA}}}val', 'true')
            logging.info("Updated existing auto-update fields setting")
    except Exception as e:
        logging.warning(f"Could not enable auto-update fields: {e}")


def _find_toc_structure(body: Any) -> list[int]:
    """
    Find the TOC structure in the document.
    Pandoc converts <ul class="toc"> to a series of list items.
    We need to find ALL consecutive elements that are part of the TOC.
    Returns list of indices of ALL TOC-related elements to remove.
    """
    toc_indices = []
    in_toc = False

    for idx, element in enumerate(body):
        # Check if this element contains TOC-related links
        if element.tag.endswith("}p"):
            hyperlinks = element.findall(".//w:hyperlink", namespaces={"w": SCHEMA})
            has_toc_link = False

            for hyperlink in hyperlinks:
                anchor = hyperlink.get(f"{{{SCHEMA}}}anchor", "")
                # TOC links typically point to work-item-anchor
                if "work-item-anchor" in anchor:
                    has_toc_link = True
                    break

            if has_toc_link:
                if not in_toc:
                    # Start of TOC
                    in_toc = True
                toc_indices.append(idx)
            elif in_toc:
                # Check if this is an empty paragraph or break between TOC items
                text = _get_paragraph_text(element)
                if not text.strip():
                    # Empty paragraph, might be part of TOC formatting
                    toc_indices.append(idx)
                else:
                    # Non-TOC content, stop
                    break

    logging.info(f"Found TOC structure: indices {toc_indices}")
    return toc_indices


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
    p_pr = para.find(".//w:pPr", namespaces={"w": SCHEMA})
    if p_pr is None:
        p_pr = parse_xml(f'<w:pPr {nsdecls("w")}><w:pStyle w:val="Caption"/></w:pPr>')
        para.insert(0, p_pr)
    else:
        p_style = p_pr.find(".//w:pStyle", namespaces={"w": SCHEMA})
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