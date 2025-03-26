import io
import logging
import sys
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.table import Table, _Cell
from lxml import etree

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

EMU_1_INCH = 914400  # 1 inch in docx in EMU (English Metric Units)
DOCX_LETTER_WIDTH_EMU = 8.5 * EMU_1_INCH  # docx LETTER width = 8.5 inch
DOCX_LETTER_SIDE_MARGIN = EMU_1_INCH  # docx left & right margins = 1 inch


def replace_table_properties(docx_bytes: bytes) -> bytes:
    # Load the document
    doc = Document(io.BytesIO(docx_bytes))
    max_width = get_available_content_width(doc)

    # Process each table in the document
    for table in doc.tables:
        process_table(table, 0, max_width)

    out = io.BytesIO()
    doc.save(out)
    docx_bytes = out.getvalue()

    return docx_bytes


def process_table(table: Table, parent_columns_count: int, max_width: int) -> None:
    tbl = table._element
    columns_count = parent_columns_count + len(table.columns)
    tblPr = tbl.find(".//w:tblPr", namespaces={"w": SCHEMA})
    if tblPr is None:
        tblPr = parse_xml(f"<w:tblPr {nsdecls('w')}/>")
        tbl.insert(0, tblPr)

    # Set table width (5000 = 100%)
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
    old_tblW = tblPr.find(".//w:tblW", namespaces={"w": SCHEMA})
    if old_tblW is not None:
        tblPr.remove(old_tblW)
    tblPr.append(tblW)

    # Set table layout to autofit
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>')
    old_tblLayout = tblPr.find(".//w:tblLayout", namespaces={"w": SCHEMA})
    if old_tblLayout is not None:
        tblPr.remove(old_tblLayout)
    tblPr.append(tblLayout)

    # Process nested tables
    for row in table.rows:
        for cell in row.cells:
            resize_images_in_cell(cell, max_width / columns_count)
            for subtable in cell.tables:
                process_table(subtable, columns_count, max_width)


def get_available_content_width(doc: DocumentObject) -> int:
    # Get the first section (assuming a single section document)
    section = doc.sections[0]
    # Provide alternative 'Letter' paper size params in case if they were not set explicitly in the document
    return int((section.page_width or DOCX_LETTER_WIDTH_EMU) - (section.left_margin or DOCX_LETTER_SIDE_MARGIN) - (section.right_margin or DOCX_LETTER_SIDE_MARGIN))


def resize_images_in_cell(cell: _Cell, max_image_width: float) -> None:
    cell_xml = cell._tc.xml
    # ruff: noqa: S320
    tree = etree.fromstring(cell_xml)

    # Find all <wp:extent> elements that define image size
    extent_elements = tree.findall(
        ".//wp:extent",
        {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"},
    )

    modified = False  # Track if any image was resized

    for extent in extent_elements:
        width = int(extent.attrib["cx"])
        height = int(extent.attrib["cy"])
        logging.debug(f"Image found, size: {width} x {height}")

        # Resize only if width exceeds max_image_width
        if width > max_image_width:
            scale_factor = max_image_width / width
            new_width = int(max_image_width)
            new_height = int(height * scale_factor)  # Maintain aspect ratio

            # Apply new size
            extent.set("cx", str(new_width))
            extent.set("cy", str(new_height))

            logging.debug(f"Resized to: {new_width} x {new_height}")
            modified = True

    # If any modification was made, update the cell XML
    if modified:
        cell._tc.clear_content()
        for child in tree.iterchildren():
            cell._tc.append(child)


# Just for manual test purposes. Accepts path to docx to process.
if __name__ == "__main__":
    args_number = 2

    if len(sys.argv) != args_number:
        logging.info("Provide <path_to_docx>")
        sys.exit(1)

    docx_path = sys.argv[1]

    with Path(docx_path).open("rb") as docx_file_reader:
        result_bytes = replace_table_properties(docx_file_reader.read())

    with Path(docx_path).open("wb") as docx_file_writer:
        docx_file_writer.write(result_bytes)

    logging.debug(f"Successfully modified table properties in {docx_path}")
