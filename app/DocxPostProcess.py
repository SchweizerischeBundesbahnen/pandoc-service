import io
import logging
import sys
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.table import Table, _Cell
from lxml import etree  # type: ignore

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

EMU_1_INCH = 914400  # 1 inch in docx in EMU (English Metric Units)
TWIPS_1_INCH = 1440  # 1 inch in docx in Twips (Twentieth of a Point)
DOCX_LETTER_WIDTH_EMU = 8.5 * EMU_1_INCH  # docx LETTER width = 8.5 inch
DOCX_LETTER_SIDE_MARGIN = EMU_1_INCH  # docx left & right margins = 1 inch

# Paper sizes in TWIPS (portrait orientation: width x height)
PAPER_SIZES = {
    "A5": {"width": 8419, "height": 11906},
    "A4": {"width": 11906, "height": 16838},
    "A3": {"width": 16838, "height": 23811},
    "B5": {"width": 9979, "height": 14144},
    "B4": {"width": 14144, "height": 20013},
    "JIS_B5": {"width": 10319, "height": 14572},
    "JIS_B4": {"width": 14572, "height": 20639},
    "LETTER": {"width": 12240, "height": 15840},
    "LEGAL": {"width": 12240, "height": 20160},
    "LEDGER": {"width": 15840, "height": 24480},
}


def process(docx_bytes: bytes, paper_size: str | None = None, orientation: str | None = None) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    _replace_size_and_orientation(doc, paper_size, orientation)
    _replace_table_properties(doc)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _replace_size_and_orientation(doc: DocumentObject, paper_size: str | None = None, orientation: str | None = None) -> None:
    # If both parameters are None, no modifications needed
    if paper_size is None and orientation is None:
        return

    for section in doc.sections:
        sect_pr = section._sectPr
        pg_sz = sect_pr.find(".//w:pgSz", namespaces={"w": SCHEMA})

        if paper_size is not None:
            pg_sz = _set_paper_size(sect_pr, pg_sz, paper_size, orientation)

        if orientation is not None:
            _set_orientation(sect_pr, pg_sz, orientation)


def _set_paper_size(sect_pr: Any, pg_sz: Any, paper_size: str, orientation: str | None) -> Any:
    """Set the paper size for a section."""
    # Normalize paper_size to uppercase for case-insensitive lookup
    paper_size_upper = paper_size.upper()

    # Get dimensions for the specified paper size
    if paper_size_upper not in PAPER_SIZES:
        raise ValueError(f"Unsupported paper size: {paper_size}. Supported sizes: {', '.join(PAPER_SIZES.keys())}")

    page_dims = PAPER_SIZES[paper_size_upper]
    width = page_dims["width"]
    height = page_dims["height"]

    # Get existing orientation if present (to preserve it when changing paper size)
    existing_orientation = pg_sz.get(f"{{{SCHEMA}}}orient") if pg_sz is not None else None

    # Apply existing orientation if orientation parameter is not specified
    if orientation is None and existing_orientation == "landscape":
        width, height = height, width

    # Create or update pg_sz element
    if pg_sz is None:
        pg_sz = parse_xml(f'<w:pgSz {nsdecls("w")} w:w="{width}" w:h="{height}"/>')
        sect_pr.append(pg_sz)
    else:
        pg_sz.set(f"{{{SCHEMA}}}w", str(width))
        pg_sz.set(f"{{{SCHEMA}}}h", str(height))

    # Preserve existing orientation attribute if present and orientation parameter not specified
    if orientation is None and existing_orientation is not None:
        pg_sz.set(f"{{{SCHEMA}}}orient", existing_orientation)

    return pg_sz


def _set_orientation(sect_pr: Any, pg_sz: Any, orientation: str) -> None:
    """Set the orientation for a section."""
    # Ensure pg_sz exists (use LETTER as default if missing)
    if pg_sz is None:
        page_dims = PAPER_SIZES["LETTER"]
        width = page_dims["width"]
        height = page_dims["height"]
        pg_sz = parse_xml(f'<w:pgSz {nsdecls("w")} w:w="{width}" w:h="{height}"/>')
        sect_pr.append(pg_sz)

    # Get current dimensions
    current_width = int(pg_sz.get(f"{{{SCHEMA}}}w", "0"))
    current_height = int(pg_sz.get(f"{{{SCHEMA}}}h", "0"))

    # Determine current and desired orientation
    current_is_landscape = current_width > current_height
    desired_is_landscape = orientation.lower() == "landscape"

    # Swap dimensions if orientations don't match
    if current_is_landscape != desired_is_landscape:
        pg_sz.set(f"{{{SCHEMA}}}w", str(current_height))
        pg_sz.set(f"{{{SCHEMA}}}h", str(current_width))

    # Set or remove the orient attribute
    if desired_is_landscape:
        pg_sz.set(f"{{{SCHEMA}}}orient", "landscape")
    # Remove orient attribute for portrait (it's the default)
    elif f"{{{SCHEMA}}}orient" in pg_sz.attrib:
        del pg_sz.attrib[f"{{{SCHEMA}}}orient"]


def _replace_table_properties(doc: DocumentObject) -> None:
    max_width = _get_available_content_width(doc)

    # Process each table in the document
    for table in doc.tables:
        _process_table(table, 0, max_width)


def _process_table(table: Table, parent_columns_count: int, max_width: int) -> None:
    tbl = table._element
    columns_count = parent_columns_count + len(table.columns)
    table_properties = tbl.find(".//w:tblPr", namespaces={"w": SCHEMA})
    if table_properties is None:
        table_properties = parse_xml(f"<w:tblPr {nsdecls('w')}/>")
        tbl.insert(0, table_properties)

    # Set table width (5000 = 100%)
    table_width = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
    old_table_width = table_properties.find(".//w:tblW", namespaces={"w": SCHEMA})
    if old_table_width is not None:
        table_properties.remove(old_table_width)
    table_properties.append(table_width)

    # Set table layout to autofit
    table_layout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>')
    old_table_layout = table_properties.find(".//w:tblLayout", namespaces={"w": SCHEMA})
    if old_table_layout is not None:
        table_properties.remove(old_table_layout)
    table_properties.append(table_layout)

    # Process nested tables
    for row in table.rows:
        for cell in row.cells:
            _resize_images_in_cell(cell, max_width / columns_count)
            for sub_table in cell.tables:
                _process_table(sub_table, columns_count, max_width)


def _get_available_content_width(doc: DocumentObject) -> int:
    # Get the first section (assuming a single section document)
    section = doc.sections[0]
    # Provide alternative 'Letter' paper size params in case if they were not set explicitly in the document
    return int((section.page_width or DOCX_LETTER_WIDTH_EMU) - (section.left_margin or DOCX_LETTER_SIDE_MARGIN) - (section.right_margin or DOCX_LETTER_SIDE_MARGIN))


def _resize_images_in_cell(cell: _Cell, max_image_width: float) -> None:
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
def main() -> int:
    MIN_ARGS = 2  # script name + docx path
    MAX_ARGS = 4  # script name + docx path + paper_size + orientation
    DOCX_PATH_ARG_INDEX = 1
    PAPER_SIZE_ARG_INDEX = 2
    ORIENTATION_ARG_INDEX = 3

    if not (MIN_ARGS <= len(sys.argv) <= MAX_ARGS):
        logging.info("Usage: <path_to_docx> [paper_size] [orientation]")
        return 1

    docx_path = sys.argv[DOCX_PATH_ARG_INDEX]
    paper_size = sys.argv[PAPER_SIZE_ARG_INDEX] if len(sys.argv) > PAPER_SIZE_ARG_INDEX and sys.argv[PAPER_SIZE_ARG_INDEX] != "None" else None
    orientation = sys.argv[ORIENTATION_ARG_INDEX] if len(sys.argv) > ORIENTATION_ARG_INDEX and sys.argv[ORIENTATION_ARG_INDEX] != "None" else None

    with Path(docx_path).open("rb") as docx_file_reader:
        result_bytes = process(docx_file_reader.read(), paper_size, orientation)

    with Path(docx_path).open("wb") as docx_file_writer:
        docx_file_writer.write(result_bytes)

    logging.debug(f"Successfully modified table properties in {docx_path}")
    return 0


if __name__ == "__main__":
    main()
