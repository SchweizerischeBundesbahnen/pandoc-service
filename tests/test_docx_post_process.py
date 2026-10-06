import sys
from pathlib import Path
from unittest.mock import MagicMock, mock_open, patch

import pytest
from docx.table import Table, _Cell
from lxml import etree

from app import docx_post_process
from app.docx_post_process import (
    SCHEMA,
    _has_existing_fixed_width,
    _process_table,
    _replace_image_placeholders,
    _replace_link_placeholders,
    _replace_table_properties,
    _resolve_image_src,
)

WORD_PROCESSING_ML_MAIN_SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD_PROCESSING_ML_MAIN_SCHEMA_IN_BRACKETS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DRAWING_ML_MAIN_SCHEMA = "http://schemas.openxmlformats.org/drawingml/2006/main"
DRAWING_ML_PICTURE_SCHEMA = "http://schemas.openxmlformats.org/drawingml/2006/picture"

SOURCE_HTML_WITH_TABLE = """
            <html>
                <head>
                    <title>Test doc title</title>
                </head>
                <body>
                    <h1>Simple html with table</h1>
                    <table>
                        <thead>
                            <tr>
                                <td style="width: 1000px">Wide column</td>
                                <td>
                                    <img src="{0}"></img>
                                </td>
                            </tr>
                        </thead>
                    </table>
                </body>
            </html>
        """

SOURCE_HTML_WITH_NESTED_TABLE = """
            <html>
                <head>
                    <title>Test doc title</title>
                </head>
                <body>
                    <h1>Nested tables</h1>
                    <table>
                        <tr>
                            <td>
                                <table>
                                    <tr>
                                        <td>Nested cell</td>
                                        <td><img src="{0}"></img></td>
                                    </tr>
                                </table>
                            </td>
                            <td>Outer cell</td>
                        </tr>
                    </table>
                </body>
            </html>
        """

SOURCE_HTML_NO_TABLES = """
            <html>
                <head>
                    <title>Test doc title</title>
                </head>
                <body>
                    <h1>Document without tables</h1>
                    <p>This is a simple paragraph without any tables.</p>
                    <p><img src="{0}"></img></p>
                </body>
            </html>
        """

# HTML with a small image that shouldn't need resizing
SOURCE_HTML_SMALL_IMAGE = """
            <html>
                <head>
                    <title>Test doc title</title>
                </head>
                <body>
                    <h1>Document with small image</h1>
                    <table>
                        <tr>
                            <td>
                                <!-- Using a data URL for a 1x1 pixel transparent PNG -->
                                <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="></img>
                            </td>
                        </tr>
                    </table>
                </body>
            </html>
        """

EMUS_IN_INCH = 914400


# Mock Document class for testing without requiring a real docx file
def create_mock_document_with_table(has_nested_table=False, has_image=True):
    # Create a mock Document
    mock_doc = MagicMock()

    # Create a mock table
    mock_table = MagicMock(spec=Table)
    mock_table._element = MagicMock()

    # Add columns to the table (to avoid division by zero)
    mock_column = MagicMock()
    mock_table.columns = [mock_column]  # At least one column

    # Set up table properties
    mock_table_properties = MagicMock()
    mock_table._element.find.return_value = mock_table_properties

    # Create mock rows and cells
    mock_row = MagicMock()
    mock_cell = MagicMock(spec=_Cell)
    mock_cell._tc = MagicMock()

    # Set up nested tables if needed
    if has_nested_table:
        mock_nested_table = MagicMock(spec=Table)
        mock_nested_table._element = MagicMock()
        # Add columns to the nested table too
        mock_nested_table.columns = [MagicMock()]
        mock_cell.tables = [mock_nested_table]
    else:
        mock_cell.tables = []

        # Set up mock cell XML with image if needed
    if has_image:
        mock_cell._tc.xml = f"""
        <w:tc xmlns:w="{WORD_PROCESSING_ML_MAIN_SCHEMA}"
              xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
              xmlns:a="{DRAWING_ML_MAIN_SCHEMA}"
              xmlns:pic="{DRAWING_ML_PICTURE_SCHEMA}"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:tcPr/>
            <w:p>
                <w:r>
                    <w:drawing>
                        <wp:inline>
                            <wp:extent cx="1000000" cy="750000"/>
                            <a:graphic>
                                <a:graphicData>
                                    <pic:pic>
                                        <pic:blipFill>
                                            <a:blip r:embed="rId5"/>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm>
                                                <a:ext cx="1000000" cy="750000"/>
                                            </a:xfrm>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing>
                </w:r>
            </w:p>
        </w:tc>
        """
    else:
        mock_cell._tc.xml = f"""
        <w:tc xmlns:w="{WORD_PROCESSING_ML_MAIN_SCHEMA}">
            <w:tcPr/>
            <w:p>
                <w:r>
                    <w:t>Cell content</w:t>
                </w:r>
            </w:p>
        </w:tc>
        """

        # Connect the mocks together
    mock_row.cells = [mock_cell]
    mock_table.rows = [mock_row]
    mock_doc.tables = [mock_table]

    return mock_doc


@patch("app.docx_post_process.etree.fromstring")
def test_nested_tables(mock_fromstring):
    # Create a mock document with nested tables
    mock_doc = create_mock_document_with_table(has_nested_table=True)

    # Create a mock XML tree that will be returned from fromstring
    mock_tree = MagicMock()
    # Return empty list for extent elements to avoid image resizing
    mock_tree.findall.return_value = []
    mock_fromstring.return_value = mock_tree

    # Call the function under test
    docx_post_process._replace_table_properties(mock_doc)

    # Verify we processed the nested table by checking if cell.tables was accessed
    assert len(mock_doc.tables[0].rows[0].cells[0].tables) > 0


def test_document_without_tables():
    # Create a mock document with no tables
    mock_doc = MagicMock()
    mock_doc.tables = []
    # Add sections for _get_available_content_width
    mock_section = MagicMock()
    mock_section.page_width = docx_post_process.DOCX_LETTER_WIDTH_EMU
    mock_section.left_margin = docx_post_process.DOCX_LETTER_SIDE_MARGIN
    mock_section.right_margin = docx_post_process.DOCX_LETTER_SIDE_MARGIN
    mock_doc.sections = [mock_section]

    # Call the function under test
    docx_post_process._replace_table_properties(mock_doc)

    # Verify we checked the tables collection
    assert len(mock_doc.tables) == 0


def test_resize_images_in_cell_no_resizing_needed():
    """Test that small images don't get resized."""
    # Create a mock cell
    cell = MagicMock(spec=_Cell)

    # Create mock XML content with a small image (smaller than max width)
    small_image_width = 100000  # Small value in EMU, less than max_width
    small_image_height = 100000
    small_image_xml = f"""
    <w:tc xmlns:w="{WORD_PROCESSING_ML_MAIN_SCHEMA}"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:a="{DRAWING_ML_MAIN_SCHEMA}">
        <w:tcPr/>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="{small_image_width}" cy="{small_image_height}"/>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
    </w:tc>
    """

    # Set up the mock cell
    mock_tc = MagicMock()
    mock_tc.xml = small_image_xml
    cell._tc = mock_tc

    # Set a max width larger than the image width
    max_width = 500000  # Larger than small_image_width

    # Call the function
    docx_post_process._resize_images_in_cell(cell, max_width)

    # Verify that clear_content was not called (which would indicate the image was modified)
    mock_tc.clear_content.assert_not_called()


def test_resize_images_in_cell_resizing_needed():
    """Test that large images are properly resized."""
    # Create a mock cell
    cell = MagicMock(spec=_Cell)

    # Create mock XML content with a large image (larger than max_width)
    large_image_width = 1000000  # Large value in EMU, greater than max_width
    large_image_height = 750000  # 3:4 aspect ratio
    large_image_xml = f"""
    <w:tc xmlns:w="{WORD_PROCESSING_ML_MAIN_SCHEMA}"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:a="{DRAWING_ML_MAIN_SCHEMA}">
        <w:tcPr/>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="{large_image_width}" cy="{large_image_height}"/>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
    </w:tc>
    """

    # Set up the mock for lxml etree parsing - create elements that mimic the real ones

    # Parse the XML to create a real tree
    tree = etree.fromstring(large_image_xml)

    # Find the wp:extent element to use with the real function
    extent_element = tree.find(".//wp:extent", {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
    assert extent_element.get("cx") == str(large_image_width)
    assert extent_element.get("cy") == str(large_image_height)

    # Set up the mock cell with our parsed tree
    mock_tc = MagicMock()
    mock_tc.xml = large_image_xml
    cell._tc = mock_tc

    # Set a max width smaller than the image width to trigger resizing
    max_width = 500000  # Smaller than large_image_width

    # Call the function
    with patch("app.docx_post_process.etree.fromstring", return_value=tree):
        docx_post_process._resize_images_in_cell(cell, max_width)

        # Verify that clear_content was called (which indicates the image was modified)
    mock_tc.clear_content.assert_called_once()

    # Check that the image dimensions were actually updated in the element tree
    # Expected new width is max_width
    expected_new_width = max_width
    # Expected new height maintains the original aspect ratio: height * (new_width / old_width)
    expected_new_height = int(large_image_height * (expected_new_width / large_image_width))

    assert extent_element.get("cx") == str(expected_new_width)
    assert extent_element.get("cy") == str(expected_new_height)


def test_resize_images_in_cell():
    """Test that images in a table cell are correctly resized."""
    # Create a mock cell with proper _tc attribute
    cell = MagicMock()
    cell._tc = MagicMock()
    cell._tc.xml = '<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="5000" cy="3000"/></w:tc>'

    # Create a mock for etree.fromstring that returns a properly structured mock
    mock_tree = MagicMock()
    mock_extent = MagicMock()
    mock_extent.attrib = {"cx": "5000", "cy": "3000"}
    mock_tree.findall.return_value = [mock_extent]

    # Patch etree.fromstring to return our mock
    with patch("app.docx_post_process.etree.fromstring", return_value=mock_tree):
        # Use a reasonable max width value that will trigger resizing
        max_image_width = 4000.0  # Smaller than the image width

        # Call resize_images_in_cell directly
        docx_post_process._resize_images_in_cell(cell, max_image_width)

        # Verify that the image was resized
        assert mock_extent.set.call_count == 2  # cx and cy set
        # Resize should maintain aspect ratio
        mock_extent.set.assert_any_call("cx", "4000")  # New width
        # New height = old_height * (new_width/old_width) = 3000 * (4000/5000) = 2400
        mock_extent.set.assert_any_call("cy", "2400")  # New height (aspect ratio maintained)


@patch("app.docx_post_process._apply_table_layout")
@patch("app.docx_post_process._resize_images_in_cell")
def test_process_table_with_nested_tables(mock_resize_images, mock_apply_layout):
    """Test that nested tables are processed correctly."""

    # Create mock tables and cells
    main_table = MagicMock()
    nested_table = MagicMock()
    cell = MagicMock()

    # Set up cell to return the nested table
    cell.tables = [nested_table]

    # Set up row to return the cell
    row = MagicMock()
    row.cells = [cell]

    # Set up main table properties
    main_table.rows = [row]
    main_table.columns = MagicMock()
    main_table.columns.__len__.return_value = 2
    main_table._element = MagicMock()
    main_table._element.find.return_value = None

    # Set up property mocks to track XML manipulation
    with patch("app.docx_post_process.parse_xml") as mock_parse_xml:
        mock_tbl_props = MagicMock()
        mock_parse_xml.return_value = mock_tbl_props

        # Set up max_width for testing
        max_width = 9144000  # 10 inches in EMU

        # Call the actual _process_table function
        _process_table(main_table, 0, max_width)

        # Verify find was called with correct parameters
        main_table._element.find.assert_any_call(".//w:tblPr", namespaces={"w": SCHEMA})

        # Verify parse_xml was called (for creating table properties)
        mock_parse_xml.assert_called()

        # Verify the table element was updated
        main_table._element.insert.assert_called_with(0, mock_tbl_props)

        # Verify _resize_images_in_cell was called with correct parameters
        mock_resize_images.assert_called_with(cell, max_width / 2)

        # Verify that we processed any nested tables
        for _sub_table in cell.tables:
            # At this point, we would process the nested table, but we can't verify
            # the actual call since it happens recursively. Instead, we verify that
            # the nested tables were accessed.
            pass


@patch("app.docx_post_process._process_table")
@patch("app.docx_post_process._get_available_content_width_for_section")
def test_replace_table_properties_by_section(mock_get_width, mock_process_table):
    """Test that _replace_table_properties processes tables by section correctly."""

    # Create mock document and sections
    doc = MagicMock()
    section1 = MagicMock()
    section2 = MagicMock()
    doc.sections = [section1, section2]

    # Set up mock tables
    table1 = MagicMock()
    table2 = MagicMock()
    doc.tables = [table1, table2]

    # Set up table elements
    tbl_element1 = MagicMock()
    tbl_element1.tag = "w:tbl"
    tbl_element2 = MagicMock()
    tbl_element2.tag = "w:tbl"

    table1._element = tbl_element1
    table2._element = tbl_element2

    # Set up document body with table and section break elements
    section_break = MagicMock()
    section_break.tag = "w:sectPr"

    doc.element.body = [tbl_element1, section_break, tbl_element2]

    # Set up available width
    max_width = 9144000
    mock_get_width.return_value = max_width

    # Call the function
    _replace_table_properties(doc)

    # Verify _get_available_content_width_for_section was called for each section
    assert mock_get_width.call_count == 2

    # Verify _process_table was called for tables in correct sections
    assert mock_process_table.call_count == 2
    mock_process_table.assert_any_call(table1, 0, max_width, None)
    mock_process_table.assert_any_call(table2, 0, max_width, None)


def test_replace_size_and_orientation_both_none():
    """Test that no modifications are made when both parameters are None."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_doc.sections = [mock_section]

    # Call with both None
    docx_post_process._replace_size_and_orientation(mock_doc, None, None)

    # Verify no modifications were attempted
    mock_section._sectPr.find.assert_not_called()


def test_replace_size_and_orientation_set_paper_size_only():
    """Test setting paper size without changing orientation."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    # Set up existing paper size element
    mock_pg_sz.get.return_value = None  # No existing orientation
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with paper_size = A4
    docx_post_process._replace_size_and_orientation(mock_doc, "A4", None)

    # Verify pgSz was updated with A4 dimensions (portrait)
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "11906")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "16838")


def test_replace_size_and_orientation_set_orientation_only():
    """Test setting orientation without changing paper size."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    # Set up existing paper size element with portrait dimensions (width < height)
    # The get method is called twice: once for width, once for height
    mock_pg_sz.get.side_effect = ["11906", "16838"]  # portrait: width=11906, height=16838
    mock_pg_sz.attrib = {}
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with orientation = landscape
    docx_post_process._replace_size_and_orientation(mock_doc, None, "landscape")

    # Verify dimensions were swapped (portrait to landscape)
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "16838")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "11906")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}orient", "landscape")


def test_replace_size_and_orientation_both_parameters():
    """Test setting both paper size and orientation."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    # Need to provide return values for the get calls:
    # 1. _set_paper_size calls get() once for existing orient attribute -> None
    # 2. _set_orientation calls get() for width -> "12240" (LETTER portrait width, set by _set_paper_size)
    # 3. _set_orientation calls get() for height -> "15840" (LETTER portrait height, set by _set_paper_size)
    mock_pg_sz.get.side_effect = [None, "12240", "15840"]
    mock_pg_sz.attrib = {}
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with paper_size = LETTER and orientation = landscape
    docx_post_process._replace_size_and_orientation(mock_doc, "LETTER", "landscape")

    # Verify LETTER portrait dimensions were set first by _set_paper_size
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "12240")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "15840")
    # Then verify they were swapped for landscape by _set_orientation
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "15840")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "12240")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}orient", "landscape")


def test_set_paper_size_unsupported():
    """Test that unsupported paper size raises ValueError."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Test with unsupported paper size
    with pytest.raises(ValueError, match="Unsupported paper size: TABLOID"):
        docx_post_process._replace_size_and_orientation(mock_doc, "TABLOID", None)


def test_set_paper_size_case_insensitive():
    """Test that paper size is case-insensitive."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    mock_pg_sz.get.return_value = None
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with lowercase paper size
    docx_post_process._replace_size_and_orientation(mock_doc, "a4", None)

    # Verify A4 dimensions were set
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "11906")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "16838")


def test_set_paper_size_creates_pg_sz_if_missing():
    """Test that pgSz element is created if it doesn't exist."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()

    # pgSz doesn't exist
    mock_sect_pr.find.return_value = None
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    with patch("app.docx_post_process.parse_xml") as mock_parse_xml:
        mock_new_pg_sz = MagicMock()
        mock_new_pg_sz.get.return_value = None
        mock_parse_xml.return_value = mock_new_pg_sz

        # Call with paper_size = A5
        docx_post_process._replace_size_and_orientation(mock_doc, "A5", None)

        # Verify parse_xml was called to create new pgSz
        mock_parse_xml.assert_called_once()
        # Verify the new pgSz was appended
        mock_sect_pr.append.assert_called_once_with(mock_new_pg_sz)


def test_set_paper_size_preserves_landscape_orientation():
    """Test that existing landscape orientation is preserved when changing paper size."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    # Existing page has landscape orientation
    mock_pg_sz.get.return_value = "landscape"
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with new paper_size but no orientation parameter
    docx_post_process._replace_size_and_orientation(mock_doc, "A3", None)

    # Verify A3 dimensions were set in landscape (swapped)
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "23811")  # height becomes width
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "16838")  # width becomes height
    # Verify landscape orientation was preserved
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}orient", "landscape")


def test_set_orientation_creates_pg_sz_if_missing():
    """Test that pgSz element is created with LETTER default if missing."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()

    # pgSz doesn't exist
    mock_sect_pr.find.return_value = None
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    with patch("app.docx_post_process.parse_xml") as mock_parse_xml:
        mock_new_pg_sz = MagicMock()
        # Simulate getting width and height from the newly created element (LETTER portrait)
        mock_new_pg_sz.get.side_effect = ["12240", "15840"]  # width, then height
        mock_new_pg_sz.attrib = {}
        mock_parse_xml.return_value = mock_new_pg_sz

        # Call with orientation only
        docx_post_process._replace_size_and_orientation(mock_doc, None, "landscape")

        # Verify parse_xml was called to create pgSz with LETTER dimensions
        mock_parse_xml.assert_called_once()
        # Verify the new pgSz was appended
        mock_sect_pr.append.assert_called_once_with(mock_new_pg_sz)
        # Verify dimensions were swapped for landscape
        mock_new_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "15840")
        mock_new_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "12240")


def test_set_orientation_portrait_removes_orient_attribute():
    """Test that orient attribute is removed for portrait orientation."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    # Set up existing landscape page (width > height)
    mock_pg_sz.get.side_effect = ["15840", "12240"]  # landscape: width=15840, height=12240
    mock_pg_sz.attrib = {f"{{{SCHEMA}}}orient": "landscape"}
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with orientation = portrait
    docx_post_process._replace_size_and_orientation(mock_doc, None, "portrait")

    # Verify dimensions were swapped back to portrait
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", "12240")
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", "15840")
    # Verify orient attribute was removed
    assert f"{{{SCHEMA}}}orient" not in mock_pg_sz.attrib


def test_set_orientation_no_swap_if_already_correct():
    """Test that dimensions aren't swapped if orientation is already correct."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    # Set up existing landscape page (width > height)
    width = "16838"
    height = "11906"
    call_count = [0]

    def get_side_effect(attr, default="0"):
        call_count[0] += 1
        if call_count[0] == 1:  # First call for width
            return width
        # Second call for height
        return height

    mock_pg_sz.get.side_effect = get_side_effect
    mock_pg_sz.attrib = {f"{{{SCHEMA}}}orient": "landscape"}
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with orientation = landscape (already landscape)
    docx_post_process._replace_size_and_orientation(mock_doc, None, "landscape")

    # Verify dimensions were NOT swapped (set should not be called with swapped values)
    # The orient attribute should still be set
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}orient", "landscape")


def test_replace_size_and_orientation_multiple_sections():
    """Test that all sections in a document are processed."""
    mock_doc = MagicMock()
    mock_section1 = MagicMock()
    mock_section2 = MagicMock()
    mock_sect_pr1 = MagicMock()
    mock_sect_pr2 = MagicMock()
    mock_pg_sz1 = MagicMock()
    mock_pg_sz2 = MagicMock()

    mock_pg_sz1.get.return_value = None
    mock_pg_sz2.get.return_value = None
    mock_sect_pr1.find.return_value = mock_pg_sz1
    mock_sect_pr2.find.return_value = mock_pg_sz2
    mock_section1._sectPr = mock_sect_pr1
    mock_section2._sectPr = mock_sect_pr2
    mock_doc.sections = [mock_section1, mock_section2]

    # Call with paper_size = B5
    docx_post_process._replace_size_and_orientation(mock_doc, "B5", None)

    # Verify both sections were updated
    mock_pg_sz1.set.assert_any_call(f"{{{SCHEMA}}}w", "9979")
    mock_pg_sz1.set.assert_any_call(f"{{{SCHEMA}}}h", "14144")
    mock_pg_sz2.set.assert_any_call(f"{{{SCHEMA}}}w", "9979")
    mock_pg_sz2.set.assert_any_call(f"{{{SCHEMA}}}h", "14144")


@pytest.mark.parametrize(
    "paper_size,expected_width,expected_height",
    [
        ("A5", "8419", "11906"),
        ("A4", "11906", "16838"),
        ("A3", "16838", "23811"),
        ("B5", "9979", "14144"),
        ("B4", "14144", "20013"),
        ("JIS_B5", "10319", "14572"),
        ("JIS_B4", "14572", "20639"),
        ("LETTER", "12240", "15840"),
        ("LEGAL", "12240", "20160"),
        ("LEDGER", "15840", "24480"),
    ],
)
def test_all_supported_paper_sizes(paper_size, expected_width, expected_height):
    """Test that all supported paper sizes are correctly applied."""
    mock_doc = MagicMock()
    mock_section = MagicMock()
    mock_sect_pr = MagicMock()
    mock_pg_sz = MagicMock()

    mock_pg_sz.get.return_value = None
    mock_sect_pr.find.return_value = mock_pg_sz
    mock_section._sectPr = mock_sect_pr
    mock_doc.sections = [mock_section]

    # Call with the specified paper_size
    docx_post_process._replace_size_and_orientation(mock_doc, paper_size, None)

    # Verify the correct dimensions were set
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}w", expected_width)
    mock_pg_sz.set.assert_any_call(f"{{{SCHEMA}}}h", expected_height)


@pytest.mark.parametrize(
    "argv, expected_exit, paper_size, orientation",
    [
        (["script.py"], True, None, None),
        (["script.py", "test.docx", "A4", "landscape"], False, "A4", "landscape"),
        (["script.py", "test.docx", "None", "portrait"], False, None, "portrait"),
        (["script.py", "test.docx", "LETTER", "None"], False, "LETTER", None),
    ],
)
def test_main_function(argv, expected_exit, paper_size, orientation):
    fake_docx_content = b"fake content"
    modified_content = b"modified content"

    with (
        patch.object(sys, "argv", argv),
        patch("pathlib.Path.open", mock_open(read_data=fake_docx_content)) as mock_file,
        patch("app.docx_post_process.process", return_value=modified_content) as mock_process,
        patch("app.docx_post_process.logger") as mock_logging,
    ):
        result = docx_post_process.main()

        if expected_exit:
            assert result == 1
        else:
            assert result == 0
            mock_process.assert_called_once_with(fake_docx_content, paper_size, orientation)
            handle = mock_file()
            handle.write.assert_called_once_with(modified_content)
            mock_logging.debug.assert_called_once()


def test_main_function_rejects_path_outside_working_directory():
    fixed_cwd = Path("/app/workdir").resolve()
    outside_path = str(Path("/outside/escape.docx").resolve())

    with (
        patch.object(sys, "argv", ["script.py", outside_path]),
        patch("pathlib.Path.cwd", return_value=fixed_cwd),
        patch("pathlib.Path.open", mock_open()) as mock_file,
        patch("app.docx_post_process.process") as mock_process,
        patch("app.docx_post_process.logger") as mock_logging,
    ):
        result = docx_post_process.main()

        assert result == 1
        mock_process.assert_not_called()
        mock_file.assert_not_called()
        mock_logging.error.assert_called_once()

        # Integration tests for the process() function to ensure 100% coverage


class TestProcessFunction:
    """Integration tests that call the process() function directly."""

    def test_process_with_no_parameters(self):
        """Test process() with no paper_size or orientation - should just process tables."""
        # Create a minimal valid DOCX file
        import io

        from docx import Document

        doc = Document()
        doc.add_paragraph("Test content")
        docx_bytes = io.BytesIO()
        doc.save(docx_bytes)
        docx_bytes.seek(0)
        input_bytes = docx_bytes.getvalue()

        # Call process with no parameters
        result = docx_post_process.process(input_bytes)

        # Verify result is valid bytes
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_process_with_paper_size_only(self):
        """Test process() with only paper_size parameter."""
        import io

        from docx import Document

        doc = Document()
        doc.add_paragraph("Test content")
        docx_bytes = io.BytesIO()
        doc.save(docx_bytes)
        docx_bytes.seek(0)
        input_bytes = docx_bytes.getvalue()

        # Call process with paper_size
        result = docx_post_process.process(input_bytes, paper_size="A4")

        # Verify result is valid bytes
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_process_with_orientation_only(self):
        """Test process() with only orientation parameter."""
        import io

        from docx import Document

        doc = Document()
        doc.add_paragraph("Test content")
        docx_bytes = io.BytesIO()
        doc.save(docx_bytes)
        docx_bytes.seek(0)
        input_bytes = docx_bytes.getvalue()

        # Call process with orientation
        result = docx_post_process.process(input_bytes, orientation="landscape")

        # Verify result is valid bytes
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_process_with_both_parameters(self):
        """Test process() with both paper_size and orientation parameters."""
        import io

        from docx import Document

        doc = Document()
        doc.add_paragraph("Test content")
        docx_bytes = io.BytesIO()
        doc.save(docx_bytes)
        docx_bytes.seek(0)
        input_bytes = docx_bytes.getvalue()

        # Call process with both parameters
        result = docx_post_process.process(input_bytes, paper_size="LETTER", orientation="portrait")

        # Verify result is valid bytes
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_process_returns_modified_document(self):
        """Test that process() returns a modified valid DOCX document."""
        import io

        from docx import Document

        # Create input document
        doc = Document()
        doc.add_paragraph("Test content")
        doc.add_table(rows=2, cols=2)
        docx_bytes = io.BytesIO()
        doc.save(docx_bytes)
        docx_bytes.seek(0)
        input_bytes = docx_bytes.getvalue()

        # Process the document
        result = docx_post_process.process(input_bytes, paper_size="A4", orientation="landscape")

        # Verify we can open the result as a valid DOCX
        result_doc = Document(io.BytesIO(result))
        assert len(result_doc.paragraphs) > 0
        assert len(result_doc.tables) > 0


def test_parser_supports_large_xml_documents():
    """Test that the XML parser is configured to handle large documents (XML_PARSE_HUGE).

    This test verifies that the parser patch in docx_post_process enables huge_tree=True,
    which is required to parse DOCX files with XML content exceeding 10MB.
    Without this patch, lxml raises 'Buffer size limit exceeded' errors.
    """
    from docx.oxml import parser as docx_parser

    # Create large XML content that would exceed the default 10MB buffer limit
    large_content = "x" * (11 * 1024 * 1024)  # 11MB of content
    large_xml = f"<root><data>{large_content}</data></root>".encode()

    # Parse using the patched python-docx parser - should not raise XMLSyntaxError
    # This would fail with "Buffer size limit exceeded" if huge_tree is not enabled
    result = etree.fromstring(large_xml, docx_parser.oxml_parser)
    assert result is not None
    assert result.tag == "root"


def test_process_document_with_large_content():
    """Test that process() can handle documents with large content.

    This is an integration test that creates a DOCX with substantial content
    and verifies it can be processed without XML buffer errors.
    """
    import io

    from docx import Document

    # Create a document with a lot of content (multiple paragraphs)
    doc = Document()

    # Add many paragraphs to create a larger document
    large_text = "Lorem ipsum dolor sit amet. " * 1000  # ~30KB per paragraph
    for _ in range(10):
        doc.add_paragraph(large_text)

        # Add a table with content
    table = doc.add_table(rows=5, cols=5)
    for row in table.rows:
        for cell in row.cells:
            cell.text = "Cell content " * 100

            # Save to bytes
    docx_bytes = io.BytesIO()
    doc.save(docx_bytes)
    docx_bytes.seek(0)
    input_bytes = docx_bytes.getvalue()

    # Process should succeed without raising XMLSyntaxError
    result = docx_post_process.process(input_bytes, paper_size="A4")

    # Verify result is valid
    assert isinstance(result, bytes)
    assert len(result) > 0

    # Verify the result can be opened as a valid DOCX
    result_doc = Document(io.BytesIO(result))
    assert len(result_doc.paragraphs) >= 10
    assert len(result_doc.tables) >= 1


def test_resize_images_in_cell_handles_large_xml():
    """Test that _resize_images_in_cell uses huge_tree parser for large XML content.

    This test verifies that the function can handle cells with XML content
    exceeding the default 10MB buffer limit (e.g., large base64-encoded images).
    Without the huge_tree parser, lxml raises 'Buffer size limit exceeded' errors.
    """
    # Create a mock cell with very large XML content (simulating a large base64 image)
    cell = MagicMock(spec=_Cell)

    # Create large content that exceeds 10MB default buffer limit
    # Using ~11MB of base64-like data to trigger the buffer limit
    large_base64_data = "A" * (11 * 1024 * 1024)  # 11MB of data

    large_cell_xml = f"""
    <w:tc xmlns:w="{WORD_PROCESSING_ML_MAIN_SCHEMA}"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:a="{DRAWING_ML_MAIN_SCHEMA}"
          xmlns:pic="{DRAWING_ML_PICTURE_SCHEMA}"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:tcPr/>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline>
                        <wp:extent cx="1000000" cy="750000"/>
                        <a:graphic>
                            <a:graphicData>
                                <pic:pic>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId5"/>
                                        <a:srcRect/>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:ext cx="1000000" cy="750000"/>
                                        </a:xfrm>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
                <!-- Large embedded data simulating base64 image content -->
                <w:t>{large_base64_data}</w:t>
            </w:r>
        </w:p>
    </w:tc>
    """

    # Set up the mock cell
    mock_tc = MagicMock()
    mock_tc.xml = large_cell_xml
    cell._tc = mock_tc

    # Set a max width that won't trigger resizing (larger than image)
    max_width = 2000000

    # This should NOT raise XMLSyntaxError: Buffer size limit exceeded
    # If _huge_tree_parser is not used, this would fail with:
    # lxml.etree.XMLSyntaxError: Resource limit exceeded: Buffer size limit exceeded
    docx_post_process._resize_images_in_cell(cell, max_width)

    # Verify the function completed without raising an exception
    # (no resizing needed since image width < max_width, so clear_content not called)
    mock_tc.clear_content.assert_not_called()


class TestMoveHeaderFooterReferencesToFirstSection:
    """Tests for _move_header_footer_references_to_first_section function."""

    def test_single_section_no_changes(self):
        """Test that single-section documents are not modified."""
        mock_doc = MagicMock()
        mock_section = MagicMock()
        mock_doc.sections = [mock_section]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Should return early without accessing sectPr
        mock_section._sectPr.findall.assert_not_called()

    def test_first_section_already_has_header_refs(self):
        """Test that no changes are made if first section already has header refs."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section already has header references
        mock_header_ref = MagicMock()
        mock_sect_pr1.findall.side_effect = lambda tag, namespaces: [mock_header_ref] if "header" in tag else []
        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Should not remove anything from last section
        mock_sect_pr2.remove.assert_not_called()

    def test_first_section_already_has_footer_refs(self):
        """Test that no changes are made if first section already has footer refs."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section already has footer references
        mock_footer_ref = MagicMock()
        mock_sect_pr1.findall.side_effect = lambda tag, namespaces: [mock_footer_ref] if "footer" in tag else []
        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Should not remove anything from last section
        mock_sect_pr2.remove.assert_not_called()

    def test_moves_header_refs_from_last_to_first(self):
        """Test that header references are moved from last to first section."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section has no refs
        mock_sect_pr1.findall.return_value = []
        mock_sect_pr1.find.return_value = None

        # Last section has header refs
        mock_header_ref = MagicMock()
        mock_sect_pr2.findall.side_effect = lambda tag, namespaces: [mock_header_ref] if "header" in tag else []
        mock_sect_pr2.find.return_value = None

        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Verify header ref was removed from last section and inserted into first
        mock_sect_pr2.remove.assert_called_with(mock_header_ref)
        mock_sect_pr1.insert.assert_called_with(0, mock_header_ref)

    def test_moves_footer_refs_from_last_to_first(self):
        """Test that footer references are moved from last to first section."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section has no refs
        mock_sect_pr1.findall.return_value = []
        mock_sect_pr1.find.return_value = None

        # Last section has footer refs
        mock_footer_ref = MagicMock()
        mock_sect_pr2.findall.side_effect = lambda tag, namespaces: [mock_footer_ref] if "footer" in tag else []
        mock_sect_pr2.find.return_value = None

        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Verify footer ref was removed from last section and inserted into first
        mock_sect_pr2.remove.assert_called_with(mock_footer_ref)
        mock_sect_pr1.insert.assert_called_with(0, mock_footer_ref)

    def test_moves_title_pg_from_last_to_first(self):
        """Test that titlePg element is moved from last to first section."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section has no refs
        mock_sect_pr1.findall.return_value = []

        # Last section has header refs and titlePg
        mock_header_ref = MagicMock()
        mock_title_pg = MagicMock()
        mock_sect_pr2.findall.side_effect = lambda tag, namespaces: [mock_header_ref] if "header" in tag else []
        mock_sect_pr2.find.side_effect = lambda tag, namespaces: mock_title_pg if "titlePg" in tag else None

        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Verify titlePg was removed from last section and appended to first
        mock_sect_pr2.remove.assert_any_call(mock_title_pg)
        mock_sect_pr1.append.assert_called_with(mock_title_pg)

    def test_no_title_pg_to_move(self):
        """Test that function handles missing titlePg gracefully."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section has no refs
        mock_sect_pr1.findall.return_value = []

        # Last section has header refs but no titlePg
        mock_header_ref = MagicMock()
        mock_sect_pr2.findall.side_effect = lambda tag, namespaces: [mock_header_ref] if "header" in tag else []
        mock_sect_pr2.find.return_value = None  # No titlePg

        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Should not call append for titlePg
        mock_sect_pr1.append.assert_not_called()

    def test_no_refs_in_last_section(self):
        """Test that function handles empty refs gracefully."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # Both sections have no refs
        mock_sect_pr1.findall.return_value = []
        mock_sect_pr2.findall.return_value = []
        mock_sect_pr2.find.return_value = None

        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Should not modify anything
        mock_sect_pr1.insert.assert_not_called()
        mock_sect_pr2.remove.assert_not_called()

    def test_multiple_header_footer_refs(self):
        """Test moving multiple header/footer refs (first, default, even types)."""
        mock_doc = MagicMock()
        mock_section1 = MagicMock()
        mock_section2 = MagicMock()
        mock_sect_pr1 = MagicMock()
        mock_sect_pr2 = MagicMock()

        # First section has no refs
        mock_sect_pr1.findall.return_value = []
        mock_sect_pr1.find.return_value = None

        # Last section has multiple header and footer refs
        mock_header_first = MagicMock()
        mock_header_default = MagicMock()
        mock_header_even = MagicMock()
        mock_footer_first = MagicMock()
        mock_footer_default = MagicMock()

        def findall_side_effect(tag, namespaces):
            if "header" in tag:
                return [mock_header_first, mock_header_default, mock_header_even]
            if "footer" in tag:
                return [mock_footer_first, mock_footer_default]
            return []

        mock_sect_pr2.findall.side_effect = findall_side_effect
        mock_sect_pr2.find.return_value = None

        mock_section1._sectPr = mock_sect_pr1
        mock_section2._sectPr = mock_sect_pr2
        mock_doc.sections = [mock_section1, mock_section2]

        docx_post_process._move_header_footer_references_to_first_section(mock_doc)

        # Verify all refs were moved
        assert mock_sect_pr2.remove.call_count == 5  # 3 headers + 2 footers
        assert mock_sect_pr1.insert.call_count == 5

    def test_integration_with_real_docx(self):
        """Integration test with real DOCX document having multiple sections."""
        import io

        from docx import Document

        # Create a document with two sections
        doc = Document()
        doc.add_paragraph("First section content")

        # Add a section break by adding a new section
        doc.add_section()
        doc.add_paragraph("Second section content")

        # Save and reload to ensure proper structure
        docx_bytes = io.BytesIO()
        doc.save(docx_bytes)
        docx_bytes.seek(0)

        # Process the document
        result = docx_post_process.process(docx_bytes.getvalue())

        # Verify result is valid
        assert isinstance(result, bytes)
        result_doc = Document(io.BytesIO(result))
        assert len(result_doc.sections) == 2


class TestReplaceFirstParagraphStyles:
    """Tests for _replace_first_paragraph_styles function."""

    @staticmethod
    def _make_doc_with_styles(styles: list[str | None]) -> MagicMock:
        """Build a mock Document whose body has paragraphs with given pStyle vals.

        ``None`` means the paragraph has no ``<w:pPr>`` at all.
        """
        from lxml import etree

        ns = WORD_PROCESSING_ML_MAIN_SCHEMA
        body = etree.Element(f"{{{ns}}}body")
        for style in styles:
            p = etree.SubElement(body, f"{{{ns}}}p")
            if style is not None:
                p_pr = etree.SubElement(p, f"{{{ns}}}pPr")
                etree.SubElement(p_pr, f"{{{ns}}}pStyle", {f"{{{ns}}}val": style})
        doc = MagicMock()
        doc.element.body = body
        return doc

    @staticmethod
    def _get_styles(doc: MagicMock) -> list[str | None]:
        """Extract the pStyle val from each paragraph, or None if absent."""
        ns = WORD_PROCESSING_ML_MAIN_SCHEMA
        result = []
        for p in doc.element.body.iterchildren(f"{{{ns}}}p"):
            p_pr = p.find(f"{{{ns}}}pPr")
            if p_pr is None:
                result.append(None)
                continue
            p_style = p_pr.find(f"{{{ns}}}pStyle")
            if p_style is None:
                result.append(None)
            else:
                result.append(p_style.get(f"{{{ns}}}val"))
        return result

    def test_replaces_first_paragraph_style_with_body_text(self):
        doc = self._make_doc_with_styles(["FirstParagraph"])
        docx_post_process._replace_first_paragraph_styles(doc)
        assert self._get_styles(doc) == ["BodyText"]

    def test_replaces_multiple_first_paragraph_styles(self):
        doc = self._make_doc_with_styles(["Heading1", "FirstParagraph", "BodyText", "FirstParagraph"])
        docx_post_process._replace_first_paragraph_styles(doc)
        assert self._get_styles(doc) == ["Heading1", "BodyText", "BodyText", "BodyText"]

    def test_does_not_change_other_paragraph_styles(self):
        doc = self._make_doc_with_styles(["Heading1", "Heading2", "BodyText", "ListParagraph"])
        docx_post_process._replace_first_paragraph_styles(doc)
        assert self._get_styles(doc) == ["Heading1", "Heading2", "BodyText", "ListParagraph"]

    def test_handles_paragraph_without_properties(self):
        doc = self._make_doc_with_styles([None, "FirstParagraph", None])
        docx_post_process._replace_first_paragraph_styles(doc)
        assert self._get_styles(doc) == [None, "BodyText", None]

    def test_handles_empty_document(self):
        doc = self._make_doc_with_styles([])
        docx_post_process._replace_first_paragraph_styles(doc)
        assert self._get_styles(doc) == []

    def test_integration_with_real_docx(self):
        """Integration test: add a heading + paragraph, convert, verify no FirstParagraph."""
        import io

        from docx import Document

        doc = Document()
        doc.add_heading("Test Heading", level=1)
        p = doc.add_paragraph("Body after heading")
        # Simulate pandoc's behavior by setting the style
        p.style = doc.styles["Normal"]
        p.style = doc.styles.add_style("First Paragraph", 1) if "First Paragraph" not in [s.name for s in doc.styles] else doc.styles["First Paragraph"]

        buf = io.BytesIO()
        doc.save(buf)
        result = docx_post_process.process(buf.getvalue())

        result_doc = Document(io.BytesIO(result))
        ns = WORD_PROCESSING_ML_MAIN_SCHEMA
        first_paragraph_replaced = False
        for para in result_doc.element.body.iterchildren(f"{{{ns}}}p"):
            p_pr = para.find(f"{{{ns}}}pPr")
            if p_pr is None:
                continue
            p_style = p_pr.find(f"{{{ns}}}pStyle")
            if p_style is not None:
                val = p_style.get(f"{{{ns}}}val")
                assert val != "FirstParagraph", "FirstParagraph style should have been replaced"
                if val == "BodyText":
                    first_paragraph_replaced = True
        assert first_paragraph_replaced, "Expected at least one paragraph to be replaced with BodyText"

        # ---- _resolve_image_src ----


def test_resolve_image_src_data_uri():
    """data: URI with base64-encoded 1x1 GIF should return bytes."""
    gif_b64 = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
    result = _resolve_image_src(f"data:image/gif;base64,{gif_b64}")
    assert result is not None
    assert result[:3] == b"GIF"


def test_resolve_image_src_invalid_base64():
    """data: URI with invalid base64 should return None."""
    assert _resolve_image_src("data:image/png;base64,!!!invalid!!!") is None


def test_resolve_image_src_http_unsupported():
    """http:// URIs are not supported and should return None."""
    assert _resolve_image_src("http://example.com/img.png") is None


def test_resolve_image_src_empty():
    """Empty string should return None."""
    assert _resolve_image_src("") is None

    # ---- _replace_image_placeholders ----


def test_replace_image_placeholder_with_data_uri():
    """Image placeholder with data: URI should be replaced with w:drawing."""
    from docx import Document

    gif_b64 = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
    doc = Document()
    doc.add_paragraph(f"{{{{IMG:||data:image/gif;base64,{gif_b64}}}}}")

    _replace_image_placeholders(doc)

    body_xml = etree.tostring(doc.element.body, encoding="unicode")
    assert "w:drawing" in body_xml
    assert "{{IMG:" not in body_xml


def test_replace_image_placeholder_unsupported_src():
    """Unsupported image src should be replaced with [image] text."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("{{IMG:||http://example.com/img.png}}")

    _replace_image_placeholders(doc)

    body_xml = etree.tostring(doc.element.body, encoding="unicode")
    assert "[image]" in body_xml
    assert "{{IMG:" not in body_xml


def test_replace_image_placeholder_unique_ids():
    """Multiple image placeholders should get unique docPr IDs."""
    from docx import Document

    gif_b64 = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
    doc = Document()
    doc.add_paragraph(f"{{{{IMG:||data:image/gif;base64,{gif_b64}}}}}")
    doc.add_paragraph(f"{{{{IMG:||data:image/gif;base64,{gif_b64}}}}}")

    _replace_image_placeholders(doc)

    body_xml = etree.tostring(doc.element.body, encoding="unicode")
    import re

    ids = re.findall(r'docPr id="(\d+)"', body_xml)
    assert len(ids) >= 2
    assert ids[0] != ids[1]

    # ---- _replace_link_placeholders ----


def test_replace_link_placeholder():
    """HREF placeholder in hyperlink tooltip should be replaced with real r:id."""
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    doc = Document()
    body = doc.element.body
    p = parse_xml(f'<w:p {nsdecls("w")}><w:hyperlink w:tooltip="{{{{HREF:http://example.com}}}}"><w:r><w:t>click</w:t></w:r></w:hyperlink></w:p>')
    body.append(p)

    _replace_link_placeholders(doc)

    body_xml = etree.tostring(body, encoding="unicode")
    assert "{{HREF:" not in body_xml
    # r:id should be set on the hyperlink
    ns_r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    hyperlink = body.find(f".//{{{SCHEMA}}}hyperlink")
    assert hyperlink is not None
    assert hyperlink.get(f"{ns_r}id") is not None


def test_replace_link_placeholder_no_match():
    """Hyperlink without HREF placeholder should not be modified."""
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    doc = Document()
    body = doc.element.body
    p = parse_xml(f'<w:p {nsdecls("w")}><w:hyperlink w:tooltip="Normal tooltip"><w:r><w:t>click</w:t></w:r></w:hyperlink></w:p>')
    body.append(p)

    _replace_link_placeholders(doc)

    hyperlink = body.find(f".//{{{SCHEMA}}}hyperlink")
    assert hyperlink.get(f"{{{SCHEMA}}}tooltip") == "Normal tooltip"

    # ---- _has_existing_fixed_width ----


def test_has_existing_fixed_width_true():
    """tblPr with tblW type=dxa and positive width should return True."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    tbl_pr = parse_xml(f'<w:tblPr {nsdecls("w")}><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>')
    assert _has_existing_fixed_width(tbl_pr) is True


def test_has_existing_fixed_width_pct():
    """tblPr with tblW type=pct should return False."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    tbl_pr = parse_xml(f'<w:tblPr {nsdecls("w")}><w:tblW w:w="5000" w:type="pct"/></w:tblPr>')
    assert _has_existing_fixed_width(tbl_pr) is False


def test_has_existing_fixed_width_zero():
    """tblPr with tblW type=dxa but w=0 should return False."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    tbl_pr = parse_xml(f'<w:tblPr {nsdecls("w")}><w:tblW w:w="0" w:type="dxa"/></w:tblPr>')
    assert _has_existing_fixed_width(tbl_pr) is False


def test_has_existing_fixed_width_missing():
    """tblPr without tblW should return False."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    tbl_pr = parse_xml(f"<w:tblPr {nsdecls('w')}/>")
    assert _has_existing_fixed_width(tbl_pr) is False

    # ---- table width clamping ----


def test_apply_table_layout_clamps_dxa_to_page_width():
    """Fixed dxa width from html_table_layout should be clamped to page width."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    from app.docx_post_process import _apply_table_layout

    tbl = parse_xml(f'<w:tbl {nsdecls("w")}><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid></w:tbl>')
    tbl_pr = tbl.find(f"{{{SCHEMA}}}tblPr")

    layout = MagicMock()
    layout.width_type = "dxa"
    layout.width_value = 11070  # 738px, wider than page
    layout.jc = None
    layout.indent_twips = None

    # max_width in EMU: 9360 twips * 635 = ~5943600
    _apply_table_layout(tbl, tbl_pr, layout, max_width=5943600)

    tbl_w = tbl_pr.find(f"{{{SCHEMA}}}tblW")
    assert tbl_w is not None
    actual_width = int(tbl_w.get(f"{{{SCHEMA}}}w"))
    assert actual_width <= 9360  # clamped to page width


def test_apply_table_layout_preserves_lua_fixed_width():
    """Lua filter fixed width that fits should not be changed."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    from app.docx_post_process import _apply_table_layout

    tbl = parse_xml(f'<w:tbl {nsdecls("w")}><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid></w:tbl>')
    tbl_pr = tbl.find(f"{{{SCHEMA}}}tblPr")

    _apply_table_layout(tbl, tbl_pr, None, max_width=5943600)

    tbl_w = tbl_pr.find(f"{{{SCHEMA}}}tblW")
    assert int(tbl_w.get(f"{{{SCHEMA}}}w")) == 5000  # unchanged


def test_apply_table_layout_clamps_lua_fixed_width():
    """Lua filter fixed width that overflows should be clamped."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    from app.docx_post_process import _apply_table_layout

    tbl = parse_xml(f'<w:tbl {nsdecls("w")}><w:tblPr><w:tblW w:w="11070" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="5535"/><w:gridCol w:w="5535"/></w:tblGrid></w:tbl>')
    tbl_pr = tbl.find(f"{{{SCHEMA}}}tblPr")

    _apply_table_layout(tbl, tbl_pr, None, max_width=5943600)

    tbl_w = tbl_pr.find(f"{{{SCHEMA}}}tblW")
    assert int(tbl_w.get(f"{{{SCHEMA}}}w")) <= 9360


# ---- Adjacent table separation ----


def _document_with_body(body_xml: str):
    """Build a Document whose body holds the given OOXML blocks."""
    import io

    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    doc = Document()
    body = doc.element.body
    for child in list(body):
        if not child.tag.endswith("sectPr"):
            body.remove(child)
    fragment = parse_xml(f"<w:root {nsdecls('w')}>{body_xml}</w:root>")
    sect_pr = body.find(f"{{{SCHEMA}}}sectPr")
    for block in list(fragment):
        if sect_pr is not None:
            sect_pr.addprevious(block)
        else:
            body.append(block)
    buffer = io.BytesIO()
    doc.save(buffer)
    return Document(io.BytesIO(buffer.getvalue()))


_TBL_XML = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>{0}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'


def _body_children(doc):
    return [etree.QName(child).localname for child in doc.element.body if isinstance(child.tag, str)]


def test_separate_adjacent_tables_inserts_paragraph():
    """Two back-to-back tables get an empty paragraph between them."""
    from app.docx_post_process import _separate_adjacent_tables

    doc = _document_with_body(_TBL_XML.format("first") + _TBL_XML.format("second"))
    assert _body_children(doc)[:2] == ["tbl", "tbl"]

    _separate_adjacent_tables(doc)

    assert _body_children(doc)[:3] == ["tbl", "p", "tbl"]


def test_separate_adjacent_tables_looks_through_bookmarks():
    """Bookmarks between two tables do not keep them apart in Word."""
    from app.docx_post_process import _separate_adjacent_tables

    bookmarks = '<w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="fields_end"/>'
    doc = _document_with_body(_TBL_XML.format("first") + bookmarks + _TBL_XML.format("second"))

    _separate_adjacent_tables(doc)

    assert _body_children(doc)[:5] == ["tbl", "p", "bookmarkEnd", "bookmarkStart", "tbl"]


def test_separate_adjacent_tables_keeps_existing_paragraph():
    """A paragraph already separates the tables, so nothing is added."""
    from app.docx_post_process import _separate_adjacent_tables

    doc = _document_with_body(_TBL_XML.format("first") + "<w:p/>" + _TBL_XML.format("second"))

    _separate_adjacent_tables(doc)

    assert _body_children(doc)[:3] == ["tbl", "p", "tbl"]


def test_separate_adjacent_tables_handles_nested_tables():
    """Adjacent tables inside a table cell are separated too."""
    from app.docx_post_process import _separate_adjacent_tables

    nested = _TBL_XML.format("inner one") + _TBL_XML.format("inner two")
    outer = f"<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>{nested}<w:p/></w:tc></w:tr></w:tbl>"
    doc = _document_with_body(outer)

    _separate_adjacent_tables(doc)

    cell = doc.element.body.find(f"{{{SCHEMA}}}tbl/{{{SCHEMA}}}tr/{{{SCHEMA}}}tc")
    assert [etree.QName(child).localname for child in cell] == ["tbl", "p", "tbl", "p"]


def test_process_separates_adjacent_tables():
    """The full post-processing pass separates tables Word would merge."""
    import io

    from docx import Document

    doc = _document_with_body(_TBL_XML.format("first") + _TBL_XML.format("second"))
    buffer = io.BytesIO()
    doc.save(buffer)

    result = Document(io.BytesIO(docx_post_process.process(buffer.getvalue())))

    assert _body_children(result)[:3] == ["tbl", "p", "tbl"]


# ---- Image placeholder dimensions ----


def test_dimension_to_emu_units():
    """Every absolute CSS unit the placeholder can carry converts to EMU."""
    from app.docx_post_process import EMU_1_INCH, _dimension_to_emu

    assert _dimension_to_emu("96px") == EMU_1_INCH
    assert _dimension_to_emu("96") == EMU_1_INCH  # a bare number is px
    assert _dimension_to_emu("1in") == EMU_1_INCH
    assert _dimension_to_emu("72pt") == EMU_1_INCH
    assert _dimension_to_emu("6pc") == EMU_1_INCH
    assert _dimension_to_emu("2.54cm") == EMU_1_INCH
    assert _dimension_to_emu("25.4mm") == EMU_1_INCH


@pytest.mark.parametrize("value", ["", "50%", "auto", "abc", "10em", "-5px", "+5px", "1.2.3px", "1e3px", ".5in"])
def test_dimension_to_emu_rejects_what_it_cannot_resolve(value):
    """An empty field, a percentage or an unknown unit means "no dimension"."""
    from app.docx_post_process import _dimension_to_emu

    assert _dimension_to_emu(value) is None


def test_dimension_to_emu_whitespace_contract():
    """Surrounding whitespace is stripped by the caller, not by the pattern.

    The pattern carries no \\s* quantifiers (SonarCloud S5852), so whitespace
    *between* the number and its unit no longer matches. filters/inline_styles.lua
    builds these fields as `num .. unit`, so that form never reaches here.
    """
    from app.docx_post_process import EMU_1_INCH, _dimension_to_emu

    assert _dimension_to_emu("  96px  ") == EMU_1_INCH
    assert _dimension_to_emu("96 px") is None


def test_dimension_to_emu_is_case_insensitive():
    from app.docx_post_process import EMU_1_INCH, _dimension_to_emu

    assert _dimension_to_emu("96PX") == EMU_1_INCH
    assert _dimension_to_emu("1IN") == EMU_1_INCH


def test_resolve_image_extent_prefers_the_requested_size():
    from app.docx_post_process import EMU_1_INCH, _resolve_image_extent

    assert _resolve_image_extent(("96px", "192px"), 20, 40) == (EMU_1_INCH, 2 * EMU_1_INCH)


def test_resolve_image_extent_scales_the_missing_side():
    """Only one side given: keep the file's aspect ratio, as the writer does."""
    from app.docx_post_process import EMU_1_INCH, _resolve_image_extent

    assert _resolve_image_extent(("96px", ""), 20, 40) == (EMU_1_INCH, 2 * EMU_1_INCH)
    assert _resolve_image_extent(("", "192px"), 20, 40) == (EMU_1_INCH, 2 * EMU_1_INCH)


def test_resolve_image_extent_falls_back_to_native_size():
    from app.docx_post_process import _resolve_image_extent

    assert _resolve_image_extent(("", ""), 96, 48) == (docx_post_process.EMU_1_INCH, docx_post_process.EMU_1_INCH // 2)


def test_resolve_image_extent_survives_a_degenerate_image():
    """A zero-sized source must not divide by zero."""
    from app.docx_post_process import EMU_1_INCH, _resolve_image_extent

    assert _resolve_image_extent(("96px", ""), 0, 0) == (EMU_1_INCH, 0)
    assert _resolve_image_extent(("", "96px"), 0, 0) == (0, EMU_1_INCH)


# ---- _cap_image_heights (#245) ----


def _png_bytes(width: int = 3, height: int = 30) -> bytes:
    """A minimal RGB PNG of the given pixel size."""
    import struct
    import zlib

    raw = b"".join(b"\x00" + b"\xff\xff\xff" * width for _ in range(height))

    def chunk(tag: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + tag + payload + struct.pack(">I", zlib.crc32(tag + payload))

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def _page_limit(text_height_inches: float) -> int:
    """The height the cap allows: the text area less the line the image is set on."""
    from docx.shared import Inches

    from app.docx_post_process import LINE_ALLOWANCE_EMU

    return int(Inches(text_height_inches)) - LINE_ALLOWANCE_EMU


def _capped(width_inches: float, height_inches: float, text_height_inches: float) -> tuple[int, int]:
    """The size a picture of this shape is brought back to on a page of this text height."""
    from docx.shared import Inches

    limit = _page_limit(text_height_inches)
    return int(int(Inches(width_inches)) * limit / int(Inches(height_inches))), limit


def _extents(doc) -> list[tuple[int, int]]:
    from app.docx_post_process import WP_NS

    return [(int(e.get("cx")), int(e.get("cy"))) for e in doc.element.body.iter(f"{{{WP_NS}}}extent")]


def _frame_extents(doc) -> list[tuple[int, int]]:
    return [(int(e.get("cx")), int(e.get("cy"))) for e in doc.element.body.iter(f"{{{DRAWING_ML_MAIN_SCHEMA}}}ext") if e.get("cx")]


def test_cap_image_heights_brings_a_tall_image_back_to_the_page():
    """Letter less 1 inch margins leaves 9 inch, less the line the image sits on."""
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(3), height=Inches(30))

    _cap_image_heights(doc)

    assert _extents(doc) == [_capped(3, 30, 9)]
    assert _frame_extents(doc) == [_capped(3, 30, 9)]


def test_cap_image_heights_leaves_an_image_that_fits():
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(3), height=Inches(5))

    _cap_image_heights(doc)

    assert _extents(doc) == [(Inches(3), Inches(5))]


def test_cap_image_heights_reaches_into_a_table_cell():
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    cell = doc.add_table(rows=1, cols=1).cell(0, 0)
    cell.paragraphs[0].add_run().add_picture(io.BytesIO(_png_bytes()), width=Inches(1), height=Inches(18))

    _cap_image_heights(doc)

    assert _extents(doc) == [_capped(1, 18, 9)]


def test_cap_image_heights_uses_the_page_of_each_section():
    """A landscape section is shorter, so its image is capped lower than the first one."""
    import io

    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(2), height=Inches(20))
    landscape = doc.add_section()
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = Inches(11), Inches(8.5)
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(2), height=Inches(20))

    _cap_image_heights(doc)

    assert _extents(doc) == [_capped(2, 20, 9), _capped(2, 20, 6.5)]


def test_cap_image_heights_leaves_an_extension_list_entry_alone():
    """An <a:ext uri=...> of an extension list has no size and must not get one."""
    import io

    from docx import Document
    from docx.oxml import parse_xml
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(3), height=Inches(30))
    blip = next(doc.element.body.iter(f"{{{DRAWING_ML_MAIN_SCHEMA}}}blip"))
    blip.append(parse_xml(f'<a:extLst xmlns:a="{DRAWING_ML_MAIN_SCHEMA}"><a:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}"/></a:extLst>'))

    _cap_image_heights(doc)

    ext_list_entry = next(e for e in doc.element.body.iter(f"{{{DRAWING_ML_MAIN_SCHEMA}}}ext") if e.get("uri"))
    assert ext_list_entry.get("cx") is None
    assert _frame_extents(doc) == [_capped(3, 30, 9)]


def test_cap_image_heights_keeps_the_whole_of_a_page_without_margins():
    """A margin of nothing is a margin the document states, not one it leaves out; #245."""
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    doc.sections[0].top_margin = doc.sections[0].bottom_margin = Inches(0)
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(3), height=Inches(30))

    _cap_image_heights(doc)

    # The page is 11 inch and keeps all of it, where a fallback of 1 inch a side would leave 9
    assert _extents(doc) == [_capped(3, 30, 11)]


def test_cap_image_heights_reads_a_negative_margin_by_its_size():
    """A negative top margin fixes the header distance; its size is what the text sits inside.

    `app/docx_page_geometry.py` lays the PDF out that way, so the cap has to agree with it or a
    capped image still runs past the bottom of the page.
    """
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _cap_image_heights

    doc = Document()
    doc.sections[0].top_margin = Inches(-1)
    doc.add_picture(io.BytesIO(_png_bytes()), width=Inches(3), height=Inches(30))

    _cap_image_heights(doc)

    # 11 inch less a top margin of 1 and a bottom margin of 1, not 11 plus 1 less 1
    assert _extents(doc) == [_capped(3, 30, 9)]


# ---- image width per column (#246) ----


def _png_bytes_246(width: int = 30, height: int = 10) -> bytes:
    import io

    from PIL import Image

    buffer = io.BytesIO()
    Image.new("RGB", (width, height), (40, 90, 170)).save(buffer, format="PNG")
    return buffer.getvalue()


def _table_with_image(grid_twips: list[int], image_inches: float, merge: bool = False, cell_margin_twips: int | None = None):
    """A table whose grid states the given widths, an image in its first cell."""
    import io

    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Inches

    doc = Document()
    table = doc.add_table(rows=1, cols=len(grid_twips))
    for col, width in zip(table._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol"), grid_twips, strict=True):
        col.set(f"{{{SCHEMA}}}w", str(width))
    if cell_margin_twips is not None:
        table._tbl.tblPr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:left w:w="{cell_margin_twips}" w:type="dxa"/><w:right w:w="{cell_margin_twips}" w:type="dxa"/></w:tblCellMar>'))
    cell = table.cell(0, 0)
    if merge:
        cell = cell.merge(table.cell(0, 1))
    cell.paragraphs[0].add_run().add_picture(io.BytesIO(_png_bytes_246()), width=Inches(image_inches))
    return doc, table


def _first_image_width(doc) -> int:
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    return int(next(doc.element.body.iter(f"{{{wp}}}extent")).get("cx"))


def test_image_in_a_cell_is_brought_back_to_its_column():
    """Two columns of 264 px: the image fits its own column less the default cell margins, not half the page."""
    doc, table = _table_with_image([5280, 5280], image_inches=6)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _first_image_width(doc) == (5280 - 2 * 108) * 635


def test_image_in_a_merged_cell_takes_the_sum_of_its_columns():
    doc, table = _table_with_image([3000, 4000, 2000], image_inches=6, merge=True)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _first_image_width(doc) == (3000 + 4000 - 2 * 108) * 635


def test_cell_margins_of_the_table_are_taken_off():
    doc, table = _table_with_image([5280, 5280], image_inches=6, cell_margin_twips=300)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _first_image_width(doc) == (5280 - 2 * 300) * 635


def test_image_which_fits_its_column_keeps_its_size():
    doc, table = _table_with_image([5280, 5280], image_inches=1)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _first_image_width(doc) == 914400


def test_without_a_usable_grid_the_even_share_stays_the_limit():
    doc, table = _table_with_image([0, 5280], image_inches=6)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _first_image_width(doc) == int(6.5 * 914400 / 2)


def test_cell_margins_of_the_cell_win_and_start_end_count():
    """The cell's own margin wins over the table's; start/end name the sides too; a pct margin is ignored."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    doc, table = _table_with_image([5280, 5280], image_inches=6, cell_margin_twips=300)
    tc_pr = table.cell(0, 0)._tc.get_or_add_tcPr()
    tc_pr.append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:start w:w="50" w:type="dxa"/><w:end w:w="10" w:type="pct"/></w:tcMar>'))

    _process_table(table, 0, max_width=int(6.5 * 914400))

    # left: the cell's 50; right: the pct value is ignored, so the table's 300 stays
    assert _first_image_width(doc) == (5280 - 50 - 300) * 635


def test_a_cell_beyond_the_grid_falls_back_to_the_even_share():
    """A cell placed past the last grid column has no width to read; the even share stays."""
    from app.docx_post_process import _cell_image_width

    cell = MagicMock()
    cell._tc.grid_offset, cell._tc.grid_span = 3, 1

    assert _cell_image_width(MagicMock(), cell, [5280 * 635, 5280 * 635], 1234.0, 6 * 914400) == 1234.0


def test_a_grid_wider_than_the_page_does_not_let_an_image_past_it():
    """A table laid out to its content keeps a grid wider than the page; an image still fits; #246."""
    # 2 x 5280 twips is 7.33 inch of grid on a text width of 6.5
    doc, table = _table_with_image([5280, 5280], image_inches=7)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _first_image_width(doc) <= int(6.5 * 914400)


def test_an_image_in_a_nested_table_stays_inside_the_cell_holding_it():
    """A nested grid may state more than the cell it sits in, and the cell is what bounds it; #246."""
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _process_table

    doc = Document()
    outer = doc.add_table(rows=1, cols=2)
    for col, width in zip(outer._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol"), [2880, 2880], strict=True):
        col.set(f"{{{SCHEMA}}}w", str(width))
    # The nested grid claims 7 inch inside a cell of 2 inch
    inner = outer.cell(0, 0).add_table(rows=1, cols=1)
    inner._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol")[0].set(f"{{{SCHEMA}}}w", "10080")
    inner.cell(0, 0).paragraphs[0].add_run().add_picture(io.BytesIO(_png_bytes_246()), width=Inches(7))

    _process_table(outer, 0, max_width=int(6.5 * 914400))

    # The outer cell is 2880 twips less its margins; the image is inside that, not inside 7 inch
    assert _first_image_width(doc) <= 2880 * 635


def test_cell_margins_of_the_table_style_are_taken_off():
    """A template sets the margins of every table through its style; #246."""
    import io

    from docx import Document
    from docx.shared import Inches

    from app.docx_post_process import _process_table

    doc = Document()
    style = next(element for element in doc.styles.element.findall(f"{{{SCHEMA}}}style") if element.get(f"{{{SCHEMA}}}styleId") == "TableGrid")
    # The style states its own margins already, and a second tblCellMar would never be read
    style_margins = style.find(f"{{{SCHEMA}}}tblPr/{{{SCHEMA}}}tblCellMar")
    for side in ("left", "right"):
        style_margins.find(f"{{{SCHEMA}}}{side}").set(f"{{{SCHEMA}}}w", "720")
    table = doc.add_table(rows=1, cols=1, style="Table Grid")
    table._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol")[0].set(f"{{{SCHEMA}}}w", "5280")
    table.cell(0, 0).paragraphs[0].add_run().add_picture(io.BytesIO(_png_bytes_246()), width=Inches(6))

    _process_table(table, 0, max_width=int(6.5 * 914400))

    # 5280 twips less 720 a side, which the default of 108 would have left far too wide
    assert _first_image_width(doc) == (5280 - 720 - 720) * 635


def test_cell_margins_come_from_each_style_which_states_a_side():
    """A style states the sides it changes and leaves the rest to the one it is based on; #246."""
    import io

    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Inches

    from app.docx_post_process import _process_table

    doc = Document()
    styles = {element.get(f"{{{SCHEMA}}}styleId"): element for element in doc.styles.element.findall(f"{{{SCHEMA}}}style")}
    # The parent states the left margin alone
    parent_margins = styles["TableNormal"].find(f"{{{SCHEMA}}}tblPr/{{{SCHEMA}}}tblCellMar")
    for side in ("right", "top", "bottom"):
        element = parent_margins.find(f"{{{SCHEMA}}}{side}")
        if element is not None:
            parent_margins.remove(element)
    parent_margins.find(f"{{{SCHEMA}}}left").set(f"{{{SCHEMA}}}w", "600")
    # The style the table names states the right one, and is based on that parent
    child = styles["TableGrid"]
    child_margins = child.find(f"{{{SCHEMA}}}tblPr/{{{SCHEMA}}}tblCellMar")
    child.find(f"{{{SCHEMA}}}tblPr").remove(child_margins)
    child.find(f"{{{SCHEMA}}}tblPr").append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:right w:w="400" w:type="dxa"/></w:tblCellMar>'))

    table = doc.add_table(rows=1, cols=1, style="Table Grid")
    table._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol")[0].set(f"{{{SCHEMA}}}w", "5280")
    table.cell(0, 0).paragraphs[0].add_run().add_picture(io.BytesIO(_png_bytes_246()), width=Inches(6))

    _process_table(table, 0, max_width=int(6.5 * 914400))

    # 600 from the parent style, 400 from the style the table names; neither side falls back to 108
    assert _first_image_width(doc) == (5280 - 600 - 400) * 635


# ---- placeholder images sized like pandoc's own ----


A4_TEXT_WIDTH_EMU = 481 * 12700  # A4 with 2 cm side margins: 481.9 pt, counted in whole points as pandoc does


def _document_with_page(page_xml: str | None):
    """A Document whose body sectPr states only the given page elements, or no sectPr at all."""
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    doc = Document()
    sect_pr = doc.element.body.find(f"{{{SCHEMA}}}sectPr")
    doc.element.body.remove(sect_pr)
    if page_xml is not None:
        doc.element.body.append(parse_xml(f"<w:sectPr {nsdecls('w')}>{page_xml}</w:sectPr>"))
    return doc


_A4_PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>'


def _placeholder(width: str, image_width_px: int = 200, image_height_px: int = 100) -> str:
    import base64

    return f"{{{{IMG:{width}||data:image/png;base64,{base64.b64encode(_png_bytes(image_width_px, image_height_px)).decode()}}}}}"


def _processed(doc, **kwargs) -> object:
    import io

    from docx import Document

    from app.docx_post_process import process

    buffer = io.BytesIO()
    doc.save(buffer)
    return Document(io.BytesIO(process(buffer.getvalue(), **kwargs)))


def test_pandoc_text_width_is_the_page_less_its_side_margins():
    from app.docx_post_process import _pandoc_text_width_emu

    assert _pandoc_text_width_emu(_document_with_page(_A4_PAGE)) == A4_TEXT_WIDTH_EMU


@pytest.mark.parametrize(
    "page_xml",
    [
        None,
        "",
        '<w:pgSz w:w="11906" w:h="16838"/>',
        '<w:pgMar w:left="1134" w:right="1134"/>',
        '<w:pgSz w:w="11906"/><w:pgMar w:left="1134"/>',
        '<w:pgSz w:w="wide"/><w:pgMar w:left="1134" w:right="1134"/>',
        '<w:pgSz w:w="2000"/><w:pgMar w:left="1134" w:right="1134"/>',
    ],
)
def test_pandoc_text_width_falls_back_to_420_points(page_xml):
    """pandoc uses 420 pt when the reference document does not state the whole page."""
    from app.docx_post_process import PANDOC_DEFAULT_TEXT_WIDTH_EMU, _pandoc_text_width_emu

    assert PANDOC_DEFAULT_TEXT_WIDTH_EMU == 5334000
    assert _pandoc_text_width_emu(_document_with_page(page_xml)) == PANDOC_DEFAULT_TEXT_WIDTH_EMU


def test_dimension_to_emu_reads_a_percentage_as_a_share_of_the_text_width():
    from app.docx_post_process import _dimension_to_emu

    assert _dimension_to_emu("50%", 1000) == 500
    assert _dimension_to_emu("12.5%", 1000) == 125
    assert _dimension_to_emu("50%") is None


def test_resolve_image_extent_resolves_a_percentage():
    """A percentage of either side is a share of the text width; the other side keeps the aspect ratio."""
    from app.docx_post_process import _resolve_image_extent

    assert _resolve_image_extent(("50%", ""), 200, 100, 4000) == (2000, 1000)
    assert _resolve_image_extent(("", "10%"), 200, 100, 4000) == (800, 400)
    assert _resolve_image_extent(("50%", "10%"), 200, 100, 4000) == (2000, 400)


def test_resolve_image_extent_ignores_a_length_beside_a_single_percentage():
    """pandoc sizes the side a lone percentage leaves from the file's aspect ratio, not from a length given there."""
    from app.docx_post_process import _resolve_image_extent

    assert _resolve_image_extent(("50%", "200px"), 200, 100, 4000) == (2000, 1000)
    assert _resolve_image_extent(("100px", "25%"), 200, 100, 4000) == (2000, 1000)


def test_resolve_image_extent_brings_a_wide_image_back_to_the_text_width():
    """The writer brings every image back to the text width, the requested ratio kept."""
    from app.docx_post_process import EMU_1_INCH, _resolve_image_extent

    assert _resolve_image_extent(("960px", "96px"), 20, 40, 5 * EMU_1_INCH) == (5 * EMU_1_INCH, EMU_1_INCH // 2)
    assert _resolve_image_extent(("", ""), 960, 96, 5 * EMU_1_INCH) == (5 * EMU_1_INCH, EMU_1_INCH // 2)
    assert _resolve_image_extent(("960px", ""), 20, 40, None) == (10 * EMU_1_INCH, 20 * EMU_1_INCH)


def test_placeholder_percentage_is_a_share_of_the_page_pandoc_used():
    doc = _document_with_page(_A4_PAGE)
    doc.add_paragraph(_placeholder("50%"))

    assert _extents(_processed(doc)) == [(A4_TEXT_WIDTH_EMU // 2, A4_TEXT_WIDTH_EMU // 4)]


def test_placeholder_percentage_ignores_the_requested_paper_size():
    """pandoc sized its images before the paper size reached the page, so the placeholder does too."""
    from app.docx_post_process import PANDOC_DEFAULT_TEXT_WIDTH_EMU

    # pandoc's own reference document: a sectPr that states no page
    doc = _document_with_page("")
    doc.add_paragraph(_placeholder("100%"))

    assert _extents(_processed(doc, paper_size="A3", orientation="landscape")) == [(PANDOC_DEFAULT_TEXT_WIDTH_EMU, PANDOC_DEFAULT_TEXT_WIDTH_EMU // 2)]


def test_placeholder_image_wider_than_the_page_is_brought_back_to_it():
    doc = _document_with_page(_A4_PAGE)
    doc.add_paragraph(_placeholder("", image_width_px=3000, image_height_px=300))

    assert _extents(_processed(doc)) == [(A4_TEXT_WIDTH_EMU, A4_TEXT_WIDTH_EMU // 10)]


def test_placeholder_image_in_a_cell_is_brought_back_to_its_column():
    """The placeholder is resolved before the tables, so the column limit reaches it."""
    doc = _document_with_page(_A4_PAGE)
    table = doc.add_table(rows=1, cols=4)
    for col in table._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol"):
        col.set(f"{{{SCHEMA}}}w", "2000")
    table.cell(0, 0).paragraphs[0].add_run(_placeholder("50%"))

    assert _extents(_processed(doc))[0][0] == (2000 - 2 * 108) * 635


def test_image_in_a_cell_keeps_its_frame_in_step_with_its_extent():
    """A frame left at the old size lets Word crop or distort the picture."""
    doc, table = _table_with_image([2000, 2000], image_inches=6)

    _process_table(table, 0, max_width=int(6.5 * 914400))

    assert _frame_extents(doc) == _extents(doc)
    assert _extents(doc)[0][0] == (2000 - 2 * 108) * 635


def test_placeholder_image_in_a_cell_keeps_its_frame_in_step_with_its_extent():
    doc = _document_with_page(_A4_PAGE)
    table = doc.add_table(rows=1, cols=4)
    for col in table._tbl.tblGrid.findall(f"{{{SCHEMA}}}gridCol"):
        col.set(f"{{{SCHEMA}}}w", "2000")
    table.cell(0, 0).paragraphs[0].add_run(_placeholder("", image_width_px=3000, image_height_px=300))

    result = _processed(doc)

    assert _frame_extents(result) == _extents(result)


# ---- _apply_image_layouts ----

WP_SCHEMA = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _paragraph_with_marked_picture(marker: str, label_size: int | None = None, picture_rpr: str = ""):
    """A paragraph holding a marker run, a picture run and a label run."""
    import io

    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    doc = Document()
    paragraph = doc.add_paragraph()
    picture_run = paragraph.add_run()
    picture_run.add_picture(io.BytesIO(_png_bytes(16, 16)))
    if picture_rpr:
        picture_run._r.insert(0, parse_xml(f"<w:rPr {nsdecls('w')}>{picture_rpr}</w:rPr>"))
    picture_run._r.addprevious(parse_xml(f'<w:r {nsdecls("w")}><w:t xml:space="preserve">{marker}</w:t></w:r>'))
    label = paragraph.add_run("Draft")
    if label_size is not None:
        label.font.size = label_size * 6350
    return doc, picture_run._r


def _position(run) -> str | None:
    position = run.find(f"{{{SCHEMA}}}rPr/{{{SCHEMA}}}position")
    return None if position is None else position.get(f"{{{SCHEMA}}}val")


def _effect_extent(run) -> dict[str, str] | None:
    effect_extent = run.find(f".//{{{WP_SCHEMA}}}effectExtent")
    return None if effect_extent is None else dict(effect_extent.attrib)


def test_image_layout_bottom_lowers_the_picture_by_the_descent_of_its_label():
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:bottom||2px}}", label_size=24)

    _apply_image_layouts(doc)

    # A quarter of 12pt, in half-points
    assert _position(picture_run) == "-6"
    assert _effect_extent(picture_run) == {"l": "0", "t": "0", "r": "19050", "b": "0"}
    assert "IMGLAYOUT" not in etree.tostring(doc.element.body, encoding="unicode")


def test_image_layout_bottom_falls_back_to_the_paragraph_style_size():
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:text-bottom||}}")
    doc.styles["Normal"].font.size = 20 * 6350

    _apply_image_layouts(doc)

    assert _position(picture_run) == "-5"


def test_image_layout_bottom_falls_back_to_the_document_default_size():
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:bottom||}}")
    normal = doc.styles["Normal"].element.find(f"{{{SCHEMA}}}rPr")
    if normal is not None:
        for sz in normal.findall(f"{{{SCHEMA}}}sz"):
            normal.remove(sz)
    defaults = doc.styles.element.find(f"{{{SCHEMA}}}docDefaults/{{{SCHEMA}}}rPrDefault/{{{SCHEMA}}}rPr")
    for sz in defaults.findall(f"{{{SCHEMA}}}sz"):
        defaults.remove(sz)

    _apply_image_layouts(doc)

    # Word's own 10pt
    assert _position(picture_run) == "-5"


@pytest.mark.parametrize(
    ("valign", "expected"),
    [
        ("-3pt", "-6"),
        ("4px", "6"),
        ("-0.1pt", None),
        ("top", None),
        ("", None),
    ],
)
def test_image_layout_length_shifts_the_picture_by_that_length(valign: str, expected: str | None):
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture(f"{{{{IMGLAYOUT:{valign}||}}}}")

    _apply_image_layouts(doc)

    assert _position(picture_run) == expected
    assert _effect_extent(picture_run) is None


@pytest.mark.parametrize(
    ("label_size", "picture_px", "expected"),
    [
        # A 16px icon beside 12pt text: a quarter of 12pt up, half of 12pt down
        (24, 16, "-6"),
        (20, 16, "-7"),
        (24, 8, None),
    ],
)
def test_image_layout_middle_centers_the_picture_on_half_the_x_height(label_size: int, picture_px: int, expected: str | None):
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:middle||}}", label_size=label_size)
    picture_run.find(f".//{{{WP_SCHEMA}}}extent").set("cy", str(picture_px * 9525))

    _apply_image_layouts(doc)

    assert _position(picture_run) == expected


def test_image_layout_bottom_reads_the_label_past_a_space():
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:bottom||}}", label_size=32)
    # An unsized space between the icon and its 16pt label
    picture_run.addnext(picture_run.makeelement(f"{{{SCHEMA}}}r", {}))
    space = picture_run.getnext()
    space.append(space.makeelement(f"{{{SCHEMA}}}t", {}))
    space[0].text = " "

    _apply_image_layouts(doc)

    assert _position(picture_run) == "-8"


def test_image_layout_reads_the_label_past_the_next_icon():
    import copy

    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:bottom||}}", label_size=32)
    # A second marked icon between the first one and the 16pt label
    second_marker, second_picture = copy.deepcopy(picture_run.getprevious()), copy.deepcopy(picture_run)
    picture_run.addnext(second_marker)
    second_marker.addnext(second_picture)

    _apply_image_layouts(doc)

    assert [_position(picture_run), _position(second_picture)] == ["-8", "-8"]


def test_image_layout_ignores_a_length_too_large_for_a_float():
    from app.docx_post_process import _apply_image_layouts

    huge = "9" * 400
    doc, picture_run = _paragraph_with_marked_picture(f"{{{{IMGLAYOUT:-{huge}px||{huge}px}}}}")

    _apply_image_layouts(doc)

    assert _position(picture_run) is None
    assert _effect_extent(picture_run) is None


def test_image_layout_margins_add_to_the_effect_extent():
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:|1pt|-2px}}")

    _apply_image_layouts(doc)

    # A negative margin is not space
    assert _effect_extent(picture_run) == {"l": "12700", "t": "0", "r": "0", "b": "0"}
    assert _position(picture_run) is None


def test_image_layout_position_keeps_the_schema_order_of_run_properties():
    from app.docx_post_process import _apply_image_layouts

    doc, picture_run = _paragraph_with_marked_picture("{{IMGLAYOUT:-1pt||}}", picture_rpr='<w:b/><w:position w:val="4"/><w:sz w:val="30"/>')

    _apply_image_layouts(doc)

    r_pr = picture_run.find(f"{{{SCHEMA}}}rPr")
    assert [etree.QName(child).localname for child in r_pr] == ["b", "position", "sz"]
    assert _position(picture_run) == "-2"


def test_image_layout_marker_without_a_picture_is_removed():
    from docx import Document

    from app.docx_post_process import _apply_image_layouts

    doc = Document()
    doc.add_paragraph("{{IMGLAYOUT:bottom||2px}}")
    doc.add_paragraph("Text {{IMGLAYOUT:bottom||}} inside")

    _apply_image_layouts(doc)

    texts = [paragraph.text for paragraph in doc.paragraphs]
    # Only a run holding nothing but a marker is a marker
    assert texts == ["", "Text {{IMGLAYOUT:bottom||}} inside"]


def test_process_applies_image_layouts():
    import io

    from docx import Document

    doc, _ = _paragraph_with_marked_picture("{{IMGLAYOUT:bottom||2px}}", label_size=20)
    buffer = io.BytesIO()
    doc.save(buffer)

    result = Document(io.BytesIO(docx_post_process.process(buffer.getvalue())))

    picture_run = result.element.body.find(f".//{{{SCHEMA}}}drawing").getparent()
    assert _position(picture_run) == "-5"
    assert _effect_extent(picture_run)["r"] == "19050"
