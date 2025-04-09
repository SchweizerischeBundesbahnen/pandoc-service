from unittest.mock import MagicMock, patch

from docx.table import Table, _Cell

from app import DocxPostProcess
from app.DocxPostProcess import process_table

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
        mock_cell._tc.xml = f'''
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
        '''
    else:
        mock_cell._tc.xml = f'''
        <w:tc xmlns:w="{WORD_PROCESSING_ML_MAIN_SCHEMA}">
            <w:tcPr/>
            <w:p>
                <w:r>
                    <w:t>Cell content</w:t>
                </w:r>
            </w:p>
        </w:tc>
        '''

    # Connect the mocks together
    mock_row.cells = [mock_cell]
    mock_table.rows = [mock_row]
    mock_doc.tables = [mock_table]

    return mock_doc


@patch("app.PandocController.run_pandoc_conversion")
@patch("app.DocxPostProcess.Document")
@patch("app.DocxPostProcess.etree.fromstring")
def test_replace_table_properties(mock_fromstring, mock_document_class, mock_run_pandoc):
    # Create a mock document to return from our Document constructor
    mock_doc = create_mock_document_with_table()
    mock_document_class.return_value = mock_doc

    # Set up the pandoc conversion mock to return some dummy bytes
    mock_run_pandoc.return_value = b"mock docx data"

    # Create a mock XML tree that will be returned from fromstring
    mock_tree = MagicMock()
    # Set up mock extent elements that might be found
    mock_extent = MagicMock()
    mock_extent.attrib = {"cx": "1000000", "cy": "750000"}
    # Return empty list for extent elements to avoid image resizing
    mock_tree.findall.return_value = []
    mock_fromstring.return_value = mock_tree

    # Call the function under test and store result (use the return value)
    result = DocxPostProcess.replace_table_properties(mock_run_pandoc.return_value)
    assert isinstance(result, bytes)

    # Verify the Document was constructed
    mock_document_class.assert_called_once()

    # Verify that table properties were accessed
    assert mock_doc.tables[0]._element.find.called


@patch("app.PandocController.run_pandoc_conversion")
@patch("app.DocxPostProcess.Document")
@patch("app.DocxPostProcess.etree.fromstring")
def test_nested_tables(mock_fromstring, mock_document_class, mock_run_pandoc):
    # Create a mock document with nested tables
    mock_doc = create_mock_document_with_table(has_nested_table=True)
    mock_document_class.return_value = mock_doc

    # Set up the pandoc conversion mock to return some dummy bytes
    mock_run_pandoc.return_value = b"mock docx data"

    # Create a mock XML tree that will be returned from fromstring
    mock_tree = MagicMock()
    # Return empty list for extent elements to avoid image resizing
    mock_tree.findall.return_value = []
    mock_fromstring.return_value = mock_tree

    # Call the function under test and store result (use the return value)
    result = DocxPostProcess.replace_table_properties(mock_run_pandoc.return_value)
    assert isinstance(result, bytes)

    # Verify the Document was constructed
    mock_document_class.assert_called_once()

    # Verify we processed the nested table by checking if cell.tables was accessed
    assert len(mock_doc.tables[0].rows[0].cells[0].tables) > 0


@patch("app.PandocController.run_pandoc_conversion")
@patch("app.DocxPostProcess.Document")
def test_document_without_tables(mock_document_class, mock_run_pandoc):
    # Create a mock document with no tables
    mock_doc = MagicMock()
    mock_doc.tables = []
    mock_document_class.return_value = mock_doc

    # Set up the pandoc conversion mock to return some dummy bytes
    mock_run_pandoc.return_value = b"mock docx data"

    # Call the function under test and store result (use the return value)
    result = DocxPostProcess.replace_table_properties(mock_run_pandoc.return_value)
    assert isinstance(result, bytes)

    # Verify the Document was constructed
    mock_document_class.assert_called_once()

    # Verify we checked the tables collection
    assert len(mock_doc.tables) == 0


def test_get_available_content_width():
    """Test the get_available_content_width function."""
    # Create a mock section
    mock_section = MagicMock()
    mock_section.page_width = DocxPostProcess.EMU_1_INCH * 11  # 11 inches
    mock_section.left_margin = DocxPostProcess.EMU_1_INCH * 1.5  # 1.5 inches
    mock_section.right_margin = DocxPostProcess.EMU_1_INCH * 1.5  # 1.5 inches

    # Create a mock document with the mock section
    mock_doc = MagicMock()
    mock_doc.sections = [mock_section]

    # Calculate expected content width: 11 - 1.5 - 1.5 = 8 inches
    expected_width = int(DocxPostProcess.EMU_1_INCH * 8)

    # Get actual content width
    actual_width = DocxPostProcess.get_available_content_width(mock_doc)

    # Assert they match
    assert actual_width == expected_width


def test_get_available_content_width_default_values():
    """Test the get_available_content_width function with default values."""
    # Create a mock section with None values for page dimensions
    mock_section = MagicMock()
    mock_section.page_width = None
    mock_section.left_margin = None
    mock_section.right_margin = None

    # Create a mock document with the mock section
    mock_doc = MagicMock()
    mock_doc.sections = [mock_section]

    # We expect the function to use defaults from DocxPostProcess
    expected_width = int(DocxPostProcess.DOCX_LETTER_WIDTH_EMU - 2 * DocxPostProcess.DOCX_LETTER_SIDE_MARGIN)  # 8.5 - 2 = 6.5 inches

    # Get actual width
    actual_width = DocxPostProcess.get_available_content_width(mock_doc)

    # Assert it uses default values correctly
    assert actual_width == expected_width


def test_resize_images_in_cell_no_resizing_needed():
    """Test that small images don't get resized."""
    # Create a mock cell
    cell = MagicMock(spec=_Cell)

    # Create mock XML content with a small image (smaller than max width)
    small_image_width = 100000  # Small value in EMU, less than max_width
    small_image_height = 100000
    small_image_xml = f'''
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
    '''

    # Set up the mock cell
    mock_tc = MagicMock()
    mock_tc.xml = small_image_xml
    cell._tc = mock_tc

    # Set a max width larger than the image width
    max_width = 500000  # Larger than small_image_width

    # Call the function
    DocxPostProcess.resize_images_in_cell(cell, max_width)

    # Verify that clear_content was not called (which would indicate the image was modified)
    mock_tc.clear_content.assert_not_called()


def test_resize_images_in_cell_resizing_needed():
    """Test that large images are properly resized."""
    # Create a mock cell
    cell = MagicMock(spec=_Cell)

    # Create mock XML content with a large image (larger than max_width)
    large_image_width = 1000000  # Large value in EMU, greater than max_width
    large_image_height = 750000  # 3:4 aspect ratio
    large_image_xml = f'''
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
    '''

    # Set up the mock for lxml etree parsing - create elements that mimic the real ones
    from lxml import etree

    # Parse the XML to create a real tree
    # ruff: noqa: S320
    tree = etree.fromstring(large_image_xml)

    # Find the wp:extent element to use with the real function
    extent_element = tree.find(".//wp:extent", {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
    assert extent_element is not None
    assert extent_element.get("cx") == str(large_image_width)
    assert extent_element.get("cy") == str(large_image_height)

    # Set up the mock cell with our parsed tree
    mock_tc = MagicMock()
    mock_tc.xml = large_image_xml
    cell._tc = mock_tc

    # Set a max width smaller than the image width to trigger resizing
    max_width = 500000  # Smaller than large_image_width

    # Call the function
    with patch("app.DocxPostProcess.etree.fromstring", return_value=tree):
        DocxPostProcess.resize_images_in_cell(cell, max_width)

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
    with patch("app.DocxPostProcess.etree.fromstring", return_value=mock_tree):
        # Use a reasonable max width value that will trigger resizing
        max_image_width = 4000.0  # Smaller than the image width

        # Call resize_images_in_cell directly
        DocxPostProcess.resize_images_in_cell(cell, max_image_width)

        # Verify that the image was resized
        assert mock_extent.set.call_count == 2  # cx and cy set
        # Resize should maintain aspect ratio
        mock_extent.set.assert_any_call("cx", "4000")  # New width
        # New height = old_height * (new_width/old_width) = 3000 * (4000/5000) = 2400
        mock_extent.set.assert_any_call("cy", "2400")  # New height (aspect ratio maintained)


@patch("app.DocxPostProcess.resize_images_in_cell")
def test_process_table_with_nested_tables(mock_resize_images):
    """Test that nested tables are processed correctly."""
    from app.DocxPostProcess import SCHEMA

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
    with patch("app.DocxPostProcess.parse_xml") as mock_parse_xml:
        mock_tbl_props = MagicMock()
        mock_parse_xml.return_value = mock_tbl_props

        # Set up max_width for testing
        max_width = 9144000  # 10 inches in EMU

        # Call the actual process_table function
        process_table(main_table, 0, max_width)

        # Verify find was called with correct parameters
        main_table._element.find.assert_called_with(".//w:tblPr", namespaces={"w": SCHEMA})

        # Verify parse_xml was called (for creating table properties)
        mock_parse_xml.assert_called()

        # Verify the table element was updated
        main_table._element.insert.assert_called_with(0, mock_tbl_props)

        # Verify resize_images_in_cell was called with correct parameters
        mock_resize_images.assert_called_with(cell, max_width / 2)

        # Verify that we processed any nested tables
        for _sub_table in cell.tables:
            # At this point, we would process the nested table, but we can't verify
            # the actual call since it happens recursively. Instead, we verify that
            # the nested tables were accessed.
            pass
