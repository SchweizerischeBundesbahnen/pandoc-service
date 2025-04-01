import io
from pathlib import Path

import pandoc  # type: ignore
from docx import Document

from app import DocxPostProcess

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

EMUS_IN_INCH = 914400


def test_replace_table_properties():
    with Path("test-data/big_image_in_base64.txt").open(encoding="utf-8") as big_image_in_base64:
        source_html = SOURCE_HTML_WITH_TABLE.format(big_image_in_base64.read())

    initial_document = pandoc.read(source_html, format="html")
    pre_process_output = pandoc.write(initial_document, format="docx")
    post_process_output = DocxPostProcess.replace_table_properties(pre_process_output)

    final_document = Document(io.BytesIO(post_process_output))

    assert len(final_document.tables) > 0
    table = final_document.tables[0]._element
    table_properties = table.find("w:tblPr", {"w": WORD_PROCESSING_ML_MAIN_SCHEMA})
    if table_properties is not None:
        # Look for the width attribute
        table_width = table_properties.find("w:tblW", {"w": WORD_PROCESSING_ML_MAIN_SCHEMA})

        if table_width is not None:
            # Get width value and type
            width_value = table_width.get(WORD_PROCESSING_ML_MAIN_SCHEMA_IN_BRACKETS + "w")
            width_type = table_width.get(WORD_PROCESSING_ML_MAIN_SCHEMA_IN_BRACKETS + "type")

            assert width_value == "5000"
            assert width_type == "pct"

    images = table.findall(".//pic:pic", {"pic": DRAWING_ML_PICTURE_SCHEMA})
    assert len(images) > 0
    image = images[0]
    image_extent = image.find(".//a:ext", {"a": DRAWING_ML_MAIN_SCHEMA})
    image_width_in_inches = round(int(image_extent.get("cx", 0)) / EMUS_IN_INCH)
    assert image_width_in_inches == 6
