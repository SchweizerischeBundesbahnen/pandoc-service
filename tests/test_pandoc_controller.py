import io
import os
import platform
import zipfile
from unittest.mock import MagicMock, patch

import pytest
from werkzeug.datastructures import FileStorage

# Import the module to test
from app.PandocController import app, version


@pytest.fixture
def client():
    """Create a test client for the Flask app."""
    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client


def test_version_endpoint():
    """Test the version endpoint."""
    with patch("pandoc.configure") as mock_configure, patch.dict(os.environ, {"PANDOC_SERVICE_VERSION": "1.0.0", "PANDOC_SERVICE_BUILD_TIMESTAMP": "2024-03-27"}):
        # Mock pandoc configuration
        mock_configure.return_value = {"version": "3.1.9"}

        # Simulate calling the version endpoint
        result = version()

        # Assertions
        assert result["python"] == platform.python_version()
        assert result["pandoc"] == "3.1.9"
        assert result["pandocService"] == "1.0.0"
        assert result["timestamp"] == "2024-03-27"


def test_get_docx_template():
    """Test the docx template retrieval endpoint."""
    with (
        patch("importlib.metadata.version", return_value="3.0.0"),
        patch("subprocess.run") as mock_subprocess,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.open", create=True) as mock_path_open,
    ):
        # Mock file content and handling
        mock_docx_content = "Mock DOCX template content"
        mock_file = MagicMock()
        mock_file.read.return_value = mock_docx_content.encode("utf-8")
        mock_path_open.return_value.__enter__.return_value = mock_file
        mock_subprocess.return_value = MagicMock()

        # Create test client
        test_client = app.test_client()

        # Send GET request
        response = test_client.get("/docx-template")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert response.data == mock_docx_content.encode("utf-8")


def test_convert_endpoint():
    """Test the conversion endpoint."""
    with patch("pandoc.read") as mock_read, patch("pandoc.write") as mock_write, patch("app.DocxPostProcess.replace_table_properties", side_effect=lambda x: x):
        # Create a test client
        test_client = app.test_client()

        # Mock pandoc read and write
        mock_read.return_value = MagicMock()
        mock_write.return_value = create_mock_docx()

        # Prepare test data
        source_format = "markdown"
        target_format = "docx"
        test_content = b"# Test Markdown Content"

        # Send POST request
        response = test_client.post(f"/convert/{source_format}/to/{target_format}", data=test_content)

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        # Verify pandoc methods were called
        mock_read.assert_called_once_with(test_content, format=source_format)
        mock_write.assert_called_once()


def test_convert_docx_with_template():
    """Test converting to DOCX with an optional template."""
    with (
        patch("pandoc.read") as mock_read,
        patch("pandoc.write") as mock_write,
        patch("time.time", return_value=1234567890),
        patch("pathlib.Path.unlink"),  # Mock file deletion to prevent errors
    ):
        # Create a test client
        test_client = app.test_client()

        # Mock pandoc read and write
        mock_read.return_value = MagicMock()
        mock_write.return_value = create_mock_docx()

        # Prepare test data
        source_format = "markdown"
        test_content = "# Test Markdown Content"

        # Create a mock template file
        template_file = FileStorage(stream=io.BytesIO(create_mock_docx()), filename="template.docx", content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Send POST request with source and template
        response = test_client.post(f"/convert/{source_format}/to/docx-with-template", data={"source": test_content, "template": template_file}, content_type="multipart/form-data")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        # Verify pandoc methods were called
        mock_read.assert_called_once_with(test_content, format=source_format)
        mock_write.assert_called_once()


def test_convert_endpoint_error_handling():
    """Test error handling in conversion endpoints."""
    with patch("pandoc.read") as mock_read:
        # Create a test client
        test_client = app.test_client()

        # Simulate decoding error
        mock_read.side_effect = UnicodeDecodeError("utf-8", b"", 0, 1, "error")

        # Send POST request
        response = test_client.post(
            "/convert/markdown/to/docx",
            data=b"\xff\xfe",  # Invalid UTF-8 data
        )

        # Assertions
        assert response.status_code == 400
        assert b"Cannot decode request body" in response.data


def test_convert_docx_with_template_no_source():
    """Test conversion endpoint with missing source."""
    # Create a test client
    test_client = app.test_client()

    # Send POST request without source
    response = test_client.post("/convert/markdown/to/docx-with-template")

    # Assertions
    assert response.status_code == 400
    assert b"No data or file provided" in response.data


def create_mock_docx(files: dict[str, bytes] = None) -> bytes:
    """
    Create a minimal valid DOCX file for testing with additional required files

    :param files: Optional dictionary of additional files to include in the DOCX
    :return: Bytes representing a valid DOCX file
    """
    # Default files if not provided
    default_files = {
        "[Content_Types].xml": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""",
        "_rels/.rels": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        "word/document.xml": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Mock Document</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>""",
        "word/_rels/document.xml.rels": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        "word/styles.xml": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr/>
        </w:rPrDefault>
    </w:docDefaults>
</w:styles>""",
    }

    # Merge default files with any provided files
    if files:
        default_files.update(files)

    # Create a BytesIO buffer to simulate a file
    buffer = io.BytesIO()

    # Create a zip file (which is the structure of a DOCX)
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        # Add all files to the zip
        for filename, content in default_files.items():
            zf.writestr(filename, content)

    # Reset buffer position
    buffer.seek(0)
    return buffer.getvalue()
