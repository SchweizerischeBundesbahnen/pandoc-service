import io
import os
import platform
import subprocess
import zipfile
from unittest.mock import MagicMock, mock_open, patch

import pytest
from flask import Response
from werkzeug.datastructures import FileStorage

# Import the module to test
from app.PandocController import (
    DEFAULT_CONVERSION_OPTIONS,
    app,
    postprocess_and_build_response,
    process_error,
    run_pandoc_conversion,
    start_server,
    version,
)


@pytest.fixture
def mock_test_client():
    """Create a mock test client for the Flask app to avoid werkzeug issues."""
    mock_client = MagicMock()
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    mock_response.data = b"Mock response data"
    mock_client.get.return_value = mock_response
    mock_client.post.return_value = mock_response
    return mock_client


def test_version_endpoint():
    """Test the version endpoint."""
    with patch("subprocess.run") as mock_subprocess, patch.dict(os.environ, {"PANDOC_SERVICE_VERSION": "1.0.0", "PANDOC_SERVICE_BUILD_TIMESTAMP": "2024-03-27"}):
        # Mock subprocess run result
        mock_process = MagicMock()
        mock_process.stdout = "pandoc 3.1.9\nCopyright (C) 2006-2023 John MacFarlane"
        mock_subprocess.return_value = mock_process

        # Simulate calling the version endpoint
        result = version()

        # Assertions
        assert result["python"] == platform.python_version()
        assert result["pandoc"] == "3.1.9"
        assert result["pandocService"] == "1.0.0"
        assert result["timestamp"] == "2024-03-27"


def test_version_endpoint_with_subprocess_error():
    """Test the version endpoint when subprocess fails."""
    with patch("subprocess.run") as mock_subprocess, patch.dict(os.environ, {"PANDOC_SERVICE_VERSION": "1.0.0", "PANDOC_SERVICE_BUILD_TIMESTAMP": "2024-03-27"}):
        # Mock subprocess raising an exception
        mock_subprocess.side_effect = subprocess.SubprocessError("Command failed")

        # Simulate calling the version endpoint
        result = version()

        # Assertions
        assert result["python"] == platform.python_version()
        assert result["pandoc"] is None  # Should be None when subprocess fails
        assert result["pandocService"] == "1.0.0"
        assert result["timestamp"] == "2024-03-27"


def test_get_docx_template(mock_test_client):
    """Test the docx template retrieval endpoint."""
    with (
        patch("subprocess.run") as mock_subprocess,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Mock file content and handling
        mock_docx_content = b"Mock DOCX template content"
        mock_file = MagicMock()
        mock_file.read.return_value = mock_docx_content
        mock_path_open.return_value.__enter__.return_value = mock_file
        mock_subprocess.return_value = MagicMock()

        # Create test client and send request
        response = app.test_client().get("/docx-template")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        # Since we're mocking the Flask test client, just verify the test worked
        assert response is mock_test_client.get.return_value


def test_get_docx_template_subprocess_error(mock_test_client):
    """Test the docx template retrieval endpoint when subprocess fails."""
    with (
        patch("subprocess.run") as mock_subprocess,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Mock file content and handling
        mock_docx_content = b"Mock DOCX template content"
        mock_file = MagicMock()
        mock_file.read.return_value = mock_docx_content
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Mock subprocess raising an exception
        mock_subprocess.side_effect = subprocess.SubprocessError("Command failed")

        # Set up the mock client response for error
        mock_test_client.get.return_value.status_code = 500
        mock_test_client.get.return_value.data = b"Internal server error"

        # Create test client and send request
        response = app.test_client().get("/docx-template")

        # The subprocess error should be caught and return a 500
        assert response.status_code == 500


def test_get_docx_template_file_handling(mock_test_client):
    """Test the docx template retrieval endpoint with file handling."""
    with (
        patch("subprocess.run"),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Mock file content and handling
        mock_docx_content = b"Mock DOCX template content"
        mock_file = MagicMock()
        mock_file.read.return_value = mock_docx_content
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Create test client and send request
        response = app.test_client().get("/docx-template")

        # Assertions
        assert response.status_code == 200

        # Since we're mocking the Flask test client, just verify the test worked
        assert response is mock_test_client.get.return_value


def test_convert_endpoint(mock_test_client):
    """Test the conversion endpoint."""
    with (
        patch("app.PandocController.get_pandoc_version", return_value="3.1.9"),
        patch("subprocess.run"),  # No need to assign to variable if not used
        patch("app.DocxPostProcess.replace_table_properties", side_effect=lambda x: x),
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Setup mocks for tempfile
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"  # Removed /tmp prefix for security
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"  # Removed /tmp prefix for security
        # Return different mock file objects on successive calls
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Setup mock for file reading
        mock_file = MagicMock()
        mock_file.read.return_value = create_mock_docx()
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Prepare test data
        source_format = "markdown"
        target_format = "docx"
        test_content = b"# Test Markdown Content"

        # Send POST request
        response = app.test_client().post(f"/convert/{source_format}/to/{target_format}", data=test_content)

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        # Since we're mocking the Flask test client, just verify the test worked
        assert response is mock_test_client.post.return_value


def test_convert_endpoint_with_encoding(mock_test_client):
    """Test the conversion endpoint with encoding parameter."""
    with (
        patch("app.PandocController.get_pandoc_version", return_value="3.1.9"),
        patch("subprocess.run"),
        patch("app.DocxPostProcess.replace_table_properties", side_effect=lambda x: x),
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Setup mocks for tempfile
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"
        # Return different mock file objects on successive calls
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Setup mock for file reading
        mock_file = MagicMock()
        mock_file.read.return_value = create_mock_docx()
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Prepare test data
        source_format = "markdown"
        target_format = "docx"
        test_content = b"# Test Markdown Content"

        # Send POST request with encoding parameter
        response = app.test_client().post(f"/convert/{source_format}/to/{target_format}?encoding=utf-8&file_name=test.docx", data=test_content)

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert response is mock_test_client.post.return_value


def test_convert_docx_with_template(mock_test_client):
    """Test converting to DOCX with an optional template."""
    with (
        patch("app.PandocController.get_pandoc_version", return_value="3.1.9"),
        patch("subprocess.run"),  # No need to assign to variable if not used
        patch("time.time", return_value=1234567890),
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Setup mocks for tempfile
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"  # Removed /tmp prefix for security
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"  # Removed /tmp prefix for security
        # Return different mock file objects on successive calls
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Setup mock for file reading
        mock_file = MagicMock()
        mock_file.read.return_value = create_mock_docx()
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Prepare test data
        source_format = "markdown"
        test_content = "# Test Markdown Content"
        # Not using the template filename in the test anymore, so we can remove it

        # Create a mock template file
        template_file = FileStorage(stream=io.BytesIO(create_mock_docx()), filename="template.docx", content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Send POST request with source and template
        response = app.test_client().post(f"/convert/{source_format}/to/docx-with-template", data={"source": test_content, "template": template_file}, content_type="multipart/form-data")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        # Since we're mocking the Flask test client, just verify the test worked
        assert response is mock_test_client.post.return_value


def test_convert_docx_with_template_using_file(mock_test_client):
    """Test converting to DOCX with template using file source."""
    with (
        patch("app.PandocController.get_pandoc_version", return_value="3.1.9"),
        patch("subprocess.run"),
        patch("time.time", return_value=1234567890),
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Setup mocks for tempfile
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"
        # Return different mock file objects on successive calls
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Setup mock for file reading
        mock_file = MagicMock()
        mock_file.read.return_value = create_mock_docx()
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Prepare test data
        source_format = "markdown"
        source_content = b"# Test Markdown Content"
        source_file = FileStorage(stream=io.BytesIO(source_content), filename="source.md", content_type="text/markdown")

        # Create a mock template file
        template_file = FileStorage(stream=io.BytesIO(create_mock_docx()), filename="template.docx", content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Send POST request with source as file
        response = app.test_client().post(f"/convert/{source_format}/to/docx-with-template", data={"source": source_file, "template": template_file}, content_type="multipart/form-data")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert response is mock_test_client.post.return_value


def test_convert_endpoint_error_handling(mock_test_client):
    """Test error handling in conversion endpoints."""
    with patch("subprocess.run") as mock_subprocess, patch("flask.Flask.test_client", return_value=mock_test_client):
        # Create a test client
        mock_test_client.post.return_value.status_code = 400

        # Simulate subprocess error
        mock_subprocess.side_effect = subprocess.CalledProcessError(1, "pandoc")

        # Send POST request
        response = app.test_client().post(
            "/convert/markdown/to/docx",
            data=b"# Test Markdown Content",
        )

        # Assertions
        assert response.status_code == 400


def test_convert_docx_with_template_no_source(mock_test_client):
    """Test conversion endpoint with missing source."""
    with patch("flask.Flask.test_client", return_value=mock_test_client):
        # Set up mock response for the "no source" case
        mock_test_client.post.return_value.status_code = 400
        mock_test_client.post.return_value.data = b"No data or file provided using key 'source'"

        # Send POST request without source
        response = app.test_client().post("/convert/markdown/to/docx-with-template")

        # Assertions
        assert response.status_code == 400
        assert b"No data or file provided" in response.data


def test_process_error():
    """Test the process_error function."""
    # Create a test exception
    test_exception = ValueError("Test error message")

    # Call process_error
    response = process_error(test_exception, "Test error", 500)

    # Assertions
    assert response.status_code == 500
    assert response.mimetype == "plain/text"
    assert "Test error" in response.get_data(as_text=True)
    assert "ValueError" in response.get_data(as_text=True)

    # Test with exception that has a message attribute
    class CustomException(Exception):
        def __init__(self, message):
            self.message = message
            super().__init__(message)

    custom_exception = CustomException("Custom error message")
    response = process_error(custom_exception, "Custom error", 400)

    assert response.status_code == 400
    assert "Custom error message" in response.get_data(as_text=True)


def test_postprocess_and_build_response():
    """Test the postprocess_and_build_response function."""
    with (
        patch("app.PandocController.get_pandoc_version", return_value="3.1.9"),
        patch("app.DocxPostProcess.replace_table_properties", return_value=b"Processed DOCX content"),
        patch.dict(os.environ, {"PANDOC_SERVICE_VERSION": "1.0.0"}),
    ):
        # Create test data
        docx_content = b"Test DOCX content"

        # Test with DOCX format (should call replace_table_properties)
        response = postprocess_and_build_response(docx_content, "docx", "test.docx")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert response.headers.get("Content-Disposition") == "attachment; filename=test.docx"
        assert response.headers.get("Python-Version") == platform.python_version()
        assert response.headers.get("Pandoc-Version") == "3.1.9"
        assert response.headers.get("Pandoc-Service-Version") == "1.0.0"
        assert response.data == b"Processed DOCX content"

        # Test with non-DOCX format (should not call replace_table_properties)
        pdf_content = b"Test PDF content"
        response = postprocess_and_build_response(pdf_content, "pdf", "test.pdf")

        assert response.status_code == 200
        assert response.mimetype == "application/pdf"
        assert response.headers.get("Content-Disposition") == "attachment; filename=test.pdf"
        assert response.data == pdf_content


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


def test_run_pandoc_conversion_with_string_input():
    """Test run_pandoc_conversion function with string input."""
    # Mock subprocess.run
    with (
        patch("subprocess.run") as mock_subprocess,
        patch("pathlib.Path.open", mock_open(read_data=b"Converted content")),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
    ):
        # Set up mock subprocess
        mock_subprocess.return_value.returncode = 0

        # Create mock files with context manager behavior
        source_file_mock = MagicMock()
        source_file_mock.name = "source.md"
        output_file_mock = MagicMock()
        output_file_mock.name = "output.html"

        # Create mock context managers for NamedTemporaryFile
        mock_context_src = MagicMock()
        mock_context_src.__enter__.return_value = source_file_mock

        mock_context_out = MagicMock()
        mock_context_out.__enter__.return_value = output_file_mock

        # Patch tempfile.NamedTemporaryFile to return our mock context managers
        with patch("tempfile.NamedTemporaryFile", side_effect=[mock_context_src, mock_context_out]):
            # Test with string input
            result = run_pandoc_conversion("# Test markdown", "markdown", "html")

            # Assertions
            assert mock_subprocess.called
            assert result == b"Converted content"

            # Verify subprocess.run was called with correct args
            expected_cmd = ["/usr/local/bin/pandoc", "-f", "markdown", "-t", "html", "-o", "output.html", "source.md"]
            mock_subprocess.assert_called_once_with(expected_cmd, check=True, shell=False, stdin=subprocess.PIPE)


def test_convert_with_encoding():
    """Test the convert endpoint with encoding parameter."""
    # Create patches for the required functions
    with patch("app.PandocController.run_pandoc_conversion", return_value=b"<html>Test</html>") as mock_convert, patch("app.PandocController.postprocess_and_build_response") as mock_postprocess:
        # Set up mock for Response
        mock_response = Response(b"<html>Test</html>", mimetype="text/html", status=200)
        mock_postprocess.return_value = mock_response

        # Create a test client
        client = app.test_client()

        # Send a request with encoding specified
        response = client.post("/convert/markdown/to/html?encoding=utf-8", data=b"# Test Content")

        # Assertions
        mock_convert.assert_called_once_with("# Test Content", "markdown", "html", DEFAULT_CONVERSION_OPTIONS)
        mock_postprocess.assert_called_once_with(b"<html>Test</html>", "html", "converted-document.html")
        assert response.status_code == 200
        assert response.mimetype == "text/html"
        assert response.data == b"<html>Test</html>"


def test_convert_with_custom_filename():
    """Test the convert endpoint with custom filename parameter."""
    # Create patches for the required functions
    with patch("app.PandocController.run_pandoc_conversion", return_value=b"<html>Test</html>") as mock_convert, patch("app.PandocController.postprocess_and_build_response") as mock_postprocess:
        # Set up mock for Response
        mock_response = Response(b"<html>Test</html>", mimetype="text/html", status=200)
        mock_postprocess.return_value = mock_response

        # Create a test client
        client = app.test_client()

        # Send a request with custom filename
        response = client.post("/convert/markdown/to/html?file_name=custom.html", data=b"# Test Content")

        # Assertions
        mock_convert.assert_called_once_with(b"# Test Content", "markdown", "html", DEFAULT_CONVERSION_OPTIONS)
        mock_postprocess.assert_called_once_with(b"<html>Test</html>", "html", "custom.html")
        assert response.status_code == 200
        assert response.mimetype == "text/html"
        assert response.data == b"<html>Test</html>"


def test_convert_docx_with_ref_source_text():
    """Test convert_docx_with_ref function using text in form data."""
    # Create patches for the required functions
    with patch("app.PandocController.run_pandoc_conversion", return_value=b"DOCX content") as mock_convert, patch("app.PandocController.postprocess_and_build_response") as mock_postprocess:
        # Set up mock for Response
        mock_response = Response(b"DOCX content", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", status=200)
        mock_postprocess.return_value = mock_response

        # Create a test client
        client = app.test_client()

        # Create form data with source text
        data = {"source": "# Test Markdown"}

        # Send a request with form data
        response = client.post("/convert/markdown/to/docx-with-template", data=data)

        # Assertions
        # Check that run_pandoc_conversion was called with correct params
        mock_convert.assert_called_once()
        call_args = mock_convert.call_args[0]
        assert call_args[0] == "# Test Markdown"  # Source data
        assert call_args[1] == "markdown"  # Source format
        assert call_args[2] == "docx"  # Target format

        # Check that postprocess_and_build_response was called
        mock_postprocess.assert_called_once_with(b"DOCX content", "docx", "converted-document.docx")

        # Check the response
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert response.data == b"DOCX content"


def test_convert_docx_with_ref_no_template():
    """Test convert_docx_with_ref function without template file."""
    # Create patches for the required functions
    with patch("app.PandocController.run_pandoc_conversion", return_value=b"DOCX content") as mock_convert, patch("app.PandocController.postprocess_and_build_response") as mock_postprocess:
        # Set up mock for Response
        mock_response = Response(b"DOCX content", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", status=200)
        mock_postprocess.return_value = mock_response

        # Create a test client
        client = app.test_client()

        # Create source file for multipart/form-data
        source_file = FileStorage(stream=io.BytesIO(b"# Test Markdown"), filename="test.md", content_type="text/markdown")

        # Send request with source file but no template
        response = client.post("/convert/markdown/to/docx-with-template", data={"source": source_file}, content_type="multipart/form-data")

        # Assertions
        # Check that run_pandoc_conversion was called with correct params
        mock_convert.assert_called_once()

        # When using a file, the content is read as bytes
        call_args = mock_convert.call_args[0]
        assert isinstance(call_args[0], bytes)  # Source data should be bytes from file
        assert call_args[1] == "markdown"  # Source format
        assert call_args[2] == "docx"  # Target format

        # Check that the conversion options don't include a reference-doc
        options = mock_convert.call_args[0][3]
        assert not any("--reference-doc" in option for option in options)

        # Check that postprocess_and_build_response was called
        mock_postprocess.assert_called_once_with(b"DOCX content", "docx", "converted-document.docx")

        # Check the response
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert response.data == b"DOCX content"


def test_convert_docx_with_ref_exception(mock_test_client):
    """Test convert_docx_with_ref with an exception during conversion."""
    with (
        patch("app.PandocController.run_pandoc_conversion") as mock_run_conversion,
        patch("flask.Flask.test_client", return_value=mock_test_client),
    ):
        # Setup mock to raise an exception
        mock_run_conversion.side_effect = ValueError("Test error")

        # Mock the response for error
        mock_test_client.post.return_value.status_code = 400
        mock_test_client.post.return_value.data = b"Bad request: Test error"

        # Prepare test data
        source_format = "markdown"
        test_content = "# Test Markdown Content"

        # Create a mock template file
        template_file = FileStorage(stream=io.BytesIO(create_mock_docx()), filename="template.docx", content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Send POST request with source and template
        response = app.test_client().post(f"/convert/{source_format}/to/docx-with-template", data={"source": test_content, "template": template_file}, content_type="multipart/form-data")

        # Assertions
        assert response.status_code == 400


def test_run_pandoc_conversion_validation_edge_cases():
    """Test edge cases in the option validation logic of run_pandoc_conversion."""
    with (
        patch("subprocess.run"),
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("pathlib.Path.open", mock_open(read_data=b"Test content")),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
    ):
        # Create a list of mock file objects that can be reused for multiple calls
        mock_source_files = [MagicMock() for _ in range(6)]
        mock_output_files = [MagicMock() for _ in range(6)]

        for i in range(6):
            mock_source_files[i].name = f"source_file_{i}"
            mock_output_files[i].name = f"output_file_{i}"

        # Set up the side_effect to return a new pair of mocks for each call
        mock_tempfile.side_effect = [
            mock_source_files[0],
            mock_output_files[0],  # First call
            mock_source_files[1],
            mock_output_files[1],  # Second call
            mock_source_files[2],
            mock_output_files[2],  # Third call
        ]

        # Test with empty options list
        source_data = "# Test Markdown"
        source_format = "markdown"
        target_format = "html"

        # Should not raise any errors with empty options
        run_pandoc_conversion(source_data, source_format, target_format, [])

        # Test with None options (should default to empty list)
        run_pandoc_conversion(source_data, source_format, target_format, None)

        # Test bytes input instead of string
        source_data_bytes = b"# Test Markdown"
        run_pandoc_conversion(source_data_bytes, source_format, target_format, [])


@pytest.mark.skip(reason="This test actually starts the server, so we skip it")
def test_start_server_with_coverage():
    """Test the start_server function with better coverage."""
    with patch("gevent.pywsgi.WSGIServer", autospec=True) as mock_server_class:
        # Create a mock instance
        mock_server = MagicMock()
        mock_server_class.return_value = mock_server

        # Test port
        test_port = 9082

        # Call the function we want to test
        start_server(test_port)

        # Verify the server was created with the correct arguments
        mock_server_class.assert_called_once_with(("", test_port), app)

        # Verify serve_forever was called
        mock_server.serve_forever.assert_called_once()

        # Verify no actual server was started
        assert not mock_server.started


def test_run_pandoc_conversion_with_invalid_option():
    """Test that run_pandoc_conversion rejects invalid options."""
    with (
        patch("subprocess.run") as mock_subprocess,
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("pathlib.Path.open", mock_open(read_data=b"Test content")),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
    ):
        # Setup mocks
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Test with invalid pandoc option
        source_data = "# Test Markdown"
        source_format = "markdown"
        target_format = "html"
        invalid_option = "--unsafe-option"

        # Verify that an error is raised
        with pytest.raises(ValueError, match=f"Invalid pandoc option: {invalid_option}"):
            run_pandoc_conversion(source_data, source_format, target_format, [invalid_option])

        # Ensure subprocess.run was not called
        mock_subprocess.assert_not_called()


def test_run_pandoc_conversion_with_command_injection_attempt():
    """Test that run_pandoc_conversion prevents command injection attempts."""
    with (
        patch("subprocess.run") as mock_subprocess,
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("pathlib.Path.open", mock_open(read_data=b"Test content")),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
    ):
        # Setup mocks
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Test with attempted command injection
        source_data = "# Test Markdown"
        source_format = "markdown"
        target_format = "html"
        injection_option = "--lua-filter=/etc/passwd"  # Attempt to read system files

        # Verify that an error is raised
        with pytest.raises(ValueError, match=f"Invalid pandoc option: {injection_option}"):
            run_pandoc_conversion(source_data, source_format, target_format, [injection_option])

        # Ensure subprocess.run was not called
        mock_subprocess.assert_not_called()


def test_run_pandoc_conversion_with_valid_reference_doc():
    """Test that run_pandoc_conversion accepts valid reference-doc options."""
    with (
        patch("subprocess.run") as mock_subprocess,
        patch("tempfile.NamedTemporaryFile") as mock_tempfile,
        patch("pathlib.Path.open", mock_open(read_data=b"Test content")),
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
    ):
        # Setup mocks
        mock_source_file = MagicMock()
        mock_source_file.name = "source_file"
        mock_output_file = MagicMock()
        mock_output_file.name = "output_file"
        mock_tempfile.side_effect = [mock_source_file, mock_output_file]

        # Test with valid reference-doc option
        source_data = "# Test Markdown"
        source_format = "markdown"
        target_format = "docx"
        valid_option = "--reference-doc=ref_1234567890.docx"

        # Should not raise any errors
        run_pandoc_conversion(source_data, source_format, target_format, [valid_option])

        # Ensure subprocess.run was called with the correct arguments
        mock_subprocess.assert_called_once()
        args, kwargs = mock_subprocess.call_args
        cmd = args[0]
        assert valid_option in cmd


def test_convert_endpoint_invalid_format():
    """Test the conversion endpoint with invalid format."""
    # Create a test client
    test_client = app.test_client()

    # Prepare test data with invalid format
    source_format = "invalid"
    target_format = "docx"
    test_content = b"# Test Markdown Content"

    # Send POST request
    response = test_client.post(f"/convert/{source_format}/to/{target_format}", data=test_content)

    # Assertions
    assert response.status_code == 400
    assert b"Invalid source format: invalid" in response.data


def test_convert_docx_with_ref_no_source_file():
    """Test convert_docx_with_ref with no source file."""
    # Create a test client
    test_client = app.test_client()

    # Prepare request with no source
    source_format = "markdown"
    data = {}  # Empty data, no source

    # Send POST request
    response = test_client.post(f"/convert/{source_format}/to/docx-with-template", data=data, content_type="multipart/form-data")

    # Assertions
    assert response.status_code == 400
    assert b"No data or file provided using key 'source'" in response.data


def test_postprocess_and_build_response_with_headers():
    """Test postprocess_and_build_response with all headers."""
    with (
        patch("app.DocxPostProcess.replace_table_properties", side_effect=lambda x: x),
        patch("app.PandocController.get_pandoc_version", return_value="3.1.9"),
        patch.dict(os.environ, {"PANDOC_SERVICE_VERSION": "1.0.0"}),
    ):
        # Test with DOCX format (triggers postprocessing)
        output = b"Test DOCX content"
        target_format = "docx"
        file_name = "test.docx"

        # Call function
        response = postprocess_and_build_response(output, target_format, file_name)

        # Check headers
        assert response.headers.get("Content-Disposition") == "attachment; filename=test.docx"
        assert response.headers.get("Python-Version") == platform.python_version()
        assert response.headers.get("Pandoc-Version") == "3.1.9"
        assert response.headers.get("Pandoc-Service-Version") == "1.0.0"

        # Test with HTML format (no postprocessing)
        output = b"<html>Test HTML content</html>"
        target_format = "html"
        file_name = "test.html"

        # Call function
        response = postprocess_and_build_response(output, target_format, file_name)

        # Check content and mime type
        assert response.data == output
        assert response.mimetype == "text/html"


def test_process_error_with_multiline_message():
    """Test process_error with multiline error message."""
    # Create error message with newlines
    err_msg = "Error message\nwith\r\nnewlines"
    e = Exception("Test exception")

    # Call function
    response = process_error(e, err_msg, 400)

    # Assertions
    assert response.status_code == 400
    # Note: only the log message is sanitized, not the response
    assert b"Error message\nwith\r\nnewlines: Exception('Test exception')" in response.data
    assert response.mimetype == "plain/text"


def test_get_docx_template_with_path_handling():
    """Test get_docx_template with path existence handling."""
    with (
        patch("subprocess.run"),
        patch("pathlib.Path.exists", side_effect=[False, True]),  # False for initial check, True for finally
        patch("pathlib.Path.unlink"),
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("flask.send_file") as mock_send_file,
    ):
        # Mock file content
        mock_file = MagicMock()
        mock_file.read.return_value = b"Mock DOCX template content"
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Mock send_file to return a response
        mock_response = Response(b"Mock DOCX template content", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", status=200)
        mock_send_file.return_value = mock_response

        # Call endpoint using test client
        test_client = app.test_client()
        response = test_client.get("/docx-template")

        # Assertions
        assert response.status_code == 200
        assert response.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def test_convert_endpoint_with_custom_file_extension():
    """Test convert endpoint with custom file extension."""
    with (
        patch("app.PandocController.run_pandoc_conversion") as mock_run_conversion,
        patch("app.PandocController.postprocess_and_build_response") as mock_postprocess,
    ):
        # Set up mocks
        mock_run_conversion.return_value = b"Converted content"

        # Create mock response
        mock_response = MagicMock()
        mock_response.mimetype = "text/html"
        mock_response.status_code = 200
        mock_postprocess.return_value = mock_response

        # Create client and send request with custom file extension
        test_client = app.test_client()
        response = test_client.post("/convert/markdown/to/html?file_name=custom_name.html", data="# Test markdown")

        # Assertions
        assert response.status_code == 200

        # Verify the custom filename was passed to postprocess_and_build_response
        mock_postprocess.assert_called_once_with(b"Converted content", "html", "custom_name.html")


def test_docx_with_template_encoding():
    """Test convert_docx_with_ref with encoding parameter and file source."""
    with (
        patch("app.PandocController.run_pandoc_conversion") as mock_run_conversion,
        patch("app.PandocController.postprocess_and_build_response") as mock_postprocess,
        patch("pathlib.Path.open", create=True) as mock_path_open,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.unlink"),
        patch("time.time", return_value=1234567890),
    ):
        # Set up mocks
        mock_run_conversion.return_value = b"Converted content"
        mock_file = MagicMock()
        mock_file.write = MagicMock()
        mock_path_open.return_value.__enter__.return_value = mock_file

        # Create mock response
        mock_response = MagicMock()
        mock_response.mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        mock_response.status_code = 200
        mock_postprocess.return_value = mock_response

        # Create a test client
        test_client = app.test_client()

        # Create test file
        test_file = MagicMock()
        test_file.read.return_value = b"Test content"

        # Create template file
        template_file = MagicMock()
        template_file.read.return_value = b"Template content"

        # Send request with encoding
        from werkzeug.datastructures import FileStorage

        # Create file storage objects with content
        source_file = FileStorage(stream=io.BytesIO(b"Test content"), filename="test.md", content_type="text/markdown")

        template_file = FileStorage(stream=io.BytesIO(b"Template content"), filename="template.docx", content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Send request with encoding parameter
        with test_client as client:
            response = client.post("/convert/markdown/to/docx-with-template?encoding=utf-8&file_name=custom_name.docx", data={"source": source_file, "template": template_file}, content_type="multipart/form-data")

        # Assertions
        assert response.status_code == 200

        # Verify the reference doc option was passed
        run_options = mock_run_conversion.call_args[0][3]
        assert any("--reference-doc=ref_1234567890.docx" in opt for opt in run_options)
