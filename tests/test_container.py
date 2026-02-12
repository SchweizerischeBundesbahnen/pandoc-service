import io
import logging
import time
from pathlib import Path
from typing import NamedTuple
import subprocess
import os

import docker
import pytest
import requests
from docker.models.containers import Container
from docx import Document
from docx.shared import RGBColor
from pptx import Presentation

# Constants for Docker resources
TEST_IMAGE_NAME = "pandoc-service-test"
TEST_IMAGE_TAG = "latest"
TEST_CONTAINER_NAME = "pandoc-service-test-container"
TEST_IMAGE_FULL = f"{TEST_IMAGE_NAME}:{TEST_IMAGE_TAG}"

SOURCE_HTML = """
            <html>
                <body>
                    <h1>Simple html with an ordered list</h1>
                    <ol>
                        <li>First</li>
                        <li>Second</li>
                        <li>Third</li>
                    </ol>
                    <p>Some <b>bold German vowels ä, ö, and ü</b> at the bottom.</p>
                </body>
            </html>
            """

SOURCE_HTML_WITH_HEADINGS = """
            <html>
                <head>
                    <title>Test doc title</title>
                </head>
                <body>
                    <h1>Simple html with several headings</h1>
                    <p>Some content 1</p>
                    <h2>Second heading with German vowels ä, ö, and ü</h2>
                    <p>Some content 2</p>
                    <h3>Third</h3>
                    <p>Some content 3</p>
                </body>
            </html>
        """

SOURCE_MARKDOWN_FOR_PPTX = """
# Slide 1 Title

First slide content with bullet points:

- Item 1
- Item 2
- Item 3

---

# Slide 2 Title

Second slide with more content.

Some **bold** and *italic* text.
"""


class TestParameters(NamedTuple):
    base_url: str
    flush_tmp_file_enabled: bool
    request_session: requests.Session
    container: Container
    __test__ = False


def _stop_and_remove_container(container: Container) -> None:
    """Helper function to stop and remove a single container."""
    try:
        logging.info(f"Stopping container: {container.name}")
        container.stop(timeout=1)
    except docker.errors.APIError as e:
        logging.warning(f"Could not stop container {container.name}: {e}")

    try:
        logging.info(f"Removing container: {container.name}")
        container.remove(force=True)
    except docker.errors.APIError as e:
        logging.error(f"Error removing container {container.name}: {e}")


def _remove_image(image) -> None:
    """Helper function to remove a single image."""
    try:
        logging.info(f"Removing image: {image.tags}")
        image.remove(force=True)
    except docker.errors.APIError as e:
        logging.error(f"Error removing image {image.tags}: {e}")


def _is_test_related_container(container: Container) -> bool:
    """Check if a container is related to our tests."""
    return (
        container.name == TEST_CONTAINER_NAME
        or (container.image.tags and TEST_IMAGE_FULL in str(container.image.tags))
        or not container.image.tags  # Intermediate containers
        or (container.image.tags and "python:3.13-alpine" in str(container.image.tags))  # Base image containers
    )


def _is_test_related_image(image) -> bool:
    """Check if an image is related to our tests."""
    return (image.tags and TEST_IMAGE_FULL in str(image.tags)) or not image.tags


def _cleanup_containers(client: docker.DockerClient) -> None:
    """Clean up test-related containers."""
    try:
        containers = client.containers.list(all=True)
        for container in containers:
            if _is_test_related_container(container):
                _stop_and_remove_container(container)
    except docker.errors.APIError as e:
        logging.error(f"Error listing containers: {e}")


def _cleanup_images(client: docker.DockerClient) -> None:
    """Clean up test-related images."""
    try:
        images = client.images.list(all=True)
        for image in images:
            if _is_test_related_image(image):
                _remove_image(image)
    except docker.errors.APIError as e:
        logging.error(f"Error listing images: {e}")


def _verify_containers(client: docker.DockerClient) -> None:
    """Verify and clean up any remaining test-related containers."""
    try:
        remaining = client.containers.list(all=True)
        remaining_test = [c for c in remaining if _is_test_related_container(c)]

        if remaining_test:
            logging.warning(f"Found {len(remaining_test)} test-related containers still remaining after cleanup")
            for container in remaining_test:
                logging.warning(f"Remaining container: {container.name} ({container.id})")
                _stop_and_remove_container(container)
    except Exception as e:
        logging.error(f"Error in container verification: {e}")


def _verify_images(client: docker.DockerClient) -> None:
    """Verify and clean up any remaining test-related images."""
    try:
        remaining_images = client.images.list(all=True)
        remaining_test_images = [i for i in remaining_images if _is_test_related_image(i)]

        if remaining_test_images:
            logging.warning(f"Found {len(remaining_test_images)} test-related images still remaining after cleanup")
            for image in remaining_test_images:
                logging.warning(f"Remaining image: {image.id} (tags: {image.tags})")
                _remove_image(image)
    except Exception as e:
        logging.error(f"Error in image verification: {e}")


def cleanup_docker_resources():
    """
    Cleanup function to remove any leftover test containers and images.
    Ensures thorough cleanup of all test-related Docker resources.
    """
    client = docker.from_env()

    # Initial cleanup
    _cleanup_containers(client)
    _cleanup_images(client)

    # Final verification
    _verify_containers(client)
    _verify_images(client)


def wait_for_container_ready(container: Container, max_wait_time: int = 60) -> None:
    """
    Wait for container to become ready by checking the /version endpoint.

    Args:
        container: Docker container to wait for
        max_wait_time: Maximum time to wait in seconds (default: 60)

    Raises:
        TimeoutError: If container does not become ready within max_wait_time
    """
    start_time = time.time()
    base_url = "http://localhost:9082"

    while time.time() - start_time < max_wait_time:
        try:
            response = requests.get(f"{base_url}/version", timeout=2)
            if response.status_code == 200:
                logging.info("Container is ready")
                return
        except requests.exceptions.RequestException as e:
            logging.debug(f"Container not ready yet, retrying: {e}")
        time.sleep(1)

    # Timeout reached, print logs for debugging
    logs = container.logs().decode("utf-8")
    raise TimeoutError(f"Container did not become ready within {max_wait_time} seconds. Logs:\n{logs}")


@pytest.fixture(scope="session", autouse=True)
def cleanup_session():
    """
    Session-level fixture to ensure cleanup happens before and after all tests.
    """
    # Clean up any leftover resources from previous runs
    try:
        cleanup_docker_resources()
    except Exception as e:
        logging.error(f"Error in pre-test cleanup: {e}")

    yield

    # Clean up after all tests are done, even if tests fail
    try:
        cleanup_docker_resources()
    except Exception as e:
        logging.error(f"Error in post-test cleanup: {e}")
        raise


@pytest.fixture(scope="module")
def pandoc_container():
    """
    Setup function for building and starting the pandoc-service image.
    Runs once per module and is cleaned up after execution

    Yields:
        Container: Built docker container
    """
    client = docker.from_env()
    container = None

    try:
        # Clean up any existing resources first
        cleanup_docker_resources()

        # Build the image with test-specific tag
        subprocess.run(
            ["docker", "build", "-t", TEST_IMAGE_FULL, "."],
            env={**os.environ, "DOCKER_BUILDKIT": "1"},
            timeout=300,
            check=True
        )

        # Run the container with test-specific name
        container = client.containers.run(image=TEST_IMAGE_FULL, detach=True, name=TEST_CONTAINER_NAME, ports={"9082": 9082}, init=True)

        # Wait for container to be ready using health check
        wait_for_container_ready(container)

        yield container

    except Exception as e:
        logging.error(f"Error in container setup: {e}")
        raise

    finally:
        try:
            # Ensure cleanup happens even if tests fail
            if container:
                logging.info("Cleaning up test container...")
                try:
                    container.stop(timeout=1)
                except docker.errors.APIError as e:
                    logging.warning(f"Could not stop container: {e}")

                try:
                    container.remove(force=True)
                except docker.errors.APIError as e:
                    logging.error(f"Could not remove container: {e}")

            # Final cleanup of any remaining resources
            cleanup_docker_resources()

        except Exception as e:
            logging.error(f"Error in container cleanup: {e}")


@pytest.fixture(scope="module")
def test_parameters(pandoc_container: Container):
    """
    Setup function for test parameters and request session.
    Runs once per module and is cleaned up after execution.

    Args:
        pandoc_container (Container): pandoc-service docker container

    Yields:
        TestParameters: The setup test parameters
    """
    base_url = "http://localhost:9082"
    flush_tmp_file_enabled = False
    request_session = requests.Session()

    try:
        yield TestParameters(base_url, flush_tmp_file_enabled, request_session, pandoc_container)
    finally:
        request_session.close()


def test_container_logs(test_parameters: TestParameters) -> None:
    logs = test_parameters.container.logs()

    assert b"Pandoc service listening port: 9082\n" in logs


def test_convert_html_to_md(test_parameters: TestParameters) -> None:
    expected_content = __load_test_file("tests/data/expected-html-to-md.md")
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="markdown", data=SOURCE_HTML)
    assert response.status_code == 200
    assert response.content.decode("utf-8") == expected_content


def test_convert_html_to_textile(test_parameters: TestParameters) -> None:
    expected_content = __load_test_file("tests/data/expected-html-to-textile.textile")
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="textile", data=SOURCE_HTML)
    assert response.status_code == 200
    assert response.content.decode("utf-8") == expected_content


def test_convert_html_to_plain(test_parameters: TestParameters) -> None:
    expected_content = __load_test_file("tests/data/expected-html-to-txt.txt")
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="plain", data=SOURCE_HTML)
    assert response.status_code == 200
    assert response.content.decode("utf-8") == expected_content


def test_convert_docx_to_plain(test_parameters: TestParameters) -> None:
    with Path("tests/data/test-input.docx").open("rb") as source_file:
        expected_content = __load_test_file("tests/data/expected-docx-to-txt.txt")
        data = ("test-input.docx", source_file.read(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="docx", target_format="plain", data=data)
        assert response.status_code == 200
        assert response.content.decode("utf-8") == expected_content


def test_convert_html_to_docx(test_parameters: TestParameters) -> None:
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="docx", data=SOURCE_HTML)
    assert response.status_code == 200

    document = Document(io.BytesIO(response.content))

    paragraphs = []
    for paragraph in document.paragraphs:
        paragraphs.append(paragraph.text)

    expected_paragraphs = [
        "Simple html with an ordered list",
        "First",
        "Second",
        "Third",
        "Some bold German vowels ä, ö, and ü at the bottom.",
    ]

    assert expected_paragraphs == paragraphs


def test_convert_with_docx_template(test_parameters: TestParameters) -> None:
    # First test without template - it has some default headings color
    response = __send_docx_with_template_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, data=SOURCE_HTML_WITH_HEADINGS, source_format="html")
    __assert_doc_contains_specific_headers_color(RGBColor(15, 71, 97), response.content)

    # Now test with 'RED' template - it forces red color for headings
    with Path("tests/data/template-red.docx").open("rb") as t:
        template = t.read()
    response = __send_docx_with_template_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, data=SOURCE_HTML_WITH_HEADINGS, source_format="html", template=template)
    __assert_doc_contains_specific_headers_color(RGBColor(255, 0, 0), response.content)


def test_version_endpoint(test_parameters: TestParameters) -> None:
    """Test that the /version endpoint returns the expected information."""
    url = f"{test_parameters.base_url}/version"
    response = test_parameters.request_session.get(url)

    # Verify response status
    assert response.status_code == 200

    # Parse response as JSON
    version_info = response.json()

    # Verify all expected fields are present
    assert "python" in version_info
    assert "pandoc" in version_info
    assert "pandocService" in version_info
    assert "timestamp" in version_info

    # Verify that values are reasonable (not empty where required)
    assert version_info["python"], "Python version should not be empty"
    assert version_info["pandoc"], "Pandoc version should not be empty"
    assert version_info["pandocService"], "Pandoc service version should not be empty"


def __send_request(base_url: str, request_session: requests.Session, source_format: str, target_format: str, data) -> requests.Response:
    url = f"{base_url}/convert/{source_format}/to/{target_format}"
    files = None
    payload = None

    if isinstance(data, tuple):
        filename, file_content, content_type = data
        files = {"source": (filename, file_content, content_type)}
    else:
        payload = data
    try:
        response = request_session.request(method="POST", url=url, data=payload, files=files, verify=True)
        if response.status_code // 100 != 2:
            logging.error(f"Error: Unexpected response: '{response}'")
            logging.error(f"Error: Response content: '{response.content}'")
        return response
    except requests.exceptions.RequestException as e:
        logging.error(f"Error: {e}")
        raise


def __send_docx_with_template_request(base_url: str, request_session: requests.Session, source_format: str, data, template=None) -> requests.Response:
    url = f"{base_url}/convert/{source_format}/to/docx-with-template"
    files = {"source": ("file.html", data)}
    if template:
        files["template"] = ("template.docx", template)
    try:
        response = request_session.request(method="POST", url=url, files=files, verify=True)
        if response.status_code // 100 != 2:
            logging.error(f"Error: Unexpected response: '{response}'")
            logging.error(f"Error: Response content: '{response.content}'")
        return response
    except requests.exceptions.RequestException as e:
        logging.error(f"Error: {e}")
        raise


def __assert_doc_contains_specific_headers_color(color, doc_content):
    document = Document(io.BytesIO(doc_content))

    # Check for specific headings colors and extract their text
    headings = []
    for paragraph in document.paragraphs:
        if paragraph.style.style_id.startswith("Heading"):
            assert color in {paragraph.style.base_style.font.color.rgb, paragraph.style.font.color.rgb}
            headings.append(paragraph.text.replace("\xa0", " "))

    expected_headings = [
        "Simple html with several headings",
        "Second heading with German vowels ä, ö, and ü",
        "Third",
    ]
    assert expected_headings == headings


def __load_test_file(file_path: str) -> str:
    with Path(file_path).open(encoding="utf-8") as file:
        file_content = file.read()
        return file_content


def __send_pptx_with_template_request(base_url: str, request_session: requests.Session, source_format: str, data, template=None) -> requests.Response:
    url = f"{base_url}/convert/{source_format}/to/pptx-with-template"
    files = {"source": ("file.md", data)}
    if template:
        files["template"] = ("template.pptx", template)
    try:
        response = request_session.request(method="POST", url=url, files=files, verify=True)
        if response.status_code // 100 != 2:
            logging.error(f"Error: Unexpected response: '{response}'")
            logging.error(f"Error: Response content: '{response.content}'")
        return response
    except requests.exceptions.RequestException as e:
        logging.error(f"Error: {e}")
        raise


def test_container_no_error_logs(test_parameters: TestParameters) -> None:
    """Verify container logs contain expected startup messages and no errors."""
    logs = test_parameters.container.logs().decode("utf-8")
    log_lines = logs.splitlines()

    # Check for critical errors (should not contain ERROR or CRITICAL level messages)
    errors = [line for line in log_lines if " - ERROR - " in line or " - CRITICAL - " in line]
    assert not errors, f"Found error logs: {errors}"

    # Check for expected startup message
    assert any("Pandoc service listening port: 9082" in line for line in log_lines), "Expected startup message not found in logs"


def test_docx_template_endpoint(test_parameters: TestParameters) -> None:
    """Test that the /docx-template endpoint returns a valid DOCX file."""
    url = f"{test_parameters.base_url}/docx-template"
    response = test_parameters.request_session.get(url)

    assert response.status_code == 200
    assert response.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    assert response.headers.get("Content-Disposition") == "attachment; filename=reference.docx"

    # Verify it's a valid DOCX by loading it
    document = Document(io.BytesIO(response.content))
    assert document is not None


def test_pptx_template_endpoint(test_parameters: TestParameters) -> None:
    """Test that the /pptx-template endpoint returns a valid PPTX file."""
    url = f"{test_parameters.base_url}/pptx-template"
    response = test_parameters.request_session.get(url)

    assert response.status_code == 200
    assert response.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    assert response.headers.get("Content-Disposition") == "attachment; filename=reference.pptx"

    # Verify it's a valid PPTX by loading it
    presentation = Presentation(io.BytesIO(response.content))
    assert presentation is not None


def test_convert_markdown_to_pptx(test_parameters: TestParameters) -> None:
    """Test converting markdown to PPTX format."""
    response = __send_request(
        base_url=test_parameters.base_url,
        request_session=test_parameters.request_session,
        source_format="markdown",
        target_format="pptx",
        data=SOURCE_MARKDOWN_FOR_PPTX,
    )
    assert response.status_code == 200

    # Verify it's a valid PPTX with expected slide count
    presentation = Presentation(io.BytesIO(response.content))
    assert presentation is not None
    assert len(presentation.slides) == 2, "Expected 2 slides from markdown with slide separator"

    # Check response headers
    assert "Pandoc-Version" in response.headers
    assert "Python-Version" in response.headers


def test_convert_with_pptx_template(test_parameters: TestParameters) -> None:
    """Test PPTX conversion with custom template."""
    # First, get the default template from the endpoint
    template_url = f"{test_parameters.base_url}/pptx-template"
    template_response = test_parameters.request_session.get(template_url)
    assert template_response.status_code == 200
    template = template_response.content

    # Test with template
    response = __send_pptx_with_template_request(
        base_url=test_parameters.base_url,
        request_session=test_parameters.request_session,
        source_format="markdown",
        data=SOURCE_MARKDOWN_FOR_PPTX,
        template=template,
    )
    assert response.status_code == 200

    # Verify it's a valid PPTX
    presentation = Presentation(io.BytesIO(response.content))
    assert presentation is not None
    assert len(presentation.slides) >= 1


def test_convert_invalid_source_format(test_parameters: TestParameters) -> None:
    """Test that invalid source format returns proper error."""
    response = __send_request(
        base_url=test_parameters.base_url,
        request_session=test_parameters.request_session,
        source_format="invalid_format",
        target_format="html",
        data="test content",
    )
    assert response.status_code == 400


def test_convert_invalid_target_format(test_parameters: TestParameters) -> None:
    """Test that invalid target format returns proper error."""
    response = __send_request(
        base_url=test_parameters.base_url,
        request_session=test_parameters.request_session,
        source_format="markdown",
        target_format="invalid_format",
        data="test content",
    )
    assert response.status_code == 400
