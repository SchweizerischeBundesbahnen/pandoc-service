import io
import logging
import time
from pathlib import Path
from typing import NamedTuple

import docker
import pytest
import requests
from docker.models.containers import Container
from docx import Document
from docx.shared import RGBColor

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


class TestParameters(NamedTuple):
    base_url: str
    flush_tmp_file_enabled: bool
    request_session: requests.Session
    container: Container
    __test__ = False


@pytest.fixture(scope="module")
def pandoc_container():
    """
    Setup function for building and starting the pandoc-service image.
    Runs once per module and is cleaned up after execution

    Yields:
        Container: Built docker container
    """
    client = docker.from_env()
    image, _ = client.images.build(path=".", tag="pandoc_service", buildargs={"APP_IMAGE_VERSION": "1.0.0"})
    container = client.containers.run(image=image, detach=True, name="pandoc_service", ports={"9082": 9082})

    time.sleep(5)

    yield container

    container.stop()
    container.remove()


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
    yield TestParameters(base_url, flush_tmp_file_enabled, request_session, pandoc_container)
    request_session.close()


def test_container_logs(test_parameters: TestParameters) -> None:
    logs = test_parameters.container.logs()

    assert logs == b"INFO:root:Pandoc service listening port: 9082\n"


def test_convert_html_to_md(test_parameters: TestParameters) -> None:
    expected_content = __load_test_file("test-data/expected-html-to-md.md")
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="markdown", data=SOURCE_HTML)
    assert response.status_code == 200
    assert response.content.decode("utf-8") == expected_content


def test_convert_html_to_textile(test_parameters: TestParameters) -> None:
    expected_content = __load_test_file("test-data/expected-html-to-textile.textile")
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="textile", data=SOURCE_HTML)
    assert response.status_code == 200
    assert response.content.decode("utf-8") == expected_content


def test_convert_html_to_plain(test_parameters: TestParameters) -> None:
    expected_content = __load_test_file("test-data/expected-html-to-txt.txt")
    response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="html", target_format="plain", data=SOURCE_HTML)
    assert response.status_code == 200
    assert response.content.decode("utf-8") == expected_content


def test_convert_docx_to_plain(test_parameters: TestParameters) -> None:
    with Path("test-data/test-input.docx").open("rb") as source_file:
        expected_content = __load_test_file("test-data/expected-docx-to-txt.txt")
        response = __send_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, source_format="docx", target_format="plain", data=source_file.read())
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
    with Path("test-data/template-red.docx").open("rb") as t:
        template = t.read()
    response = __send_docx_with_template_request(base_url=test_parameters.base_url, request_session=test_parameters.request_session, data=SOURCE_HTML_WITH_HEADINGS, source_format="html", template=template)
    __assert_doc_contains_specific_headers_color(RGBColor(255, 0, 0), response.content)


def __send_request(base_url: str, request_session: requests.Session, source_format: str, target_format: str, data) -> requests.Response:
    url = f"{base_url}/convert/{source_format}/to/{target_format}"
    try:
        response = request_session.request(method="POST", url=url, data=data, verify=True)
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
