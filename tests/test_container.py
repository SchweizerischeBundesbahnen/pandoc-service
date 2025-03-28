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


def __load_test_file(file_path: str) -> str:
    with Path(file_path).open(encoding="utf-8") as file:
        file_content = file.read()
        return file_content
