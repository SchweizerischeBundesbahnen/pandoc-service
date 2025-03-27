import time
from typing import NamedTuple

import docker
import pytest
import requests
from docker.models.containers import Container


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


def test_container_no_error_logs(test_parameters: TestParameters) -> None:
    logs = test_parameters.container.logs()

    assert logs == b"INFO:root:Pandoc service listening port: 9082\n"
