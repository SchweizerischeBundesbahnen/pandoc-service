"""The pandoc of the image holds to PANDOC_MAX_HEAP: a document beyond it is answered 413; #271.

The other container tests run without a memory limit, so pandoc runs without a heap limit there.
This one starts a container of the same image with a small limit of its own.
"""

from __future__ import annotations

import time

import pytest
import requests

from tests.test_container import TEST_IMAGE_FULL

PORT = 9093
HEAP_LIMIT = "32m"


@pytest.fixture(scope="module")
def heap_limited_url(pandoc_container):
    """A container of the test image, which the session fixture built, with a heap limit of 32 MB."""
    import docker

    container = docker.from_env().containers.run(
        image=TEST_IMAGE_FULL,
        detach=True,
        name="pandoc-service-heap-limit-test",
        ports={"9082": PORT},
        environment={"PANDOC_MAX_HEAP": HEAP_LIMIT},
        init=True,
    )
    url = f"http://localhost:{PORT}"
    try:
        deadline = time.time() + 60
        while time.time() < deadline:
            try:
                if requests.get(f"{url}/version", timeout=2).status_code == 200:
                    break
            except requests.exceptions.RequestException:
                # The service is still starting and refuses the connection: try again
                pass
            time.sleep(1)
        else:
            pytest.fail("The heap-limited container did not become ready")
        yield url, container
    finally:
        container.remove(force=True)


def test_the_startup_log_names_the_stated_limit(heap_limited_url):
    _, container = heap_limited_url

    assert f"pandoc heap limited to {HEAP_LIMIT}" in container.logs().decode()


def test_a_document_within_the_limit_converts(heap_limited_url):
    url, _ = heap_limited_url

    response = requests.post(f"{url}/convert/html/to/docx", data=b"<p>Hello</p>", timeout=60)

    assert response.status_code == 200


def test_a_document_beyond_the_limit_is_answered_too_large(heap_limited_url):
    url, _ = heap_limited_url
    rows = "".join(f"<tr><td>Row {row}</td><td>Some text of row {row} with a few words</td></tr>" for row in range(20_000))

    response = requests.post(f"{url}/convert/html/to/docx", data=f"<table>{rows}</table>".encode(), timeout=120)

    assert response.status_code == 413
    assert f"({HEAP_LIMIT})" in response.text
    assert "PANDOC_MAX_HEAP" in response.text
