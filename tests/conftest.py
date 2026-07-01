"""Shared pytest fixtures for all tests."""

import logging
import os
import subprocess
from unittest.mock import patch

import pytest
import requests

from tests.test_container import (
    TEST_IMAGE_FULL,
    TEST_CONTAINER_NAME,
    TestParameters,
    cleanup_docker_resources,
    wait_for_container_ready,
)


@pytest.fixture(autouse=True)
def disable_metrics_server():
    """Disable the metrics server and the lifespan Chromium startup for all tests.

    Disabling SVG conversion keeps the FastAPI lifespan from launching a Chromium
    browser during TestClient-based controller tests. The browser-based
    ChromiumManager and SvgProcessor tests construct a real ChromiumManager
    directly (mirroring weasyprint-service), so they are unaffected by this flag;
    health monitoring is left at its default (enabled) for those tests.
    """
    with patch.dict(os.environ, {"METRICS_SERVER_ENABLED": "false", "ENABLE_SVG_CONVERSION": "false"}):
        yield


@pytest.fixture(scope="session", autouse=True)
def cleanup_session():
    """Session-level fixture to ensure cleanup happens before and after all tests."""
    try:
        cleanup_docker_resources()
    except Exception as e:
        logging.error(f"Error in pre-test cleanup: {e}")

    yield

    try:
        cleanup_docker_resources()
    except Exception as e:
        logging.error(f"Error in post-test cleanup: {e}")
        raise


@pytest.fixture(scope="session")
def pandoc_container():
    """Build and start the pandoc-service container. Shared across a test module."""
    import docker

    client = docker.from_env()
    container = None

    try:
        cleanup_docker_resources()

        subprocess.run(["docker", "build", "-t", TEST_IMAGE_FULL, "."], env={**os.environ, "DOCKER_BUILDKIT": "1"}, timeout=300, check=True)

        container = client.containers.run(image=TEST_IMAGE_FULL, detach=True, name=TEST_CONTAINER_NAME, ports={"9082": 9082}, init=True)

        wait_for_container_ready(container)

        yield container

    except Exception as e:
        logging.error(f"Error in container setup: {e}")
        raise

    finally:
        try:
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

            cleanup_docker_resources()

        except Exception as e:
            logging.error(f"Error in container cleanup: {e}")


@pytest.fixture(scope="session")
def test_parameters(pandoc_container):
    """Test parameters and request session for container-based tests."""
    base_url = "http://localhost:9082"
    flush_tmp_file_enabled = False
    request_session = requests.Session()

    try:
        yield TestParameters(base_url, flush_tmp_file_enabled, request_session, pandoc_container)
    finally:
        request_session.close()
