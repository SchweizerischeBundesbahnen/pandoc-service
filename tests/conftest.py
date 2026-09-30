"""Shared pytest fixtures for all tests."""

import logging
import os
import subprocess
from unittest.mock import patch

import pytest
import requests

from tests.test_container import (
    TEST_CONTAINER_NAME,
    TEST_IMAGE_FULL,
    TestParameters,
    cleanup_docker_resources,
    wait_for_container_ready,
)

logger = logging.getLogger(__name__)

# The image carries a Debian base, pandoc, tectonic and Playwright's Chromium. On a warm cache it is
# built in seconds; on a cold one, which every CI runner has and which a local `docker builder prune`
# leaves behind, it has taken past five minutes and timed out, and every container test then errors
# at setup with nothing about the build in the message. The timeout is here to stop a hung build, not
# to measure a cold one.
BUILD_TIMEOUT_SECONDS = 1800


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
    # Pre-test cleanup is best effort.
    except Exception as e:  # noqa: BLE001
        logger.error(f"Error in pre-test cleanup: {e}")

    yield

    try:
        cleanup_docker_resources()
    except Exception as e:
        logger.error(f"Error in post-test cleanup: {e}")
        raise


@pytest.fixture(scope="session")
def pandoc_container():
    """Build and start the pandoc-service container. Shared across a test module."""
    import docker

    client = docker.from_env()
    container = None

    try:
        cleanup_docker_resources()

        subprocess.run(["docker", "build", "-t", TEST_IMAGE_FULL, "."], env={**os.environ, "DOCKER_BUILDKIT": "1"}, timeout=BUILD_TIMEOUT_SECONDS, check=True)

        container = client.containers.run(image=TEST_IMAGE_FULL, detach=True, name=TEST_CONTAINER_NAME, ports={"9082": 9082}, init=True)

        wait_for_container_ready(container)

        yield container

    except Exception as e:
        logger.error(f"Error in container setup: {e}")
        raise

    finally:
        try:
            if container:
                logger.info("Cleaning up test container...")
                try:
                    container.stop(timeout=1)
                except docker.errors.APIError as e:
                    logger.warning(f"Could not stop container: {e}")

                try:
                    container.remove(force=True)
                except docker.errors.APIError as e:
                    logger.error(f"Could not remove container: {e}")

            cleanup_docker_resources()

        # Container cleanup is best effort.
        except Exception as e:  # noqa: BLE001
            logger.error(f"Error in container cleanup: {e}")


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
