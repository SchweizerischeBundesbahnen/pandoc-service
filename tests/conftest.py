"""Shared pytest fixtures for all tests."""

import os
from unittest.mock import patch

import pytest


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
