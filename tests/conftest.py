"""Shared pytest fixtures for all tests."""

import os
from unittest.mock import patch

import pytest


@pytest.fixture(autouse=True)
def disable_metrics_server():
    """Disable metrics server for all tests to avoid port conflicts."""
    with patch.dict(os.environ, {"METRICS_SERVER_ENABLED": "false"}):
        yield
