"""Tests for metrics server module."""

import os
from unittest.mock import patch

import pytest
from starlette.testclient import TestClient

from app.metrics_server import (
    DEFAULT_METRICS_PORT,
    MetricsServer,
    get_metrics_port,
    is_metrics_server_enabled,
    metrics_app,
)
from app.pandoc_metrics import reset_pandoc_metrics


@pytest.fixture(autouse=True)
def reset_metrics():
    """Reset metrics before each test."""
    reset_pandoc_metrics()
    yield


class TestMetricsEndpoint:
    """Tests for /metrics endpoint."""

    def test_metrics_endpoint_returns_prometheus_format(self):
        """Test that metrics endpoint returns Prometheus format."""
        client = TestClient(metrics_app)
        response = client.get("/metrics")

        assert response.status_code == 200
        assert "text/plain" in response.headers["content-type"] or "text/plain" in response.headers.get("content-type", "")

        # Check for expected metric names in response
        content = response.text
        assert "pandoc_conversions_total" in content or "uptime_seconds" in content

    def test_metrics_endpoint_includes_uptime(self):
        """Test that metrics endpoint includes uptime."""
        client = TestClient(metrics_app)
        response = client.get("/metrics")

        assert response.status_code == 200
        assert "uptime_seconds" in response.text


class TestGetMetricsPort:
    """Tests for get_metrics_port function."""

    def test_default_port(self):
        """Test default metrics port."""
        with patch.dict(os.environ, {}, clear=True):
            # Remove METRICS_PORT if present
            os.environ.pop("METRICS_PORT", None)
            port = get_metrics_port()
            assert port == DEFAULT_METRICS_PORT

    def test_custom_port_from_env(self):
        """Test custom metrics port from environment."""
        with patch.dict(os.environ, {"METRICS_PORT": "9999"}):
            port = get_metrics_port()
            assert port == 9999

    def test_invalid_port_falls_back_to_default(self):
        """Test that invalid port falls back to default."""
        with patch.dict(os.environ, {"METRICS_PORT": "not_a_number"}):
            port = get_metrics_port()
            assert port == DEFAULT_METRICS_PORT

    def test_port_below_minimum(self):
        """Test that port below minimum falls back to default."""
        with patch.dict(os.environ, {"METRICS_PORT": "80"}):
            port = get_metrics_port()
            assert port == DEFAULT_METRICS_PORT

    def test_port_above_maximum(self):
        """Test that port above maximum falls back to default."""
        with patch.dict(os.environ, {"METRICS_PORT": "70000"}):
            port = get_metrics_port()
            assert port == DEFAULT_METRICS_PORT


class TestIsMetricsServerEnabled:
    """Tests for is_metrics_server_enabled function."""

    def test_enabled_by_default(self):
        """Test that metrics server is enabled by default."""
        with patch.dict(os.environ, {}, clear=True):
            os.environ.pop("METRICS_SERVER_ENABLED", None)
            assert is_metrics_server_enabled() is True

    def test_explicitly_enabled(self):
        """Test explicitly enabling metrics server."""
        for value in ["true", "True", "TRUE", "1", "yes", "on"]:
            with patch.dict(os.environ, {"METRICS_SERVER_ENABLED": value}):
                assert is_metrics_server_enabled() is True

    def test_explicitly_disabled(self):
        """Test explicitly disabling metrics server."""
        for value in ["false", "False", "FALSE", "0", "no", "off"]:
            with patch.dict(os.environ, {"METRICS_SERVER_ENABLED": value}):
                assert is_metrics_server_enabled() is False


class TestMetricsServer:
    """Tests for MetricsServer class."""

    def test_server_initialization(self):
        """Test MetricsServer initialization."""
        server = MetricsServer(port=9999)
        assert server.port == 9999
        assert server.is_running is False

    def test_server_default_port(self):
        """Test MetricsServer with default port."""
        server = MetricsServer()
        assert server.port == DEFAULT_METRICS_PORT

    def test_server_start_stop(self):
        """Test starting and stopping the metrics server."""
        import anyio

        async def run_test():
            server = MetricsServer(port=19182)  # Use a different port to avoid conflicts

            assert server.is_running is False

            await server.start()
            assert server.is_running is True

            await server.stop()
            assert server.is_running is False

        anyio.run(run_test)

    def test_server_double_start(self):
        """Test that double start is handled gracefully."""
        import anyio

        async def run_test():
            server = MetricsServer(port=19183)

            await server.start()
            # Second start should be a no-op
            await server.start()
            assert server.is_running is True

            await server.stop()

        anyio.run(run_test)

    def test_server_stop_without_start(self):
        """Test that stopping without starting is handled gracefully."""
        import anyio

        async def run_test():
            server = MetricsServer(port=19184)

            # Should not raise
            await server.stop()
            assert server.is_running is False

        anyio.run(run_test)
