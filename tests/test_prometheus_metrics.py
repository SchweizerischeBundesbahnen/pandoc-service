"""Tests for prometheus metrics module."""

import os
from unittest.mock import MagicMock, patch

import pytest

from app.pandoc_metrics import PandocMetrics, get_pandoc_metrics, reset_pandoc_metrics
from app.prometheus_metrics import (
    increment_conversion_failure,
    increment_conversion_success,
    increment_template_conversion,
    observe_post_processing_duration,
    observe_request_body_size,
    observe_response_body_size,
    observe_subprocess_duration,
    update_gauges_from_pandoc_metrics,
)


@pytest.fixture(autouse=True)
def reset_metrics():
    """Reset metrics before each test."""
    reset_pandoc_metrics()
    yield


class TestPandocMetrics:
    """Tests for PandocMetrics class."""

    def test_initial_state(self):
        """Test initial metrics state."""
        metrics = PandocMetrics()
        assert metrics.total_conversions == 0
        assert metrics.failed_conversions == 0
        assert metrics.active_conversions == 0
        assert metrics.avg_conversion_time_ms == 0.0

    def test_record_conversion_start(self):
        """Test recording conversion start."""
        metrics = PandocMetrics()
        metrics.record_conversion_start()
        assert metrics.active_conversions == 1

        metrics.record_conversion_start()
        assert metrics.active_conversions == 2

    def test_record_conversion_success(self):
        """Test recording successful conversion."""
        metrics = PandocMetrics()
        metrics.record_conversion_start()
        metrics.record_conversion_success(100.0)

        assert metrics.total_conversions == 1
        assert metrics.active_conversions == 0
        assert metrics.avg_conversion_time_ms == 100.0

    def test_record_conversion_failure(self):
        """Test recording failed conversion."""
        metrics = PandocMetrics()
        metrics.record_conversion_start()
        metrics.record_conversion_failure()

        assert metrics.failed_conversions == 1
        assert metrics.active_conversions == 0

    def test_error_rate_calculation(self):
        """Test error rate calculation."""
        metrics = PandocMetrics()

        # No conversions yet
        assert metrics.get_error_rate() == 0.0

        # 1 success, 1 failure = 50% error rate
        metrics.record_conversion_success(100.0)
        metrics.record_conversion_failure()
        assert metrics.get_error_rate() == 50.0

    def test_average_conversion_time(self):
        """Test average conversion time calculation."""
        metrics = PandocMetrics()

        metrics.record_conversion_success(100.0)
        assert metrics.avg_conversion_time_ms == 100.0

        metrics.record_conversion_success(200.0)
        assert metrics.avg_conversion_time_ms == 150.0

    def test_uptime_seconds(self):
        """Test uptime calculation."""
        metrics = PandocMetrics()
        uptime = metrics.get_uptime_seconds()
        assert uptime >= 0.0

    def test_get_metrics_dict(self):
        """Test getting metrics as dictionary."""
        metrics = PandocMetrics()
        metrics.record_conversion_success(100.0)
        metrics.set_pandoc_version("3.1.9")

        result = metrics.get_metrics()

        assert result["total_conversions"] == 1
        assert result["failed_conversions"] == 0
        assert result["error_rate_percent"] == 0.0
        assert result["avg_conversion_time_ms"] == 100.0
        assert result["active_conversions"] == 0
        assert result["pandoc_version"] == "3.1.9"
        assert "uptime_seconds" in result

    def test_thread_safety(self):
        """Test thread safety of metrics operations."""
        import threading

        metrics = PandocMetrics()
        threads = []

        def record_conversions():
            for _ in range(100):
                metrics.record_conversion_start()
                metrics.record_conversion_success(10.0)

        # Start multiple threads
        for _ in range(10):
            t = threading.Thread(target=record_conversions)
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        # All conversions should be recorded
        assert metrics.total_conversions == 1000
        assert metrics.active_conversions == 0


class TestPrometheusMetricsHelpers:
    """Tests for prometheus metrics helper functions."""

    def test_increment_conversion_success(self):
        """Test incrementing successful conversion counter."""
        # This should not raise
        increment_conversion_success("markdown", "docx", 1.5)

    def test_increment_conversion_failure(self):
        """Test incrementing failed conversion counter."""
        increment_conversion_failure("markdown", "pdf")

    def test_increment_template_conversion(self):
        """Test incrementing template conversion counter."""
        increment_template_conversion("docx")

    def test_observe_subprocess_duration(self):
        """Test observing subprocess duration."""
        observe_subprocess_duration(0.5)

    def test_observe_post_processing_duration(self):
        """Test observing post-processing duration."""
        observe_post_processing_duration("docx", 0.1)

    def test_observe_request_body_size(self):
        """Test observing request body size."""
        observe_request_body_size(1024)

    def test_observe_response_body_size(self):
        """Test observing response body size."""
        observe_response_body_size(2048)


class TestUpdateGauges:
    """Tests for gauge update function."""

    def test_update_gauges_from_pandoc_metrics(self):
        """Test updating gauges from PandocMetrics."""
        metrics = PandocMetrics()
        metrics.record_conversion_success(150.0)
        metrics.set_pandoc_version("3.1.9")

        with patch.dict(os.environ, {"PANDOC_SERVICE_VERSION": "1.0.0"}):
            # Should not raise
            update_gauges_from_pandoc_metrics(metrics)

    def test_update_gauges_handles_errors(self):
        """Test that gauge update handles errors gracefully."""
        # Create a mock that raises an exception
        mock_metrics = MagicMock()
        mock_metrics.get_metrics.side_effect = RuntimeError("Test error")

        # Should not raise, just log the error
        update_gauges_from_pandoc_metrics(mock_metrics)


class TestGetPandocMetrics:
    """Tests for global metrics singleton."""

    def test_get_pandoc_metrics_singleton(self):
        """Test that get_pandoc_metrics returns singleton."""
        metrics1 = get_pandoc_metrics()
        metrics2 = get_pandoc_metrics()
        assert metrics1 is metrics2

    def test_reset_pandoc_metrics(self):
        """Test resetting global metrics."""
        metrics1 = get_pandoc_metrics()
        metrics1.record_conversion_success(100.0)

        reset_pandoc_metrics()

        metrics2 = get_pandoc_metrics()
        assert metrics2.total_conversions == 0
        assert metrics1 is not metrics2
