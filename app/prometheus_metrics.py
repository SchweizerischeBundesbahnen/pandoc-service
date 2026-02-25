"""
Prometheus metrics collectors for pandoc-service.

This module defines custom Prometheus metrics that expose conversion
and application-level metrics for monitoring and observability.

Note: Counters are incremented when events occur (not synced from external state).
      Gauges are updated periodically to reflect current state.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from prometheus_client import Counter, Gauge, Histogram, Info

if TYPE_CHECKING:
    from app.pandoc_metrics import PandocMetrics


logger = logging.getLogger(__name__)


# Conversion counters - labeled by source and target format
pandoc_conversions_total = Counter(
    "pandoc_conversions_total",
    "Total number of successful document conversions",
    ["source_format", "target_format"],
)

pandoc_conversion_failures_total = Counter(
    "pandoc_conversion_failures_total",
    "Total number of failed document conversions",
    ["source_format", "target_format"],
)

# Template conversion counters
pandoc_template_conversions_total = Counter(
    "pandoc_template_conversions_total",
    "Total number of conversions using custom templates",
    ["target_format"],
)

# Conversion duration histograms
pandoc_conversion_duration_seconds = Histogram(
    "pandoc_conversion_duration_seconds",
    "Document conversion duration in seconds",
    ["source_format", "target_format"],
    buckets=[0.1, 0.25, 0.5, 1.0, 2.5, 5.0, 10.0, 30.0, 60.0],
)

pandoc_subprocess_duration_seconds = Histogram(
    "pandoc_subprocess_duration_seconds",
    "Time spent in pandoc subprocess in seconds",
    buckets=[0.1, 0.25, 0.5, 1.0, 2.5, 5.0, 10.0, 30.0],
)

pandoc_post_processing_duration_seconds = Histogram(
    "pandoc_post_processing_duration_seconds",
    "Time spent in DOCX/PPTX post-processing in seconds",
    ["target_format"],
    buckets=[0.01, 0.05, 0.1, 0.25, 0.5, 1.0, 2.5, 5.0],
)

# Request/response size histograms
pandoc_request_body_bytes = Histogram(
    "pandoc_request_body_bytes",
    "Input document size in bytes",
    buckets=[1024, 10240, 102400, 1048576, 10485760, 104857600],  # 1KB, 10KB, 100KB, 1MB, 10MB, 100MB
)

pandoc_response_body_bytes = Histogram(
    "pandoc_response_body_bytes",
    "Output document size in bytes",
    buckets=[1024, 10240, 102400, 1048576, 10485760, 104857600],  # 1KB, 10KB, 100KB, 1MB, 10MB, 100MB
)

# Error rate gauges
pandoc_conversion_error_rate_percent = Gauge(
    "pandoc_conversion_error_rate_percent",
    "Document conversion error rate as percentage",
)

avg_pandoc_conversion_time_seconds = Gauge(
    "avg_pandoc_conversion_time_seconds",
    "Average document conversion time in seconds",
)

# Service lifecycle metrics
uptime_seconds = Gauge(
    "uptime_seconds",
    "Service uptime in seconds",
)

active_conversions = Gauge(
    "active_conversions",
    "Current number of active document conversions",
)

# Service info
pandoc_info = Info(
    "pandoc",
    "Pandoc service information",
)


# Helper functions to increment counters (called when events occur)
def increment_conversion_success(source_format: str, target_format: str, duration_seconds: float) -> None:
    """Increment successful conversion counter and record duration."""
    pandoc_conversions_total.labels(source_format=source_format, target_format=target_format).inc()
    pandoc_conversion_duration_seconds.labels(source_format=source_format, target_format=target_format).observe(duration_seconds)


def increment_conversion_failure(source_format: str, target_format: str) -> None:
    """Increment failed conversion counter."""
    pandoc_conversion_failures_total.labels(source_format=source_format, target_format=target_format).inc()


def increment_template_conversion(target_format: str) -> None:
    """Increment template conversion counter."""
    pandoc_template_conversions_total.labels(target_format=target_format).inc()


def observe_subprocess_duration(duration_seconds: float) -> None:
    """Record pandoc subprocess duration."""
    pandoc_subprocess_duration_seconds.observe(duration_seconds)


def observe_post_processing_duration(target_format: str, duration_seconds: float) -> None:
    """Record post-processing duration."""
    pandoc_post_processing_duration_seconds.labels(target_format=target_format).observe(duration_seconds)


def observe_request_body_size(size_bytes: int) -> None:
    """Record input document size."""
    pandoc_request_body_bytes.observe(size_bytes)


def observe_response_body_size(size_bytes: int) -> None:
    """Record output document size."""
    pandoc_response_body_bytes.observe(size_bytes)


def update_gauges_from_pandoc_metrics(pandoc_metrics: PandocMetrics) -> None:
    """
    Update Prometheus gauges from PandocMetrics current state.

    This function should be called before serving metrics to ensure
    gauges reflect the current state. It ONLY updates gauges, not counters.

    Note: Counters are incremented when events occur via the increment_* functions.

    Args:
        pandoc_metrics: PandocMetrics instance to collect metrics from
    """
    try:
        metrics = pandoc_metrics.get_metrics()

        # Update gauges only - cast to float for type safety
        error_rate = metrics["error_rate_percent"]
        uptime = metrics["uptime_seconds"]
        active = metrics["active_conversions"]
        avg_time = metrics["avg_conversion_time_ms"]

        if isinstance(error_rate, (int, float)):
            pandoc_conversion_error_rate_percent.set(float(error_rate))
        if isinstance(uptime, (int, float)):
            uptime_seconds.set(float(uptime))
        if isinstance(active, (int, float)):
            active_conversions.set(float(active))
        if isinstance(avg_time, (int, float)):
            avg_pandoc_conversion_time_seconds.set(float(avg_time) / 1000.0)

        # Update service info
        version = metrics.get("pandoc_version")
        service_ver = metrics.get("service_version")
        pandoc_info.info(
            {
                "version": str(version) if version else "unknown",
                "service_version": str(service_ver) if service_ver else "unknown",
            }
        )

        logger.debug("Prometheus gauges updated from PandocMetrics")

    except Exception as e:
        logger.error("Failed to update Prometheus gauges: %s", e, exc_info=True)
