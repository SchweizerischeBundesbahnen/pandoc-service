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
    from app.chromium_manager import ChromiumManager
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


# --- SVG-to-PNG (Chromium) metrics ---
# These mirror the ChromiumManager metrics used in weasyprint-service. Gauge
# names are prefixed with "chromium_" to avoid colliding with the service-level
# gauges above (e.g. uptime_seconds, active_conversions).
svg_conversions_total = Counter(
    "svg_conversions_total",
    "Total number of successful SVG to PNG conversions",
)

svg_conversion_failures_total = Counter(
    "svg_conversion_failures_total",
    "Total number of failed SVG to PNG conversions",
)

svg_conversion_duration_seconds = Histogram(
    "svg_conversion_duration_seconds",
    "SVG to PNG conversion duration in seconds",
    buckets=[0.05, 0.1, 0.25, 0.5, 1.0, 2.0, 5.0],
)

chromium_restarts_total = Counter(
    "chromium_restarts_total",
    "Total number of Chromium browser restarts",
)

svg_conversion_error_rate_percent = Gauge(
    "svg_conversion_error_rate_percent",
    "SVG conversion error rate as percentage",
)

avg_svg_conversion_time_seconds = Gauge(
    "avg_svg_conversion_time_seconds",
    "Average SVG to PNG conversion time in seconds",
)

chromium_uptime_seconds = Gauge(
    "chromium_uptime_seconds",
    "Chromium browser uptime in seconds",
)

chromium_consecutive_failures = Gauge(
    "chromium_consecutive_failures",
    "Current number of consecutive Chromium health check failures",
)

chromium_cpu_percent = Gauge(
    "chromium_cpu_percent",
    "Current Chromium CPU usage percentage",
)

chromium_memory_bytes = Gauge(
    "chromium_memory_bytes",
    "Current Chromium memory usage in bytes",
)

chromium_queue_size = Gauge(
    "chromium_queue_size",
    "Current number of requests waiting for an SVG conversion slot",
)

chromium_active_conversions = Gauge(
    "chromium_active_conversions",
    "Current number of in-flight SVG conversions",
)

# Browser info
chromium_info = Info(
    "chromium",
    "Chromium browser information",
)


def initialize_pandoc_info(pandoc_version: str, service_version: str) -> None:
    """
    Initialize pandoc service info metric once at startup.

    This must be called once during application startup. prometheus_client.Info.info()
    raises ValueError if called more than once with different values.

    Args:
        pandoc_version: Version of pandoc binary
        service_version: Version of the service application
    """
    pandoc_info.info(
        {
            "version": pandoc_version if pandoc_version else "unknown",
            "service_version": service_version if service_version else "unknown",
        }
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


def increment_svg_conversion_success(duration_seconds: float) -> None:
    """Increment successful SVG conversion counter and record duration."""
    svg_conversions_total.inc()
    svg_conversion_duration_seconds.observe(duration_seconds)


def increment_svg_conversion_failure() -> None:
    """Increment failed SVG conversion counter."""
    svg_conversion_failures_total.inc()


def increment_chromium_restart() -> None:
    """Increment Chromium restart counter."""
    chromium_restarts_total.inc()


def update_gauges_from_chromium_manager(chromium_manager: ChromiumManager) -> None:
    """
    Update Prometheus gauges from ChromiumManager current state.

    Called before serving metrics so gauges reflect current state. It only
    updates gauges; counters are incremented when events occur via the
    increment_* functions.

    Args:
        chromium_manager: ChromiumManager instance to collect metrics from.
    """
    try:
        metrics = chromium_manager.get_metrics()

        svg_conversion_error_rate_percent.set(float(metrics["error_svg_conversion_rate_percent"]))
        avg_svg_conversion_time_seconds.set(float(metrics["avg_svg_conversion_time_ms"]) / 1000.0)
        chromium_uptime_seconds.set(float(metrics["uptime_seconds"]))
        chromium_consecutive_failures.set(float(metrics.get("consecutive_failures", 0)))
        chromium_cpu_percent.set(float(metrics["current_cpu_percent"]))
        chromium_memory_bytes.set(float(metrics["current_chromium_memory_mb"]) * 1024 * 1024)  # MB -> bytes
        chromium_queue_size.set(float(metrics["queue_size"]))
        chromium_active_conversions.set(float(metrics["active_pdf_generations"]))

        chromium_version = chromium_manager.get_version()
        if chromium_version:
            chromium_info.info({"version": chromium_version})

        logger.debug("Prometheus gauges updated from ChromiumManager")

    except Exception as e:  # noqa: BLE001
        logger.exception("Failed to update Prometheus gauges: %s", e)


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

        logger.debug("Prometheus gauges updated from PandocMetrics")

    except Exception as e:
        logger.exception("Failed to update Prometheus gauges: %s", e)
