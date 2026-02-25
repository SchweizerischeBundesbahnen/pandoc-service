"""
Internal metrics tracking for pandoc-service.

This module provides a PandocMetrics class that tracks conversion statistics
and performance metrics internally. These metrics are then exposed via
Prometheus through the prometheus_metrics module.
"""

from __future__ import annotations

import os
import threading
import time
from dataclasses import dataclass, field


@dataclass
class PandocMetrics:
    """
    Tracks pandoc conversion metrics and statistics.

    This class maintains internal state for:
    - Conversion counts (success/failure)
    - Timing statistics (average, total)
    - Active conversion tracking
    - Service uptime
    """

    # Conversion counts
    total_conversions: int = 0
    failed_conversions: int = 0

    # Timing (in milliseconds)
    total_conversion_time_ms: float = 0.0
    avg_conversion_time_ms: float = 0.0

    # Active conversions
    active_conversions: int = 0

    # Service start time
    start_time: float = field(default_factory=time.time)

    # Thread lock for thread-safe updates (RLock allows reentrant locking)
    _lock: threading.RLock = field(default_factory=threading.RLock, repr=False)

    # Cached pandoc version
    _pandoc_version: str | None = field(default=None, repr=False)

    def record_conversion_start(self) -> None:
        """Record that a conversion has started."""
        with self._lock:
            self.active_conversions += 1

    def record_conversion_success(self, duration_ms: float) -> None:
        """
        Record a successful conversion.

        Args:
            duration_ms: Conversion duration in milliseconds
        """
        with self._lock:
            self.total_conversions += 1
            self.active_conversions = max(0, self.active_conversions - 1)
            self.total_conversion_time_ms += duration_ms
            self.avg_conversion_time_ms = self.total_conversion_time_ms / self.total_conversions

    def record_conversion_failure(self) -> None:
        """Record a failed conversion."""
        with self._lock:
            self.failed_conversions += 1
            self.active_conversions = max(0, self.active_conversions - 1)

    def get_error_rate(self) -> float:
        """
        Calculate the current error rate.

        Returns:
            Error rate as a percentage (0-100)
        """
        with self._lock:
            total_attempts = self.total_conversions + self.failed_conversions
            if total_attempts == 0:
                return 0.0
            return (self.failed_conversions / total_attempts) * 100.0

    def get_uptime_seconds(self) -> float:
        """
        Get service uptime.

        Returns:
            Uptime in seconds
        """
        return time.time() - self.start_time

    def set_pandoc_version(self, version: str | None) -> None:
        """Set the cached pandoc version."""
        self._pandoc_version = version

    def get_metrics(self) -> dict[str, int | float | str | None]:
        """
        Get all metrics as a dictionary.

        Returns:
            Dictionary containing all current metrics
        """
        with self._lock:
            return {
                "total_conversions": self.total_conversions,
                "failed_conversions": self.failed_conversions,
                "error_rate_percent": self.get_error_rate(),
                "avg_conversion_time_ms": self.avg_conversion_time_ms,
                "active_conversions": self.active_conversions,
                "uptime_seconds": self.get_uptime_seconds(),
                "pandoc_version": self._pandoc_version,
                "service_version": os.environ.get("PANDOC_SERVICE_VERSION", "unknown"),
            }


class _MetricsHolder:
    """Holder class for the global PandocMetrics singleton."""

    instance: PandocMetrics | None = None


def get_pandoc_metrics() -> PandocMetrics:
    """
    Get the global PandocMetrics instance.

    Returns:
        The global PandocMetrics singleton
    """
    if _MetricsHolder.instance is None:
        _MetricsHolder.instance = PandocMetrics()
    return _MetricsHolder.instance


def reset_pandoc_metrics() -> None:
    """Reset the global PandocMetrics instance (useful for testing)."""
    _MetricsHolder.instance = None
