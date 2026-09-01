"""Tests for the shared environment helpers."""

import pytest

from app.constants import (
    DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT,
    DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS,
    get_graceful_shutdown_timeout,
    get_max_concurrent_pandoc_conversions,
)


def test_graceful_shutdown_timeout_defaults(monkeypatch):
    """Without the variable the timeout falls back to the default."""
    monkeypatch.delenv("GRACEFUL_SHUTDOWN_TIMEOUT", raising=False)

    assert get_graceful_shutdown_timeout() == DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT


def test_graceful_shutdown_timeout_reads_the_variable(monkeypatch):
    """A valid value is used as is."""
    monkeypatch.setenv("GRACEFUL_SHUTDOWN_TIMEOUT", "45")

    assert get_graceful_shutdown_timeout() == 45


@pytest.mark.parametrize("value", ["not-a-number", "0", "301"])
def test_graceful_shutdown_timeout_rejects_bad_values(monkeypatch, value):
    """An invalid or out of range value falls back to the default."""
    monkeypatch.setenv("GRACEFUL_SHUTDOWN_TIMEOUT", value)

    assert get_graceful_shutdown_timeout() == DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT


def test_max_concurrent_conversions_defaults(monkeypatch):
    """Without the variable the limit falls back to the default."""
    monkeypatch.delenv("MAX_CONCURRENT_PANDOC_CONVERSIONS", raising=False)

    assert get_max_concurrent_pandoc_conversions() == DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS


def test_max_concurrent_conversions_reads_the_variable(monkeypatch):
    """A valid value is used as is."""
    monkeypatch.setenv("MAX_CONCURRENT_PANDOC_CONVERSIONS", "5")

    assert get_max_concurrent_pandoc_conversions() == 5


@pytest.mark.parametrize("value", ["not-a-number", "0", "101"])
def test_max_concurrent_conversions_rejects_bad_values(monkeypatch, value):
    """An invalid or out of range value falls back to the default."""
    monkeypatch.setenv("MAX_CONCURRENT_PANDOC_CONVERSIONS", value)

    assert get_max_concurrent_pandoc_conversions() == DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS
