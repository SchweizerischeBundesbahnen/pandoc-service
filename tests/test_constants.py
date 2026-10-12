"""Tests for the shared environment helpers."""

import pytest

from app import constants
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


@pytest.fixture
def cgroup(tmp_path, monkeypatch):
    """The memory limit files of the container, none of them there until a test writes one."""
    v2, v1 = tmp_path / "memory.max", tmp_path / "memory.limit_in_bytes"
    monkeypatch.setattr(constants, "_CGROUP_MEMORY_LIMITS", (v2, v1))
    monkeypatch.delenv("PANDOC_MAX_HEAP", raising=False)
    monkeypatch.delenv("MAX_CONCURRENT_PANDOC_CONVERSIONS", raising=False)
    return v2, v1


@pytest.mark.parametrize("value", ["2g", "512M", "1048576", "300k"])
def test_a_stated_pandoc_heap_limit_is_taken_as_it_is(cgroup, monkeypatch, value):
    cgroup[0].write_text("8589934592\n")
    monkeypatch.setenv("PANDOC_MAX_HEAP", value)

    assert constants.get_pandoc_max_heap() == value.lower()


@pytest.mark.parametrize("value", ["0", "off", "none", "false", "OFF"])
def test_a_pandoc_heap_limit_can_be_turned_off(cgroup, monkeypatch, value):
    cgroup[0].write_text("8589934592\n")
    monkeypatch.setenv("PANDOC_MAX_HEAP", value)

    assert constants.get_pandoc_max_heap() is None


def test_the_pandoc_heap_limit_shares_the_container_memory_less_a_reserve(cgroup, monkeypatch):
    cgroup[0].write_text("8589934592\n")  # 8 GiB less 1 GiB, for two conversions
    assert constants.get_pandoc_max_heap() == "3584m"

    monkeypatch.setenv("MAX_CONCURRENT_PANDOC_CONVERSIONS", "4")
    assert constants.get_pandoc_max_heap() == "1792m"


def test_a_small_container_gives_the_pandoc_processes_half_its_memory(cgroup):
    cgroup[0].write_text(str(1536 * 1024 * 1024))  # less 1 GiB would leave a third of it

    assert constants.get_pandoc_max_heap() == "384m"


def test_an_invalid_pandoc_heap_limit_falls_back_to_the_container_memory(cgroup, monkeypatch):
    cgroup[0].write_text("8589934592\n")
    monkeypatch.setenv("PANDOC_MAX_HEAP", "lots")

    assert constants.get_pandoc_max_heap() == "3584m"


def test_the_container_memory_is_read_from_cgroup_v1_as_well(cgroup):
    cgroup[1].write_text("4294967296\n")

    assert constants.get_pandoc_max_heap() == "1536m"  # 4 GiB less 1 GiB, for two conversions


@pytest.mark.parametrize(("v2", "v1"), [(None, None), ("max\n", None), (None, "9223372036854771712\n")])
def test_a_container_without_a_memory_limit_gets_no_pandoc_heap_limit(cgroup, v2, v1):
    for path, value in zip(cgroup, (v2, v1), strict=True):
        if value is not None:
            path.write_text(value)

    assert constants.get_pandoc_max_heap() is None


def test_the_pandoc_heap_limits_never_add_up_to_more_than_the_container_memory(cgroup, monkeypatch, caplog):
    cgroup[0].write_text(str(2 * 1024 * 1024 * 1024))
    monkeypatch.setenv("MAX_CONCURRENT_PANDOC_CONVERSIONS", "40")

    assert constants.get_pandoc_max_heap() == "25m"  # 1 GiB, half of the memory, for 40 conversions
    assert "40 conversions at the same time leave each pandoc 25 MB" in caplog.text
