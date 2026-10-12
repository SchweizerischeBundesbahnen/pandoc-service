"""Small shared helpers for environment configuration."""

import functools
import logging
import os
import re
from pathlib import Path

# API version for compatibility checking with docx-exporter.
# Increment this ONLY when making breaking changes to the API contract.
# Minor updates and bug fixes should NOT change this version.
API_VERSION = 1

_TRUTHY_VALUES = ("true", "1", "yes", "on")

# Bounds for the number of conversions which may run at the same time
DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS = 2
MIN_CONCURRENT_PANDOC_CONVERSIONS = 1
MAX_CONCURRENT_PANDOC_CONVERSIONS = 100

# Graceful shutdown bounds, in seconds
DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT = 30
MIN_GRACEFUL_SHUTDOWN_TIMEOUT = 1
MAX_GRACEFUL_SHUTDOWN_TIMEOUT = 300

# A pandoc stopped by its heap limit exits with this code
PANDOC_HEAP_EXHAUSTED_EXIT_CODE = 251

logger = logging.getLogger(__name__)


def get_bool_env(name: str, default: bool = False) -> bool:
    """Read a boolean environment variable."""
    return os.environ.get(name, str(default).lower()).lower() in _TRUTHY_VALUES


def get_max_concurrent_pandoc_conversions() -> int:
    """
    Read the conversion limit from the MAX_CONCURRENT_PANDOC_CONVERSIONS variable.

    Each conversion holds a worker thread and a pandoc process, and a PDF target adds a
    tectonic run on top. The limit bounds how many of those the container carries at once.

    Returns:
        Number of conversions allowed at the same time (default: 2). An invalid or out of
        range value falls back to the default.
    """
    value = os.environ.get("MAX_CONCURRENT_PANDOC_CONVERSIONS", str(DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS))
    try:
        limit = int(value)
        if not (MIN_CONCURRENT_PANDOC_CONVERSIONS <= limit <= MAX_CONCURRENT_PANDOC_CONVERSIONS):
            logger.warning(
                "MAX_CONCURRENT_PANDOC_CONVERSIONS must be between %d and %d, using default: %d",
                MIN_CONCURRENT_PANDOC_CONVERSIONS,
                MAX_CONCURRENT_PANDOC_CONVERSIONS,
                DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS,
            )
            return DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS
    except ValueError:
        logger.warning("Invalid MAX_CONCURRENT_PANDOC_CONVERSIONS value '%s', using default: %d", value, DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS)
        return DEFAULT_MAX_CONCURRENT_PANDOC_CONVERSIONS
    else:
        return limit


def get_graceful_shutdown_timeout() -> int:
    """
    Read the graceful shutdown timeout from the GRACEFUL_SHUTDOWN_TIMEOUT variable.

    On SIGTERM uvicorn stops accepting requests and waits for the running ones to
    finish. This timeout bounds that wait, so a stuck conversion cannot hold the
    container open until Docker sends SIGKILL.

    Returns:
        Timeout in seconds (default: 30). An invalid or out of range value falls back to the default.
    """
    value = os.environ.get("GRACEFUL_SHUTDOWN_TIMEOUT", str(DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT))
    try:
        timeout = int(value)
        if not (MIN_GRACEFUL_SHUTDOWN_TIMEOUT <= timeout <= MAX_GRACEFUL_SHUTDOWN_TIMEOUT):
            logger.warning(
                "GRACEFUL_SHUTDOWN_TIMEOUT must be between %d and %d, using default: %d",
                MIN_GRACEFUL_SHUTDOWN_TIMEOUT,
                MAX_GRACEFUL_SHUTDOWN_TIMEOUT,
                DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT,
            )
            return DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT
    except ValueError:
        logger.warning("Invalid GRACEFUL_SHUTDOWN_TIMEOUT value '%s', using default: %d", value, DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT)
        return DEFAULT_GRACEFUL_SHUTDOWN_TIMEOUT
    else:
        return timeout


# The memory of the container left to the service, Chromium and tectonic, which take some 0.6 GB
# beside a large conversion, and the least share of the memory the pandoc processes get together
_NON_PANDOC_MEMORY = 1024 * 1024 * 1024
_MIN_PANDOC_MEMORY_SHARE = 0.5
# A heap below this fails most documents: the configuration earns a warning, but the budget holds
_SMALL_PANDOC_HEAP_MB = 64
_PANDOC_HEAP_RE = re.compile(r"\d+[kmg]?", re.IGNORECASE)
_NO_PANDOC_HEAP_LIMIT = ("0", "none", "off", "false")
# The memory limit of the container: cgroup v2, then v1. v1 states no limit as a value near 2**63.
_CGROUP_MEMORY_LIMITS = (Path("/sys/fs/cgroup/memory.max"), Path("/sys/fs/cgroup/memory/memory.limit_in_bytes"))
_NO_CGROUP_LIMIT_ABOVE = 1 << 60


def get_pandoc_max_heap() -> str | None:
    """
    The heap limit of each pandoc process, in the form of ``+RTS -M``, or None for none.

    pandoc beyond its limit stops with "Heap exhausted" instead of growing until the kernel kills
    it, or kills another process of the container. PANDOC_MAX_HEAP states it, such as ``2g``, or
    ``0`` for no limit. Unset, the memory limit of the container less 1 GB for the rest of it, at
    least half of the memory, is shared between the conversions which may run at the same time.
    A container without a memory limit gets none.
    """
    value = os.environ.get("PANDOC_MAX_HEAP", "").strip().lower()
    if value in _NO_PANDOC_HEAP_LIMIT:
        return None
    if value:
        if _PANDOC_HEAP_RE.fullmatch(value):
            return value
        logger.warning("Invalid PANDOC_MAX_HEAP value '%s', deriving the limit from the memory of the container", value)
    memory = _container_memory_limit()
    if memory is None:
        return None
    pandoc_memory = max(memory - _NON_PANDOC_MEMORY, int(memory * _MIN_PANDOC_MEMORY_SHARE))
    conversions = get_max_concurrent_pandoc_conversions()
    share = max(pandoc_memory // conversions // (1024 * 1024), 1)
    if share < _SMALL_PANDOC_HEAP_MB:
        _warn_of_a_small_heap(share, conversions)
    return f"{share}m"


@functools.cache
def _warn_of_a_small_heap(share: int, conversions: int) -> None:
    """Warn once that the memory of the container leaves each conversion too little."""
    logger.warning(
        "%d conversions at the same time leave each pandoc %d MB, which fails most documents: raise the memory of the container or lower MAX_CONCURRENT_PANDOC_CONVERSIONS",
        conversions,
        share,
    )


def _container_memory_limit() -> int | None:
    """The memory limit of the container in bytes, or None where it has none or none can be read."""
    for path in _CGROUP_MEMORY_LIMITS:
        try:
            value = path.read_text(encoding="ascii").strip()
        except OSError:
            continue
        return int(value) if value.isdigit() and int(value) < _NO_CGROUP_LIMIT_ABOVE else None
    return None
