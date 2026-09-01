"""Small shared helpers for environment configuration."""

import logging
import os

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
