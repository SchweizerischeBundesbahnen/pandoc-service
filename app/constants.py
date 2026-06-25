"""Small shared helpers for environment configuration."""

import os

_TRUTHY_VALUES = ("true", "1", "yes", "on")


def get_bool_env(name: str, default: bool = False) -> bool:
    """Read a boolean environment variable."""
    return os.environ.get(name, str(default).lower()).lower() in _TRUTHY_VALUES
