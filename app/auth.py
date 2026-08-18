"""Optional API key authentication for the pandoc service.

Authentication is disabled by default. It activates when the ``API_KEY``
environment variable holds at least one non-empty key. Several keys can be
configured as a comma-separated list, which allows key rotation without
downtime.

Clients send the key in one of two headers:

- ``X-API-Key: <key>``
- ``Authorization: Bearer <key>``

The key is checked twice. ``is_request_authorized`` serves the middleware,
which rejects a request before its body is buffered. ``require_api_key`` is the
route dependency, which documents both schemes in the OpenAPI schema and covers
a route the middleware misses.
"""

from __future__ import annotations

import logging
import os
import secrets
from typing import TYPE_CHECKING, Annotated

from fastapi import Depends, HTTPException, Request, status
from fastapi.security import APIKeyHeader, HTTPAuthorizationCredentials, HTTPBearer

if TYPE_CHECKING:
    from collections.abc import Mapping

logger = logging.getLogger(__name__)

API_KEY_ENV_VAR = "API_KEY"
API_KEY_HEADER_NAME = "X-API-Key"
AUTHORIZATION_HEADER_NAME = "Authorization"
BEARER_SCHEME = "bearer"

UNAUTHORIZED_MESSAGE = "Invalid or missing API key"

# The paths the key protects. The middleware and the route dependency read this
# one definition, so the two cannot drift apart.
PROTECTED_PATHS = frozenset({"/docx-template", "/pptx-template"})
PROTECTED_PATH_PREFIXES = ("/convert/",)

# auto_error=False keeps both schemes optional, so a missing header reaches
# require_api_key instead of failing inside the security dependency.
_api_key_header = APIKeyHeader(name=API_KEY_HEADER_NAME, auto_error=False, description="API key, required when the service is started with API_KEY configured.")
_bearer_scheme = HTTPBearer(auto_error=False, description="API key sent as a bearer token, required when the service is started with API_KEY configured.")


class ApiKeyError(HTTPException):
    """Raised when a request carries no valid API key.

    The controller answers it in plain text, the format the other errors of
    this service use.
    """

    def __init__(self) -> None:
        super().__init__(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=UNAUTHORIZED_MESSAGE,
            headers={"WWW-Authenticate": "Bearer"},
        )


def get_api_keys() -> tuple[str, ...]:
    """
    Read the configured API keys from the environment.

    Returns:
        Tuple of non-empty keys parsed from the API_KEY variable. Empty when
        authentication is disabled.
    """
    raw = os.environ.get(API_KEY_ENV_VAR, "")
    return tuple(key for key in (part.strip() for part in raw.split(",")) if key)


def is_auth_enabled() -> bool:
    """
    Check whether API key authentication is active.

    Returns:
        True if at least one API key is configured.
    """
    return bool(get_api_keys())


def _matches_any(candidate: str, api_keys: tuple[str, ...]) -> bool:
    """
    Compare a candidate key against all configured keys in constant time.

    The comparison runs on bytes: secrets.compare_digest raises TypeError for a
    str holding a character above U+007F, and a header may carry such a byte.
    Starlette decodes header bytes as latin-1, so encoding back to latin-1
    restores the bytes the client sent, and it cannot fail for such a value.

    Every key is checked without an early exit, so the comparison time does not
    depend on which key matches.

    Args:
        candidate: Key presented by the client.
        api_keys: Configured keys.

    Returns:
        True if the candidate matches one of the configured keys.
    """
    presented = candidate.encode("latin-1", errors="replace")
    matched = False
    for api_key in api_keys:
        if secrets.compare_digest(presented, api_key.encode()):
            matched = True
    return matched


def is_protected_path(path: str) -> bool:
    """
    Check whether a request path needs an API key.

    A trailing slash is ignored, so the redirect variant of a path is treated
    like the path itself.

    Args:
        path: Path of the incoming request.

    Returns:
        True if the path belongs to a protected endpoint.
    """
    return path.rstrip("/") in PROTECTED_PATHS or path.startswith(PROTECTED_PATH_PREFIXES)


def _extract_credentials(headers: Mapping[str, str]) -> list[str]:
    """
    Collect the API keys a request presents.

    The headers are read directly, which is what the middleware needs: it runs
    before routing, so the security dependencies are not available there.

    Args:
        headers: Headers of the incoming request.

    Returns:
        Non-empty candidate keys, from the X-API-Key header and from a bearer
        token in the Authorization header.
    """
    candidates = [headers.get(API_KEY_HEADER_NAME)]
    authorization = headers.get(AUTHORIZATION_HEADER_NAME)
    if authorization:
        scheme, _, credentials = authorization.partition(" ")
        if scheme.lower() == BEARER_SCHEME:
            candidates.append(credentials.strip())
    return [candidate for candidate in candidates if candidate]


def _is_authorized(presented: list[str], api_keys: tuple[str, ...], path: str) -> bool:
    """
    Match the presented keys against the configured ones and log a rejection.

    Both schemes are advertised as alternatives, so either credential admits
    the request. A stale header next to a valid bearer token must not reject.

    Args:
        presented: Candidate keys the request carries.
        api_keys: Configured keys.
        path: Path of the request, logged on a rejection.

    Returns:
        True if one of the candidates matches a configured key.
    """
    matched = False
    for candidate in presented:
        if _matches_any(candidate, api_keys):
            matched = True
    if not matched:
        # Never log the presented value, only the reason and the target path.
        logger.warning("Rejected unauthenticated request to %s: %s", path, "invalid API key" if presented else "missing API key")
    return matched


def is_request_authorized(request: Request) -> bool:
    """
    Check a request before its body is read.

    The middleware calls this ahead of the size check, so an unauthenticated
    request costs no memory.

    Args:
        request: Incoming request.

    Returns:
        True if authentication is disabled, the path is open, or the request
        carries a valid key.
    """
    api_keys = get_api_keys()
    if not api_keys or not is_protected_path(request.url.path):
        return True
    return _is_authorized(_extract_credentials(request.headers), api_keys, request.url.path)


def require_api_key(
    request: Request,
    header_key: Annotated[str | None, Depends(_api_key_header)] = None,
    bearer: Annotated[HTTPAuthorizationCredentials | None, Depends(_bearer_scheme)] = None,
) -> None:
    """
    Reject the request when API key authentication fails.

    The dependency is a no-op while authentication is disabled, which keeps the
    previous behavior for deployments without API_KEY.

    Args:
        request: Incoming request, used for the rejection log only.
        header_key: Value of the X-API-Key header, if present.
        bearer: Credentials from the Authorization header, if present.

    Raises:
        ApiKeyError: 401 when the key is missing or invalid.
    """
    api_keys = get_api_keys()
    if not api_keys:
        return

    presented = [candidate for candidate in (header_key, bearer.credentials if bearer else None) if candidate]
    if not _is_authorized(presented, api_keys, request.url.path):
        raise ApiKeyError
