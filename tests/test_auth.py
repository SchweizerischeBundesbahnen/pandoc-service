"""Tests for the optional API key authentication."""

from unittest.mock import patch

import pytest
from fastapi import Depends, FastAPI
from starlette.responses import Response
from starlette.testclient import TestClient

from app import pandoc_controller
from app.auth import API_KEY_ENV_VAR, get_api_keys, is_auth_enabled, is_protected_path, require_api_key
from app.pandoc_controller import app

SIMPLE_HTML = "<html><body><h1>Hello</h1></body></html>"
CONVERT_HTML_TO_DOCX = "/convert/html/to/docx"

PROTECTED_ENDPOINTS = [
    ("post", CONVERT_HTML_TO_DOCX),
    ("post", "/convert/html/to/docx-with-template"),
    ("post", "/convert/html/to/pptx-with-template"),
    ("get", "/docx-template"),
    ("get", "/pptx-template"),
]


@pytest.fixture
def protected_client() -> TestClient:
    """Client for a minimal app guarded by the API key dependency."""
    protected_app = FastAPI()

    @protected_app.get("/protected", dependencies=[Depends(require_api_key)])
    async def protected() -> dict[str, str]:
        return {"status": "ok"}

    return TestClient(protected_app)


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("", ()),
        ("   ", ()),
        (",,", ()),
        ("secret", ("secret",)),
        ("  secret  ", ("secret",)),
        ("first,second", ("first", "second")),
        ("first, second , ,third", ("first", "second", "third")),
    ],
)
def test_get_api_keys_parsing(monkeypatch: pytest.MonkeyPatch, raw: str, expected: tuple[str, ...]) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, raw)
    assert get_api_keys() == expected


def test_get_api_keys_unset(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(API_KEY_ENV_VAR, raising=False)
    assert get_api_keys() == ()
    assert is_auth_enabled() is False


def test_is_auth_enabled_with_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert is_auth_enabled() is True


def test_no_key_configured_allows_request(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.delenv(API_KEY_ENV_VAR, raising=False)
    assert protected_client.get("/protected").status_code == 200


def test_valid_api_key_header(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"X-API-Key": "secret"}).status_code == 200


def test_api_key_header_name_is_case_insensitive(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"x-api-key": "secret"}).status_code == 200


def test_valid_bearer_token(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"Authorization": "Bearer secret"}).status_code == 200


def test_any_configured_key_is_accepted(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "first,second")
    assert protected_client.get("/protected", headers={"X-API-Key": "first"}).status_code == 200
    assert protected_client.get("/protected", headers={"X-API-Key": "second"}).status_code == 200


def test_missing_key_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    response = protected_client.get("/protected")
    assert response.status_code == 401
    assert response.headers["WWW-Authenticate"] == "Bearer"


def test_invalid_key_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"X-API-Key": "wrong"}).status_code == 401


def test_empty_key_header_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"X-API-Key": ""}).status_code == 401


def test_invalid_bearer_token_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"Authorization": "Bearer wrong"}).status_code == 401


def test_non_bearer_authorization_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"Authorization": "Basic secret"}).status_code == 401


def test_either_credential_is_enough(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    """Both schemes are alternatives, so a stale credential next to a valid one does not reject."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"X-API-Key": "wrong", "Authorization": "Bearer secret"}).status_code == 200
    assert protected_client.get("/protected", headers={"X-API-Key": "secret", "Authorization": "Bearer wrong"}).status_code == 200


def test_both_credentials_invalid_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"X-API-Key": "wrong", "Authorization": "Bearer also-wrong"}).status_code == 401


def test_non_ascii_key_header_is_rejected(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    """A header byte above 0x7F must answer 401, not fail the comparison."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    assert protected_client.get("/protected", headers={"X-API-Key": "sécret".encode()}).status_code == 401


def test_non_ascii_configured_key_is_accepted(monkeypatch: pytest.MonkeyPatch, protected_client: TestClient) -> None:
    """A configured key holding a non-ASCII character still matches the bytes the client sends."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "sécret")
    assert protected_client.get("/protected", headers={"X-API-Key": "sécret".encode()}).status_code == 200
    assert protected_client.get("/protected", headers={"X-API-Key": "secret"}).status_code == 401


@pytest.mark.parametrize(("method", "path"), PROTECTED_ENDPOINTS)
def test_protected_endpoints_require_key(monkeypatch: pytest.MonkeyPatch, method: str, path: str) -> None:
    """Every conversion and template endpoint rejects a request without a key."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    with TestClient(app) as test_client:
        body = {"content": SIMPLE_HTML} if method == "post" else {}
        response = getattr(test_client, method)(path, **body)
        assert response.status_code == 401
        assert response.headers["WWW-Authenticate"] == "Bearer"
        assert response.headers["content-type"].startswith("text/plain")
        assert response.text == "Invalid or missing API key"


def test_open_endpoints_stay_open(monkeypatch: pytest.MonkeyPatch) -> None:
    """The endpoints the healthcheck and the schema need are reachable without a key."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    with TestClient(app) as test_client:
        assert test_client.get("/version").status_code == 200
        assert test_client.get("/health").status_code in (200, 503)
        assert test_client.get("/static/openapi.json").status_code == 200


def test_conversion_accepts_valid_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    with (
        patch("app.pandoc_controller.run_pandoc_conversion", return_value=b"DOCX content"),
        patch("app.pandoc_controller.postprocess_and_build_response", return_value=Response(b"DOCX content")),
        TestClient(app) as test_client,
    ):
        assert test_client.post(CONVERT_HTML_TO_DOCX, content=SIMPLE_HTML, headers={"X-API-Key": "secret"}).status_code == 200
        assert test_client.post(CONVERT_HTML_TO_DOCX, content=SIMPLE_HTML, headers={"Authorization": "Bearer secret"}).status_code == 200


def test_conversion_open_without_key(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(API_KEY_ENV_VAR, raising=False)
    with (
        patch("app.pandoc_controller.run_pandoc_conversion", return_value=b"DOCX content"),
        patch("app.pandoc_controller.postprocess_and_build_response", return_value=Response(b"DOCX content")),
        TestClient(app) as test_client,
    ):
        assert test_client.post(CONVERT_HTML_TO_DOCX, content=SIMPLE_HTML).status_code == 200


@pytest.mark.parametrize(
    ("path", "expected"),
    [
        ("/convert/html/to/docx", True),
        ("/convert/html/to/docx-with-template", True),
        ("/convert/html/to/pptx-with-template", True),
        ("/convert/html/to/unknown", True),
        ("/docx-template", True),
        ("/pptx-template", True),
        ("/docx-template/", True),
        ("/health", False),
        ("/version", False),
        ("/static/openapi.json", False),
        ("/api/docs", False),
        ("/", False),
        ("/converted", False),
    ],
)
def test_is_protected_path(path: str, expected: bool) -> None:
    assert is_protected_path(path) is expected


def test_unauthenticated_oversized_body_is_rejected_before_the_size_check(monkeypatch: pytest.MonkeyPatch) -> None:
    """The key is vetted first, so an anonymous upload never reaches the body buffer."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    monkeypatch.setattr(pandoc_controller, "data_limit", 10)
    with TestClient(app) as test_client:
        response = test_client.post(CONVERT_HTML_TO_DOCX, content=b"x" * 100)
        assert response.status_code == 401
        assert response.headers["WWW-Authenticate"] == "Bearer"


def test_authenticated_oversized_body_still_hits_the_size_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    monkeypatch.setattr(pandoc_controller, "data_limit", 10)
    with TestClient(app) as test_client:
        response = test_client.post(CONVERT_HTML_TO_DOCX, content=b"x" * 100, headers={"X-API-Key": "secret"})
        assert response.status_code == 413


def test_open_path_keeps_the_size_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    """An open endpoint is not touched by the key check, only by the size check."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    monkeypatch.setattr(pandoc_controller, "data_limit", 10)
    with TestClient(app) as test_client:
        assert test_client.post("/version", content=b"x" * 100).status_code == 413


def test_unknown_protected_path_is_rejected(monkeypatch: pytest.MonkeyPatch) -> None:
    """The middleware answers before routing, so an unrouted convert path also needs the key."""
    monkeypatch.setenv(API_KEY_ENV_VAR, "secret")
    with TestClient(app) as test_client:
        assert test_client.post("/convert/html/to/unknown", content=SIMPLE_HTML).status_code == 401
