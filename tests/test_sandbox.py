"""Tests for the pandoc sandbox, which keeps a document from naming its own resources."""

import pytest

from app.pandoc_controller import _build_pandoc_command, is_sandbox_enabled


def build(**overrides: object) -> list[str]:
    """Build a conversion command with the arguments a plain HTML to DOCX run uses."""
    arguments: dict[str, object] = {
        "source_format": "html",
        "target_format": "docx",
        "source_path": "/tmp/source.html",
        "output_path": "/tmp/output.docx",
        "validated_options": [],
        "apply_docx_latex_filters": False,
    }
    arguments.update(overrides)
    return _build_pandoc_command(**arguments)  # type: ignore[arg-type]


def test_sandbox_is_on_by_default(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("PANDOC_SANDBOX", raising=False)

    assert is_sandbox_enabled() is True
    assert "--sandbox" in build()


@pytest.mark.parametrize("value", ["true", "1", "yes", "on", "TRUE"])
def test_sandbox_stays_on_for_truthy_values(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv("PANDOC_SANDBOX", value)

    assert "--sandbox" in build()


@pytest.mark.parametrize("value", ["false", "0", "no", "off"])
def test_sandbox_can_be_turned_off(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    """An operator who needs a document to load its own resources has to say so."""
    monkeypatch.setenv("PANDOC_SANDBOX", value)

    assert is_sandbox_enabled() is False
    assert "--sandbox" not in build()


@pytest.mark.parametrize(
    ("source_format", "target_format"),
    [("html", "docx"), ("html", "pdf"), ("docx", "pdf"), ("markdown", "html"), ("html", "pptx")],
)
def test_every_conversion_is_sandboxed(monkeypatch: pytest.MonkeyPatch, source_format: str, target_format: str) -> None:
    monkeypatch.delenv("PANDOC_SANDBOX", raising=False)

    assert "--sandbox" in build(source_format=source_format, target_format=target_format)


def test_sandbox_leaves_the_other_arguments_alone(monkeypatch: pytest.MonkeyPatch) -> None:
    """The flag is added, nothing else about the invocation changes."""
    monkeypatch.setenv("PANDOC_SANDBOX", "false")
    without = build(validated_options=["--toc"])
    monkeypatch.setenv("PANDOC_SANDBOX", "true")
    with_sandbox = build(validated_options=["--toc"])

    assert [argument for argument in with_sandbox if argument != "--sandbox"] == without
