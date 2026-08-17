"""Tests for the pandoc sandbox, which keeps a document from naming its own resources."""

import pytest

from app.pandoc_controller import FILTERS, _build_pandoc_command, is_sandbox_enabled


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


@pytest.mark.parametrize("value", ["enabled", " on ", "of", "", "  ", "yes please"])
def test_a_value_which_is_not_understood_keeps_the_sandbox(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    """A typo must not open the service, which is what a plain truthy check would do."""
    monkeypatch.setenv("PANDOC_SANDBOX", value)

    assert is_sandbox_enabled() is True
    assert "--sandbox" in build()


@pytest.mark.parametrize("target_format", ["pdf", "latex"])
def test_raw_tex_of_the_document_is_dropped_on_the_tex_paths(monkeypatch: pytest.MonkeyPatch, target_format: str) -> None:
    """tectonic runs outside the sandbox and reads what the TeX names, so the TeX goes first."""
    monkeypatch.delenv("PANDOC_SANDBOX", raising=False)
    command = build(target_format=target_format)
    strip = f"--lua-filter={FILTERS['strip_raw_tex']}"

    assert strip in command
    other_filters = [index for index, argument in enumerate(command) if argument.startswith("--lua-filter=") and argument != strip]
    assert all(command.index(strip) < index for index in other_filters), "the raw TeX of the document is dropped before the filters which emit their own"


@pytest.mark.parametrize("target_format", ["docx", "html", "pptx", "odt"])
def test_the_other_targets_keep_their_raw_blocks(monkeypatch: pytest.MonkeyPatch, target_format: str) -> None:
    monkeypatch.delenv("PANDOC_SANDBOX", raising=False)

    assert f"--lua-filter={FILTERS['strip_raw_tex']}" not in build(target_format=target_format)


def test_the_filter_is_not_added_without_the_sandbox(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PANDOC_SANDBOX", "false")

    assert f"--lua-filter={FILTERS['strip_raw_tex']}" not in build(target_format="pdf")


@pytest.mark.parametrize("source_format", ["markdown", "html", "latex", "docx", "epub", "fb2"])
def test_images_named_by_the_document_are_dropped_on_the_tex_paths(monkeypatch: pytest.MonkeyPatch, source_format: str) -> None:
    """An image becomes \\includegraphics, and the engine resolving it reads a file of the container.

    The filter asks the media bag rather than the source format, so an EPUB whose
    XHTML names an address is covered like a markdown source is.
    """
    monkeypatch.delenv("PANDOC_SANDBOX", raising=False)

    assert f"--lua-filter={FILTERS['strip_document_images']}" in build(source_format=source_format, target_format="pdf")


def test_images_are_kept_on_the_other_targets(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("PANDOC_SANDBOX", raising=False)

    assert f"--lua-filter={FILTERS['strip_document_images']}" not in build(source_format="markdown", target_format="docx")


def test_no_filter_is_added_without_the_sandbox(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PANDOC_SANDBOX", "false")
    command = build(target_format="pdf")

    assert f"--lua-filter={FILTERS['strip_document_images']}" not in command
    assert f"--lua-filter={FILTERS['strip_raw_tex']}" not in command
