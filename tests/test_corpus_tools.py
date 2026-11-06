"""Tests for tools/corpus: how a conversion of the corpus is compared with Word's PDF."""

from __future__ import annotations

import io
import json
from pathlib import Path

import pytest
from PIL import Image

from tools.corpus import corpus
from tools.corpus.pdf_lines import Line, document_lines, normalized


def test_text_is_compared_with_single_spaces_and_short_leaders():
    assert normalized("  Introduction . . . . . . . 3 ") == "Introduction ... 3"
    assert normalized("one\u2010two") == "one-two"


def test_a_character_outside_the_bmp_comes_back_whole():
    assert normalized("a \ud83d\ude00 b") == "a \U0001f600 b"


def _line(page: int, y: float, text: str) -> Line:
    return Line(page, 72, y, 100, text)


def test_a_line_matches_the_nearest_line_of_its_text_on_its_page():
    reference = [[_line(0, 100, "Cell one"), _line(0, 200, "Cell one")], [_line(1, 100, "Only in Word")]]
    converted = [[_line(0, 203, "Cell one"), _line(0, 101, "Cell one")], [_line(1, 100, "Moved")], [_line(2, 100, "Only in Word")]]
    matches, anywhere = corpus._match(reference, converted)
    assert [match["dy"] for match in matches] == [1, 3, None]
    assert anywhere == 3


REFERENCE = Path(__file__).resolve().parents[1] / "corpus" / "reference" / "macos"


def test_the_lines_of_word_are_runs_on_one_baseline():
    """The cells of a row of the table are lines of their own: a gap of several letters lies between them."""
    lines = document_lines(REFERENCE / "test_attribute_name_alignment.pdf")[0]
    assert [line.text for line in lines[:3]] == ["Attribute Name Alignment Test", "EL-101 - User name must be validated when a new user account is created", "Linked Work"]
    assert lines[1].y == pytest.approx(126.5, abs=0.1)
    assert lines[2].x == pytest.approx(73, abs=0.1)


def test_a_pdf_without_text_has_no_lines(tmp_path: Path):
    buffer = io.BytesIO()
    Image.new("RGB", (612, 792), "white").save(buffer, "PDF")
    (tmp_path / "image.pdf").write_bytes(buffer.getvalue())
    assert document_lines(tmp_path / "image.pdf") == [[]]


def _result(same: int, shift: float, pages: tuple[int, int] = (1, 1)) -> dict:
    return {"reference": "macos", "total": {}, "documents": {"doc": {"pages": list(pages), "lines": 10, "same_page": same, "any_page": 10, "mean_shift": shift, "problems": []}}}


@pytest.fixture
def baseline(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    path = tmp_path / "baseline.json"
    monkeypatch.setattr(corpus, "BASELINE", path)
    return path


def test_the_baseline_keeps_each_run_of_each_reference(baseline: Path):
    corpus.write_baseline(_result(9, 1.0), "image")
    corpus.write_baseline(_result(10, 0.5), "word-fonts")
    stored = json.loads(baseline.read_text())
    assert stored["macos"]["image"]["documents"]["doc"]["same_page"] == 9
    assert stored["macos"]["word-fonts"]["documents"]["doc"]["mean_shift"] == 0.5
    assert "problems" not in stored["macos"]["image"]["documents"]["doc"]


def test_fewer_lines_on_their_page_a_larger_shift_or_another_page_count_got_worse(baseline: Path):
    corpus.write_baseline(_result(9, 1.0), "image")
    assert corpus.regressions(_result(9, 1.4), "image") == []
    assert corpus.regressions(_result(10, 0.2), "image") == []
    assert len(corpus.regressions(_result(8, 1.0), "image")) == 1
    assert len(corpus.regressions(_result(9, 1.6), "image")) == 1
    assert len(corpus.regressions(_result(9, 1.0, pages=(1, 2)), "image")) == 1


def test_a_run_without_a_baseline_is_reported(baseline: Path):
    assert corpus.regressions(_result(9, 1.0), "word-fonts") == ["no baseline for macos word-fonts"]


def test_the_report_lists_the_documents_furthest_off_first():
    result = _result(9, 1.0)
    result["documents"]["other"] = {**result["documents"]["doc"], "mean_shift": 5.0}
    result["total"] = {"same_page": 90.0, "any_page": 100.0, "pages_as_reference": 2, "documents": 2, "mean_shift": 3.0}
    text = corpus.report(result)
    assert text.index("| other |") < text.index("| doc |")
