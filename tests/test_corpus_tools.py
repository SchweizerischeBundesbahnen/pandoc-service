"""Tests for tools/corpus: how a conversion of the corpus is compared with Word's PDF."""

from __future__ import annotations

import io
import json
from pathlib import Path

import pytest
from PIL import Image

from tools.corpus import corpus
from tools.corpus.pdf_lines import Line, _continues, _Run, document_lines, normalized


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


def test_the_next_cell_on_the_baseline_starts_a_line_of_its_own():
    """TeX writes the lines of a row which breaks between its lines cell after cell: two cell margins lie between."""
    run = _Run(0, 72, 700, 200, 12)
    assert _continues(run, 700, 204, 12), "the next word, a space on"
    assert not _continues(run, 700, 210.8, 12), "the next cell, two margins of 5.4 points on"


def test_a_pdf_without_text_has_no_lines(tmp_path: Path):
    buffer = io.BytesIO()
    Image.new("RGB", (612, 792), "white").save(buffer, "PDF")
    (tmp_path / "image.pdf").write_bytes(buffer.getvalue())
    assert document_lines(tmp_path / "image.pdf") == [[]]


def _result(same: int, shift: float, pages: tuple[int, int] = (1, 1), regions_matched: int = 50) -> dict:
    document = {"pages": list(pages), "lines": 10, "same_page": same, "any_page": 10, "mean_shift": shift, "words": 40, "words_in_place": 40, "regions": 50, "regions_matched": regions_matched, "problems": []}
    return {"reference": "macos", "total": {}, "documents": {"doc": document}}


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


def test_a_failed_conversion_leaves_no_pdf_of_an_earlier_run(tmp_path: Path):
    """The score would read the old PDF and miss that the conversion now fails."""
    import httpx

    docx = tmp_path / "doc.docx"
    docx.write_bytes(b"docx")
    (tmp_path / "doc.pdf").write_bytes(b"%PDF old")
    client = httpx.Client(transport=httpx.MockTransport(lambda _: httpx.Response(400, text="Bad Request")))
    assert corpus.convert_one(client, "http://service", docx, tmp_path) == "conversion failed: Bad Request"
    assert not (tmp_path / "doc.pdf").exists()


def test_the_free_corpus_names_open_fonts_for_the_fonts_of_word():
    from tools.corpus import free_fonts

    xml = '<w:rFonts w:ascii="segoe ui" w:hAnsi="Aptos" w:cs="Times New Roman" w:eastAsia="Aptos"/><a:latin typeface="Aptos Display"/><w:rFonts w:ascii="Symbol"/>'
    assert free_fonts.free_part("word/document.xml", xml) == ('<w:rFonts w:ascii="DejaVu Sans" w:hAnsi="Liberation Sans" w:cs="Liberation Serif" w:eastAsia="Aptos"/><a:latin typeface="Liberation Sans"/><w:rFonts w:ascii="Symbol"/>')


def test_the_free_corpus_sets_the_equations_in_latin_modern_math():
    from tools.corpus import free_fonts

    settings = "<w:settings><w:rsids/><w:themeFontLang/></w:settings>"
    assert free_fonts.free_part("word/settings.xml", settings) == '<w:settings><w:rsids/><m:mathPr><m:mathFont m:val="Latin Modern Math"/></m:mathPr><w:themeFontLang/></w:settings>'
    assert free_fonts.free_part("word/settings.xml", "<w:settings><m:mathPr/></w:settings>") == "<w:settings><m:mathPr/></w:settings>"


def test_the_free_corpus_keeps_every_other_part(tmp_path: Path):
    import zipfile

    from tools.corpus import free_fonts

    source = tmp_path / "source"
    source.mkdir()
    with zipfile.ZipFile(source / "a.docx", "w") as docx:
        docx.writestr("word/styles.xml", '<w:rFonts w:ascii="Calibri"/>')
        docx.writestr("word/media/image1.png", b"\x89PNG")
    assert free_fonts.main([str(source), str(tmp_path / "free")]) == 0
    with zipfile.ZipFile(tmp_path / "free" / "a.docx") as docx:
        assert docx.read("word/styles.xml") == b'<w:rFonts w:ascii="Carlito"/>'
        assert docx.read("word/media/image1.png") == b"\x89PNG"
    assert free_fonts.main([]) == 2


def test_a_word_is_in_place_within_two_points_across_and_down():
    from tools.corpus.pdf_lines import Word

    reference = [[Word(0, 72, 100, "one"), Word(0, 100, 100, "two"), Word(0, 72, 120, "three")]]
    converted = [[Word(0, 73.5, 101, "one"), Word(0, 103, 100, "two"), Word(0, 72, 120, "four")]]
    assert corpus._words_in_place(reference, converted) == 1
    assert corpus._words_in_place(reference, []) == 0


def _mask(boxes: list[tuple[int, int, int, int]], size: tuple[int, int] = (56, 56)) -> Image.Image:
    from PIL import ImageDraw

    from tools.corpus import regions

    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    for box in boxes:
        draw.rectangle(box, fill=regions.INK_PIXEL)
    return mask


def test_a_region_matches_when_its_ink_lies_within_the_tolerance():
    from tools.corpus import regions

    result = regions.compare_masks(_mask([(4, 4, 20, 20)]), _mask([(6, 6, 22, 22)]))
    assert (result.inked, result.matched, result.mismatched) == (1, 1, [])


def test_a_region_with_ink_on_one_side_only_does_not_match():
    from tools.corpus import regions

    result = regions.compare_masks(_mask([(4, 4, 20, 20)]), _mask([(32, 32, 50, 50)]))
    assert (result.inked, result.matched) == (2, 0)
    assert result.mismatched == [(0, 0, 14, 14), (14, 14, 28, 28)]


def test_a_larger_converted_page_is_compared_over_its_size():
    from tools.corpus import regions

    result = regions.compare_masks(_mask([]), _mask([(4, 60, 20, 70)], size=(56, 84)))
    assert (result.inked, result.matched) == (1, 0)


def test_the_regions_of_a_document_without_a_conversion_match_nothing(tmp_path: Path):
    from tools.corpus import regions

    pdf = REFERENCE / "test_alignment.pdf"
    pages = regions.compare_documents(pdf, None, tmp_path, "doc")
    assert pages
    assert all(page.matched == 0 for page in pages)
    assert (tmp_path / "doc-1.webp").exists()
    same = regions.compare_documents(pdf, pdf)
    assert all(page.matched == page.inked for page in same)


def test_fewer_regions_matched_got_worse(baseline: Path):
    corpus.write_baseline(_result(10, 0.5), "image")
    assert corpus.regressions(_result(10, 0.5, regions_matched=49), "image") == []
    assert corpus.regressions(_result(10, 0.5, regions_matched=45), "image") == ["doc: 90.0% regions matched, was 100.0%"]


def test_a_page_the_reference_lacks_matches_nothing(tmp_path: Path):
    from tools.corpus import regions

    one, more = REFERENCE / "test_alignment.pdf", REFERENCE / "test_fit_to_page.pdf"
    pages = regions.compare_documents(one, more, tmp_path, "doc")
    assert len(pages) > 1
    assert all(page.matched == 0 for page in pages[1:])
    assert sum(page.inked for page in pages[1:]) > 0
    assert (tmp_path / f"doc-{len(pages)}.webp").exists()
