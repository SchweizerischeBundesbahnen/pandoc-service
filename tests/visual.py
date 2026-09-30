"""Page-image comparison for the container tests.

A test converts its document through the service, turns the result into a
PDF, and compares every page with a reference PNG in tests/data/expected.
The comparison allows a small share of differing pixels, so antialiasing
does not fail a run, while an image of a different size or place does.

To create or refresh the references, run the tests with
UPDATE_VISUAL_REFERENCES=1 and look at every changed PNG before committing it.
On a failure the page and a diff image are written to tests/output.
"""

from __future__ import annotations

import os
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageChops

EXPECTED_DIR = Path(__file__).parent / "data" / "expected"
OUTPUT_DIR = Path(__file__).parent / "output"
DPI = 50
# A pixel differs when its grey value moves by more than this.
PIXEL_TOLERANCE = 32
# The share of differing pixels a page may have.
MAX_DIFFERING_SHARE = 0.002


def render_pages(pdf: bytes) -> list[Image.Image]:
    """Render every page of a PDF as a greyscale image."""
    document = pdfium.PdfDocument(pdf)
    try:
        # TODO Avoid stacking all pages in memory. Rather, use a generator to yield pages one by one.
        return [page.render(scale=DPI / 72, grayscale=True).to_pil() for page in document]
    finally:
        document.close()


def _differing_share(actual: Image.Image, expected: Image.Image) -> tuple[float, Image.Image]:
    diff = ImageChops.difference(actual, expected).point(lambda value: 255 if value > PIXEL_TOLERANCE else 0)
    differing = diff.histogram()[255]
    return differing / (actual.width * actual.height), diff


def assert_pages_match(name: str, pdf: bytes) -> None:
    """Compare the pages of a PDF with the references named <name>_page_<n>.png."""
    pages = render_pages(pdf)
    if os.environ.get("UPDATE_VISUAL_REFERENCES") == "1":
        # A refresh which rendered nothing would take the references away and put nothing in their
        # place, and the next run would have no baseline left to fail against
        assert pages, f"{name}: the PDF has no pages, so the references are kept as they are"
        EXPECTED_DIR.mkdir(parents=True, exist_ok=True)
        for stale in EXPECTED_DIR.glob(f"{name}_page_*.png"):
            stale.unlink()
        for number, page in enumerate(pages, start=1):
            page.save(EXPECTED_DIR / f"{name}_page_{number}.png")
        return

    references = sorted(EXPECTED_DIR.glob(f"{name}_page_*.png"), key=lambda path: int(path.stem.rsplit("_", 1)[1]))
    assert references, f"no reference for {name}; run with UPDATE_VISUAL_REFERENCES=1 and review the PNGs"
    assert len(pages) == len(references), f"{name}: {len(pages)} pages, the reference has {len(references)}"

    failures = []
    for number, (page, reference) in enumerate(zip(pages, references, strict=True), start=1):
        expected = Image.open(reference).convert("L")
        if page.size != expected.size:
            failures.append(f"page {number}: size {page.size}, the reference is {expected.size}")
            continue
        share, diff = _differing_share(page, expected)
        if share > MAX_DIFFERING_SHARE:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            page.save(OUTPUT_DIR / f"{name}_page_{number}.png")
            diff.save(OUTPUT_DIR / f"{name}_page_{number}_diff.png")
            failures.append(f"page {number}: {share:.2%} of the pixels differ, see {OUTPUT_DIR}")
    assert not failures, f"{name}:\n" + "\n".join(failures)
