"""Where the ink of a converted page stands apart from the reference, square by square.

Both pages are rendered at the same scale and cut into squares of CELL points. A square with ink in
either page matches when most of the ink of each page lies near ink of the other: within TOLERANCE
points, so anti-aliasing and a line off by a fraction of a point still match. Pictures, borders,
shading and equations count as much as text, which the lines and the words of the score leave out.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

import pypdfium2 as pdfium
from PIL import Image, ImageChops, ImageDraw, ImageFilter

if TYPE_CHECKING:
    from pathlib import Path

# About 5 mm
CELL = 14
# Pixels per point of the rendering
SCALE = 2
# A gray darker than this is ink
INK = 160
# How far ink may stand from ink of the other page, in points
TOLERANCE = 1.5
# The share of the ink of each page that has to lie near ink of the other
MATCH = 0.8
# Fewer pixels of ink than this in a square are dust: the square counts as empty
DUST = 4
INK_PIXEL = 255


@dataclass
class PageRegions:
    """The squares of one page: how many hold ink, how many match, and which do not, in points from the top left."""

    inked: int = 0
    matched: int = 0
    mismatched: list[tuple[int, int, int, int]] = field(default_factory=list)


def ink(image: Image.Image) -> Image.Image:
    """A mask of the ink of a grayscale rendering: 255 where ink, 0 elsewhere."""
    return image.point(lambda value: INK_PIXEL if value < INK else 0)


def _count(mask: Image.Image, box: tuple[int, int, int, int]) -> int:
    return mask.crop(box).histogram()[INK_PIXEL]


def compare_masks(reference: Image.Image, converted: Image.Image) -> PageRegions:
    """Compare two ink masks of one page square by square. A page of another size is compared over the larger one."""
    width, height = max(reference.width, converted.width), max(reference.height, converted.height)
    reference, converted = (_padded(mask, width, height) for mask in (reference, converted))
    size = 2 * round(TOLERANCE * SCALE) + 1
    near_reference = ImageChops.multiply(converted, reference.filter(ImageFilter.MaxFilter(size)))
    near_converted = ImageChops.multiply(reference, converted.filter(ImageFilter.MaxFilter(size)))
    result = PageRegions()
    step = CELL * SCALE
    for top in range(0, height, step):
        for left in range(0, width, step):
            box = (left, top, min(left + step, width), min(top + step, height))
            in_reference, in_converted = _count(reference, box), _count(converted, box)
            if in_reference < DUST and in_converted < DUST:
                continue
            result.inked += 1
            if in_reference >= DUST and in_converted >= DUST and _count(near_converted, box) >= MATCH * in_reference and _count(near_reference, box) >= MATCH * in_converted:
                result.matched += 1
            else:
                result.mismatched.append(tuple(round(value / SCALE) for value in box))  # ty: ignore[invalid-argument-type]
    return result


def _padded(mask: Image.Image, width: int, height: int) -> Image.Image:
    if mask.size == (width, height):
        return mask
    padded = Image.new("L", (width, height), 0)
    padded.paste(mask, (0, 0))
    return padded


def render(page: pdfium.PdfPage) -> Image.Image:
    return page.render(scale=SCALE, grayscale=True).to_pil().convert("L")


def difference_image(reference: Image.Image, converted: Image.Image, regions: PageRegions) -> Image.Image:
    """The reference in gray, the conversion over it in red, each square which does not match framed in orange."""
    width, height = max(reference.width, converted.width), max(reference.height, converted.height)
    base = Image.new("RGB", (width, height), "white")
    base.paste(Image.new("RGB", reference.size, (150, 150, 150)), (0, 0), ink(reference))
    base.paste(Image.new("RGB", converted.size, (220, 30, 30)), (0, 0), ink(converted))
    draw = ImageDraw.Draw(base)
    for left, top, right, bottom in regions.mismatched:
        draw.rectangle((left * SCALE, top * SCALE, right * SCALE - 1, bottom * SCALE - 1), outline=(255, 140, 0), width=2)
    return base


def compare_documents(reference_pdf: Path, converted_pdf: Path | None, difference_target: Path | None = None, prefix: str = "") -> list[PageRegions]:
    """Compare each page of the reference with the page of the conversion of its number, through the longer of the two:
    a page one of them lacks is blank there and matches nothing.

    With DIFFERENCE_TARGET, each page also gets its difference image there, as PREFIX-N.webp.
    """
    reference = pdfium.PdfDocument(reference_pdf)
    converted = pdfium.PdfDocument(converted_pdf) if converted_pdf is not None and converted_pdf.exists() else None
    pages = []
    try:
        for number in range(max(len(reference), len(converted) if converted is not None else 0)):
            reference_image = render(reference[number]) if number < len(reference) else None
            converted_image = render(converted[number]) if converted is not None and number < len(converted) else None
            size = (reference_image or converted_image).size  # ty: ignore[unresolved-attribute]
            reference_image = reference_image or Image.new("L", size, 255)
            converted_image = converted_image or Image.new("L", size, 255)
            regions = compare_masks(ink(reference_image), ink(converted_image))
            pages.append(regions)
            if difference_target is not None:
                difference_target.mkdir(parents=True, exist_ok=True)
                difference_image(reference_image, converted_image, regions).save(difference_target / f"{prefix}-{number + 1}.webp", quality=70)
    finally:
        reference.close()
        if converted is not None:
            converted.close()
    return pages
