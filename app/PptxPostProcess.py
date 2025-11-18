import io
import logging
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.util import Inches

# Standard slide sizes (width x height in inches)
SLIDE_SIZES = {
    "16:9": {"width": 10, "height": 5.625},      # Widescreen
    "4:3": {"width": 10, "height": 7.5},         # Standard
    "LETTER": {"width": 10, "height": 7.5},      # Letter (same as 4:3)
    "LEDGER": {"width": 13.333, "height": 10},   # Ledger/Tabloid
    "A4": {"width": 10.833, "height": 7.5},      # A4
    "A3": {"width": 14, "height": 10.5},         # A3
}


def process(pptx_bytes: bytes, slide_size: str | None = None) -> bytes:
    """
    Process a PPTX file to apply slide size modifications.

    Args:
        pptx_bytes: The PPTX file content as bytes
        slide_size: The desired slide size (e.g., "16:9", "4:3", "A4")

    Returns:
        Modified PPTX content as bytes
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    _apply_slide_size(prs, slide_size)
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def _apply_slide_size(prs: Presentation, slide_size: str | None = None) -> None:
    """
    Apply slide size to a presentation.

    Args:
        prs: The Presentation object
        slide_size: The desired slide size (e.g., "16:9", "4:3", "A4")
    """
    # If no slide size specified, no modifications needed
    if slide_size is None:
        return

    # Normalize slide_size to uppercase for case-insensitive lookup
    slide_size_upper = slide_size.upper()

    # Get dimensions for the specified slide size
    if slide_size_upper not in SLIDE_SIZES:
        raise ValueError(f"Unsupported slide size: {slide_size}. Supported sizes: {', '.join(SLIDE_SIZES.keys())}")

    slide_dims = SLIDE_SIZES[slide_size_upper]
    width = Inches(slide_dims["width"])
    height = Inches(slide_dims["height"])

    # Apply the slide size to the presentation
    prs.slide_width = width
    prs.slide_height = height

    logging.debug(f"Applied slide size {slide_size_upper}: {slide_dims['width']}\" x {slide_dims['height']}\"")


# Just for manual test purposes. Accepts path to pptx to process.
def main() -> int:
    import sys

    MIN_ARGS = 2  # script name + pptx path
    MAX_ARGS = 3  # script name + pptx path + slide_size
    PPTX_PATH_ARG_INDEX = 1
    SLIDE_SIZE_ARG_INDEX = 2

    if not (MIN_ARGS <= len(sys.argv) <= MAX_ARGS):
        logging.info("Usage: <path_to_pptx> [slide_size]")
        return 1

    pptx_path = sys.argv[PPTX_PATH_ARG_INDEX]
    slide_size = sys.argv[SLIDE_SIZE_ARG_INDEX] if len(sys.argv) > SLIDE_SIZE_ARG_INDEX and sys.argv[SLIDE_SIZE_ARG_INDEX] != "None" else None

    with Path(pptx_path).open("rb") as pptx_file_reader:
        result_bytes = process(pptx_file_reader.read(), slide_size)

    with Path(pptx_path).open("wb") as pptx_file_writer:
        pptx_file_writer.write(result_bytes)

    logging.debug(f"Successfully modified slide size in {pptx_path}")
    return 0


if __name__ == "__main__":
    main()
