import io
import logging

from pptx import Presentation
from pptx.util import Inches

# Standard slide sizes (width x height in inches)
SLIDE_SIZES = {
    "16:9": {"width": 10, "height": 5.63},  # 16:9 aspect ratio
    "WIDESCREEN": {"width": 13.33, "height": 7.5},  # Widescreen
    "4:3": {"width": 10, "height": 7.5},  # Standard
    "LETTER": {"width": 10, "height": 7.5},  # Letter (same as 4:3)
    "LEDGER": {"width": 13.333, "height": 10},  # Ledger/Tabloid
    "A4": {"width": 10.833, "height": 7.5},  # A4
    "A3": {"width": 14, "height": 10.5},  # A3
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


def _apply_slide_size(prs: Presentation, slide_size: str | None = None) -> None:  # type: ignore[valid-type]
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
    prs.slide_width = width  # type: ignore[attr-defined]
    prs.slide_height = height  # type: ignore[attr-defined]

    # Sanitize user input for logging to prevent log injection (CWE-117)
    safe_slide_size = slide_size_upper.replace("\r\n", "").replace("\n", "")
    logging.debug(f'Applied slide size {safe_slide_size}: {slide_dims["width"]}" x {slide_dims["height"]}"')
