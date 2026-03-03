import logging
from io import BytesIO
from typing import TypedDict
from zipfile import ZIP_DEFLATED, ZipFile

from defusedxml import ElementTree


# Standard slide sizes (width x height in inches)
class Dimensions(TypedDict):
    width: float
    height: float


SLIDE_SIZES: dict[str, Dimensions] = {
    "16:9": {"width": 10, "height": 5.63},  # 16:9 aspect ratio
    "WIDESCREEN": {"width": 13.33, "height": 7.5},  # Widescreen
    "4:3": {"width": 10, "height": 7.5},  # Standard
    "LETTER": {"width": 10, "height": 7.5},  # Letter (same as 4:3)
    "LEDGER": {"width": 13.333, "height": 10},  # Ledger/Tabloid
    "A4": {"width": 10.833, "height": 7.5},  # A4
    "A3": {"width": 14, "height": 10.5},  # A3
}
PPTX_NAMESPACE = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def inches_to_emu(x: float) -> int:
    return int(x * 914400)


def process(pptx_bytes: bytes, slide_size: str | None = None) -> bytes:
    """
    Process a PPTX file to apply slide size modifications.

    Args:
        pptx_bytes: The PPTX file content as bytes
        slide_size: The desired slide size (e.g., "16:9", "4:3", "A4")

    Returns:
        Modified PPTX content as bytes
    """
    prs = BytesIO(pptx_bytes)
    return _apply_slide_size(prs, slide_size)


def _apply_slide_size(prs: BytesIO, slide_size: str | None = None) -> bytes:  # type: ignore[valid-type]
    """
    Apply slide size to a presentation.

    Args:
        prs: The Presentation object as BytesIO buffer
        slide_size: The desired slide size (e.g., "16:9", "4:3", "A4")
    """
    # If no slide size specified, no modifications needed
    if slide_size is None:
        return prs.getvalue()

    # Normalize slide_size to uppercase for case-insensitive lookup
    slide_size_upper = slide_size.upper()

    # Get dimensions for the specified slide size
    if slide_size_upper not in SLIDE_SIZES:
        raise ValueError(f"Unsupported slide size: {slide_size}. Supported sizes: {', '.join(SLIDE_SIZES.keys())}")

    slide_dims = SLIDE_SIZES[slide_size_upper]
    # Convert inches to emu
    width = inches_to_emu(slide_dims["width"])
    height = inches_to_emu(slide_dims["height"])
    buf = BytesIO()
    with ZipFile(prs, "r") as zip_in, ZipFile(buf, "w", ZIP_DEFLATED) as zip_out:
        if not any(item.filename == "ppt/presentation.xml" for item in zip_in.infolist()):
            raise ValueError("Invalid pptx: presentation.xml missing")
        for item in zip_in.infolist():
            data = zip_in.read(item.filename)
            # Skip files that are not presentation.xml in zip folder
            if item.filename != "ppt/presentation.xml":
                zip_out.writestr(item, data)
                continue
            # Parse presentation.xml document
            tree = ElementTree.parse(BytesIO(data))
            root = tree.getroot()
            if root is None:
                continue
            # Find slide size element
            sld_sz = root.find("p:sldSz", PPTX_NAMESPACE)
            if sld_sz is not None:
                # Apply slide sizes
                sld_sz.set("cx", str(width))
                sld_sz.set("cy", str(height))
            data = ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)
            # Write to out buffer
            zip_out.writestr(item, data)

    # Sanitize user input for logging to prevent log injection (CWE-117)
    safe_slide_size = slide_size_upper.replace("\r\n", "").replace("\n", "")
    logging.debug(f'Applied slide size {safe_slide_size}: {slide_dims["width"]}" x {slide_dims["height"]}"')
    return buf.getvalue()
