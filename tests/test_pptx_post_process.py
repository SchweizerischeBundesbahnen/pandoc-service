import io
from defusedxml import ElementTree
from zipfile import ZipFile
import re

import pytest

from app.PptxPostProcess import SLIDE_SIZES, _apply_slide_size, process, PPTX_NAMESPACE, inches_to_emu

SLIDE_PATH_PATTERN = re.compile(r"ppt/slides/[a-zA-Z0-9]+.xml")


def create_test_pptx() -> bytes:
    """
    Read a minimal test PPTX file.
    Test PPTX includes one slide and one text box reading "Test Content"

    """
    with open("tests/default.pptx", "rb") as default_pres:
        return default_pres.read()


def find_presentation_information(prs: io.BytesIO) -> tuple[int, int, int]:
    """Read slide dimensions from presentation.xml"""
    with ZipFile(prs, "r") as zip_in:
        for item in zip_in.infolist():
            if item.filename != "ppt/presentation.xml":
                continue
            # Read data
            data = zip_in.read(item.filename)
            tree = ElementTree.parse(io.BytesIO(data))
            root = tree.getroot()
            # Find slide size element
            sld_sz = root.find("p:sldSz", PPTX_NAMESPACE)
            sld_id_lst = list(root.find("p:sldIdLst", PPTX_NAMESPACE).iter())[1:]  # Remove first element since list includes sld_id_lst
            num_slides = len(sld_id_lst)
            return (int(sld_sz.get("cx")), int(sld_sz.get("cy")), num_slides)
        raise ValueError("Invalid Presentation")


def find_first_text_box_content(prs: io.BytesIO) -> str:
    """Read and return the text content of the first text box in the first slide"""
    with ZipFile(prs, "r") as zip_in:
        for item in zip_in.infolist():
            if re.match(SLIDE_PATH_PATTERN, item.filename) is None:
                continue
            # Read data
            data = zip_in.read(item.filename)
            tree = ElementTree.parse(io.BytesIO(data))
            root = tree.getroot()
            # Find first text box element
            text_box = root.find(".//a:t", PPTX_NAMESPACE)  # .// XPath prefix for recursive search
            return text_box.text
        raise ValueError("Invalid Presentation")


def test_process_with_no_slide_size():
    """Test processing PPTX without specifying slide size (no modifications)."""
    pptx_bytes = create_test_pptx()
    original_width, original_height, _ = find_presentation_information(io.BytesIO(pptx_bytes))

    result_bytes = process(pptx_bytes, slide_size=None)
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    # Dimensions should remain unchanged
    assert result_width == original_width
    assert result_height == original_height


def test_process_with_16_9_slide_size():
    """Test processing PPTX with 16:9 slide size."""
    pptx_bytes = create_test_pptx()

    result_bytes = process(pptx_bytes, slide_size="16:9")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["16:9"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["16:9"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_process_with_4_3_slide_size():
    """Test processing PPTX with 4:3 slide size."""
    pptx_bytes = create_test_pptx()

    result_bytes = process(pptx_bytes, slide_size="4:3")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["4:3"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["4:3"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_process_with_case_insensitive_slide_size():
    """Test that slide size is case insensitive."""
    pptx_bytes = create_test_pptx()

    # Test lowercase
    result_bytes = process(pptx_bytes, slide_size="a4")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["A4"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["A4"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_process_with_invalid_slide_size():
    """Test processing PPTX with invalid slide size raises ValueError."""
    pptx_bytes = create_test_pptx()

    with pytest.raises(ValueError, match="Unsupported slide size: INVALID"):
        process(pptx_bytes, slide_size="INVALID")


def test_process_with_letter_slide_size():
    """Test processing PPTX with LETTER slide size."""
    pptx_bytes = create_test_pptx()

    result_bytes = process(pptx_bytes, slide_size="LETTER")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["LETTER"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["LETTER"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_process_with_a3_slide_size():
    """Test processing PPTX with A3 slide size."""
    pptx_bytes = create_test_pptx()

    result_bytes = process(pptx_bytes, slide_size="A3")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["A3"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["A3"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_process_with_widescreen_slide_size():
    """Test processing PPTX with WIDESCREEN slide size."""
    pptx_bytes = create_test_pptx()

    result_bytes = process(pptx_bytes, slide_size="WIDESCREEN")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["WIDESCREEN"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["WIDESCREEN"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_apply_slide_size_none():
    """Test _apply_slide_size with None does nothing."""
    prs = create_test_pptx()
    buf = io.BytesIO(prs)
    original_width, original_height, _ = find_presentation_information(buf)

    result_bytes = _apply_slide_size(buf, slide_size=None)
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    # Dimensions should remain unchanged
    assert result_width == original_width
    assert result_height == original_height


def test_apply_slide_size_with_valid_size():
    """Test _apply_slide_size with valid slide size."""
    prs = create_test_pptx()
    buf = io.BytesIO(prs)
    result_bytes = _apply_slide_size(buf, slide_size="16:9")
    result_width, result_height, _ = find_presentation_information(io.BytesIO(result_bytes))

    expected_width = inches_to_emu(SLIDE_SIZES["16:9"]["width"])
    expected_height = inches_to_emu(SLIDE_SIZES["16:9"]["height"])

    assert result_width == expected_width
    assert result_height == expected_height


def test_apply_slide_size_with_invalid_size():
    """Test _apply_slide_size with invalid slide size raises ValueError."""
    prs = create_test_pptx()

    with pytest.raises(ValueError, match="Unsupported slide size: BADSIZE"):
        _apply_slide_size(io.BytesIO(prs), slide_size="BADSIZE")


def test_slide_sizes_constant():
    """Test that SLIDE_SIZES constant has expected entries."""
    expected_sizes = ["16:9", "WIDESCREEN", "4:3", "LETTER", "LEDGER", "A4", "A3"]

    for size in expected_sizes:
        assert size in SLIDE_SIZES
        assert "width" in SLIDE_SIZES[size]
        assert "height" in SLIDE_SIZES[size]
        assert isinstance(SLIDE_SIZES[size]["width"], (int, float))
        assert isinstance(SLIDE_SIZES[size]["height"], (int, float))


def test_process_returns_bytes():
    """Test that process returns bytes."""
    pptx_bytes = create_test_pptx()

    result = process(pptx_bytes, slide_size="16:9")

    assert isinstance(result, bytes)
    assert len(result) > 0


def test_process_preserves_content():
    """Test that processing preserves presentation content."""
    prs = create_test_pptx()

    # Process the PPTX
    result_bytes = process(prs, slide_size="16:9")
    result_width, result_height, num_slides = find_presentation_information(io.BytesIO(result_bytes))

    # Check width and height
    assert result_width == inches_to_emu(SLIDE_SIZES["16:9"]["width"])
    assert result_height == inches_to_emu(SLIDE_SIZES["16:9"]["height"])
    # Check that content is preserved
    assert num_slides == 1
    # Find the textbox in the result
    text_box_content = find_first_text_box_content(io.BytesIO(result_bytes))
    assert text_box_content == "Test Content"
