"""
SVG processing utilities.

Features:
- Convert SVG <svg> to <img src="data:image/svg+xml;base64,...">
- Replace base64 SVG <img> with base64 PNG using Chromium via CDP (Chrome DevTools Protocol)
- Handle SVG dimensions, including vw/vh/% via viewBox

This is a port of the SvgProcessor from weasyprint-service, kept deliberately
close to its counterpart so the shared "SVG conversion" logic can later be
extracted into a common library.
"""

from __future__ import annotations

import base64
import logging
import math
import os
import re
import xml.etree.ElementTree as ET
from typing import TYPE_CHECKING, ClassVar

from bs4 import BeautifulSoup, Tag

# Defusedxml exposes ElementTree as a module.
from defusedxml import ElementTree as det  # noqa: N813

if TYPE_CHECKING:  # used only for type hints
    from xml.etree.ElementTree import Element

    from app.chromium_manager import ChromiumManager


class SvgProcessor:
    """
    Class for processing SVG images in HTML and converting them to PNG via CDP.
    """

    # MIME/constants
    SPECIAL_UNITS = ("vw", "vh", "%")

    # What a browser is asked to rasterize at most, after the scale factor: one side, and the pixels of it
    MAX_RENDER_SIDE = 10_000
    MAX_RENDER_PIXELS = 50_000_000

    # What a CSS px is worth in the units a length can be stated in, the absolute ones alone. A unit
    # which stands for something else on every element is not here: it cannot be read without that
    # element. The same table as `app/html_image_pre_process.py`, which sizes a raster by it.
    ABSOLUTE_UNITS_IN_PX: ClassVar[dict[str, float]] = {"px": 1.0, "in": 96.0, "cm": 96 / 2.54, "mm": 96 / 25.4, "pt": 96 / 72, "pc": 16.0}
    IMAGE_PNG = "image/png"
    IMAGE_SVG = "image/svg+xml"
    NON_SVG_CONTENT_TYPES = ("image/jpeg", "image/png", "image/gif")
    VIEWBOX_PARTS_COUNT = 4  # min-x, min-y, width, height
    DATA_PREFIX = "data:"
    SVG_NS = "http://www.w3.org/2000/svg"  # NOSONAR

    def __init__(
        self,
        chromium_manager: ChromiumManager | None = None,
        device_scale_factor: float | None = None,
        logger: logging.Logger | None = None,
    ) -> None:
        """
        Initialize SvgProcessor with CDP-based conversion.

        Args:
            chromium_manager: ChromiumManager instance for CDP-based conversion.
            device_scale_factor: Device scale factor for rendering. If None, reads DEVICE_SCALE_FACTOR (default 1.0).
            logger: Optional logger; if None, a module-level logger is used.
        """
        self.chromium_manager = chromium_manager
        self.device_scale_factor = self._parse_float(os.environ.get("DEVICE_SCALE_FACTOR"), 1.0) if device_scale_factor is None else float(device_scale_factor)
        self.log = logger or logging.getLogger(__name__)

    # ---------------- Public API ----------------

    async def process_svg(self, input_html: BeautifulSoup) -> BeautifulSoup:
        """
        Process <svg> and <img src="data:..."> in the HTML.
        - Replace only top-level <svg> with <img data:image/svg+xml;base64,...>
        - Convert base64 SVG images inside <img> to base64 PNG via CDP
        """
        self.log.info("Starting SVG processing in HTML")
        parsed_html = self.replace_inline_svgs_with_img(input_html)
        result = await self.replace_img_base64(parsed_html)
        self.log.info("Completed SVG processing")
        return result

    def replace_inline_svgs_with_img(self, parsed_html: BeautifulSoup) -> BeautifulSoup:
        """
        Replace only top-level <svg>...</svg> with <img src="data:image/svg+xml;base64,...">.
        Skips nested <svg> (those having an <svg> ancestor). Preserves width/height if present.
        """
        top_level_svgs: list[Tag] = [node for node in parsed_html.find_all("svg") if isinstance(node, Tag) and node.find_parent("svg") is None]

        self.log.debug("Found %d top-level SVG tags to replace with img tags", len(top_level_svgs))
        for svg in top_level_svgs:
            svg_str = str(svg)
            self.log.debug("Converting inline SVG to data URL, size: %d characters", len(svg_str))
            b64 = base64.b64encode(svg_str.encode("utf-8")).decode("ascii")
            img: Tag = parsed_html.new_tag("img")

            # Set attributes in the specific order expected by tests: height, src, width
            height = svg.get("height")
            if isinstance(height, str):
                img.attrs["height"] = height

            img.attrs["src"] = f"data:{self.IMAGE_SVG};base64,{b64}"

            width = svg.get("width")
            if isinstance(width, str):
                img.attrs["width"] = width

            svg.replace_with(img)

        return parsed_html

    async def replace_img_base64(self, parsed_html: BeautifulSoup) -> BeautifulSoup:
        """
        Replace base64 SVG images with PNG equivalents in HTML <img> tags via CDP.
        """
        img_nodes = parsed_html.find_all("img")
        self.log.debug("Found %d img tags to check for SVG data URLs", len(img_nodes))
        converted_count = 0
        for node in img_nodes:
            if not isinstance(node, Tag):
                continue

            src = self._get_attr_str(node, "src")
            parsed = self._parse_data_url_base64(src) if src else None
            if not parsed:
                continue

            content_type, content_base64 = parsed

            svg = self.get_svg(content_type, content_base64)
            if svg is None:
                continue

            drawn_size = self._drawn_size_px(node, svg)
            image_type, image_content = await self.replace_svg_with_png(svg, self._render_size_px(drawn_size))
            replaced_content_base64 = self.to_base64(image_content)

            # Skip if nothing changed
            if replaced_content_base64 == content_base64:
                continue

            # Preserve the rendered size by setting it explicitly on the <img>
            self._apply_img_dimensions(node, drawn_size)

            node["src"] = f"data:{image_type};base64,{replaced_content_base64}"
            converted_count += 1

        if converted_count > 0:
            self.log.info("Converted %d SVG data URLs to PNG", converted_count)
        return parsed_html

    def _apply_img_dimensions(self, node: Tag, drawn_size: tuple[int | None, int | None] | None) -> None:
        """Best-effort: give the <img> the size it is drawn at, as an attribute the writers read.

        The PNG carries the pixels of the SVG times the scale factor, so an image left to size itself
        would come out that many times too large. `drawn_size` is what `_drawn_size_px` read; None
        leaves the element as the document wrote it, the target being the one to resolve the size.

        The size is written as an attribute and not only as a style, because the writer of every target
        reads the attribute, while `filters/inline_styles.lua` carries a style onto it for docx alone.
        """
        try:
            if drawn_size is None:
                return
            width, height = drawn_size

            style_val = self._get_attr_str(node, "style") or ""
            style_parts = [part.strip() for part in style_val.split(";") if part.strip()]
            style_parts = [part for part in style_parts if not part.lower().startswith(("width:", "height:"))]

            if width is not None:
                node["width"] = f"{width}px"
                style_parts.append(f"width: {width}px")
            if height is not None:
                node["height"] = f"{height}px"
                style_parts.append(f"height: {height}px")

            node["style"] = "; ".join(style_parts)

        # Applying dimensions is best effort.
        except Exception as e:  # noqa: BLE001
            # Log at debug level to avoid noise but prevent silent pass
            logging.getLogger(__name__).debug("Failed to apply img dimensions from SVG: %s", e)

    def _drawn_size_px(self, node: Tag, svg: Element) -> tuple[int | None, int | None] | None:
        """The size this image is drawn at, in px, or None where only the target can resolve it.

        Either side may be None, meaning the raster carries it: a single side is enough for a writer,
        which keeps the ratio of the image it is given.

        Where the document states a size this can read, that size is the one drawn, brought inside
        `max-width` and `max-height`. A side the document states twice over - a width and a height
        both - is held by the cap on its own axis and by no other: the document has chosen the shape
        already, and a cap on one axis only trims that axis. A side the document leaves out follows
        the ratio of the SVG, which a viewBox is what makes it keep, and where a cap catches that
        side both sides give way together, so the drawing keeps its shape rather than sitting in a
        box of empty space. Where the document states nothing this can read, the width of the SVG is
        the size, as it always was, and a cap shrinks the whole of the drawing. Only a percentage is
        left alone: it is a share of a width no one here knows.
        """
        style = self._style_declarations(node)
        stated_width, stated_height = style.get("width"), style.get("height")
        if (stated_width or "").endswith("%") or (stated_height or "").endswith("%"):
            return None

        width, height = self._px_value(stated_width), self._px_value(stated_height)
        max_width, max_height = self._px_value(style.get("max-width")), self._px_value(style.get("max-height"))
        own_width, own_height, _ = self.extract_svg_dimensions_as_px(svg)

        if width is None and height is None:
            # Both sides follow the drawing, so a cap shrinks the whole of it, with the ratio it had
            if not isinstance(own_width, int):
                return None
            return self._whole_drawing_inside_the_caps(own_width, own_height, max_width, max_height)

        width, height = self._inside_its_own_cap(width, max_width), self._inside_its_own_cap(height, max_height)
        ratio = self._ratio_of(svg, own_width, own_height)

        # The side the document leaves out follows the one it states, where there is a ratio to follow
        if ratio is not None and width is None and height is not None:
            return self._both_sides_giving_way(max(1, math.ceil(height / ratio)), height, max_width)
        if ratio is not None and height is None and width is not None:
            height, width = self._both_sides_giving_way(max(1, math.ceil(width * ratio)), width, max_height)

        return width, height

    def _ratio_of(self, svg: Element, own_width: int | None, own_height: int | None) -> float | None:
        """The height of the drawing over its width, or None where it has none to be scaled by.

        The drawing is scaled into the size asked for, which a viewBox is what makes possible. Without
        one there is no ratio to complete a single side with, and the raster carries it instead.
        """
        if not own_width or not own_height or self.parse_viewbox(svg) == (None, None):
            return None
        return own_height / own_width

    @staticmethod
    def _inside_its_own_cap(side: int | None, cap: int | None) -> int | None:
        """A side held by the cap on that same axis, which is the only one to hold it."""
        return side if side is None or cap is None else min(side, cap)

    @staticmethod
    def _both_sides_giving_way(following: int, stated: int, cap: int | None) -> tuple[int, int]:
        """The pair where a cap catches the side which follows the drawing, and the stated side with it.

        The document states the one side and leaves the other to the drawing, so the shape is the
        drawing's to keep. Holding the following side alone would leave the image in a box wider than
        itself, with the drawing letterboxed inside it.
        """
        if cap is None or following <= cap:
            return following, stated
        return cap, max(1, round(stated * cap / following))

    @staticmethod
    def _whole_drawing_inside_the_caps(own_width: int, own_height: int | None, max_width: int | None, max_height: int | None) -> tuple[int, None]:
        """The width of a drawing both of whose sides follow it, brought inside the caps it is given.

        The height is left to the raster, which carries the ratio of the SVG, so a cap on it is met by
        shrinking the width in the same measure.
        """
        factor = 1.0
        for cap, side in ((max_width, own_width), (max_height, own_height)):
            if cap is not None and side is not None and side > cap:
                factor = min(factor, cap / side)
        return (own_width if math.isclose(factor, 1.0) else max(1, round(own_width * factor))), None

    def _render_size_px(self, drawn_size: tuple[int | None, int | None] | None) -> tuple[int, int] | None:
        """The size to rasterize at, or None to rasterize the SVG at its own size.

        The PNG is made at the size the image is drawn at times the scale factor, which is what keeps
        an image sharp in print. Without it an image the document enlarges would be rasterized at the
        size of its SVG and blown up from there.
        """
        if drawn_size is None:
            return None
        width, height = drawn_size
        if width is None or height is None:
            return None

        # A document could otherwise size an image into a screenshot of any size, and the memory of the
        # browser taking it is shared with every conversion running beside this one. The image is still
        # drawn at the size the document gives; only its detail falls back to the size of the SVG.
        scale = self.device_scale_factor or 1.0
        if max(width, height) * scale > self.MAX_RENDER_SIDE or width * height * scale * scale > self.MAX_RENDER_PIXELS:
            self.log.warning("Rasterizing an SVG at its own size: %dx%d px at a scale of %.2f is too large", width, height, scale)
            return None
        return width, height

    @classmethod
    def _px_value(cls, value: str | None) -> int | None:
        """A CSS length in px, where the unit it is stated in says how long it is on its own.

        A unit which stands for something else on every element - a percentage, `em`, `vw` - reads as
        none: it cannot be resolved without the layout the target has and this does not. A bare number
        is px, which is how pandoc and `filters/inline_styles.lua` both read one.
        """
        if value is None:
            return None
        # The space around a value is stripped rather than matched: two `\s*` on either side of a unit
        # which may be empty is a run of spaces this could divide in as many ways as it is long. CSS
        # puts no space between a number and its unit, so none is read between them either.
        # The two ways of writing a number start on different characters, so neither can be read as the
        # other and nothing is left to backtrack over
        match = re.fullmatch(r"(\d+(?:\.\d+)?|\.\d+)([a-z]*)", value.strip())
        if match is None:
            return None
        factor = cls.ABSOLUTE_UNITS_IN_PX.get(match.group(2) or "px")
        if factor is None:
            return None
        length = float(match.group(1)) * factor
        # A number long enough to reach infinity would raise on the way to an int
        return math.ceil(length) if math.isfinite(length) and length > 0 else None

    def _style_declarations(self, node: Tag) -> dict[str, str]:
        """The inline style of the element, as property to value, lowercased.

        A later declaration wins, unless an earlier one is `!important` - the cascade of a style attribute.
        """
        style_val = (self._get_attr_str(node, "style") or "").lower()
        declarations: dict[str, str] = {}
        important: set[str] = set()
        for part in style_val.split(";"):
            name, separator, value = part.partition(":")
            if not separator:
                continue
            name = name.strip()
            value = value.strip()
            marked = value.endswith("!important")
            if marked:
                value = value[: -len("!important")].strip()
            if name in important and not marked:
                continue
            declarations[name] = value
            if marked:
                important.add(name)
        return declarations

    # ---------------- Core helpers ----------------

    def _parse_data_url_base64(self, src: str | None) -> tuple[str, str] | None:
        """
        Parse a data URL of the form "data:<content-type>;base64,<payload>".
        Returns a tuple (content_type, base64_payload) or None if not applicable.
        """
        if not src or not src.startswith(self.DATA_PREFIX) or ";base64," not in src:
            return None
        header, b64data = src.split(";base64,", 1)
        if not header.startswith(self.DATA_PREFIX):
            return None
        content_type = header[len(self.DATA_PREFIX) :]
        return content_type, b64data

    def get_svg(self, content_type: str, content_base64: str) -> Element | None:
        """
        Decode and validate base64 content as SVG. Allows incorrect MIME types (common in the wild).
        """
        if content_type in self.NON_SVG_CONTENT_TYPES:
            self.log.debug("Skipping non-SVG content type: %s", content_type)
            return None

        try:
            decoded_content = base64.b64decode(content_base64)
            if b"\0" in decoded_content:
                self.log.debug("Skipping binary content (contains null bytes)")
                return None

            possible_svg_content = decoded_content.decode("utf-8")
            return self.svg_from_string(possible_svg_content)
        except Exception as e:  # noqa: BLE001
            self.log.error("Failed to decode base64 content: %s", e)
            return None

    async def replace_svg_with_png(self, svg: Element, render_size: tuple[int, int] | None = None) -> tuple[str, str | bytes]:
        """
        Convert SVG Element to PNG bytes using CDP.
        Returns tuple of (mime, content). If conversion fails, returns original SVG.

        `render_size` is the size the document draws the image at, where it gives one: the PNG is rasterized
        at that size times the scale factor, so it stays sharp wherever the image is enlarged.
        """
        updated_svg = self.ensure_mandatory_attributes(svg)

        width, height, updated_svg = self.extract_svg_dimensions_as_px(updated_svg)
        if not width or not height:
            self.log.warning("Invalid or undefined dimensions for SVG (width: %s, height: %s)", width, height)
            return self.without_changes(svg)

        if render_size is not None:
            width, height = render_size
            updated_svg = self.replace_svg_size_attributes(updated_svg, width, height)
        self.log.debug("Converting SVG (%dx%d px) to PNG with scale factor %.2f", width, height, self.device_scale_factor)

        svg_content = self.svg_to_string(updated_svg)

        # Convert via CDP
        if self.chromium_manager:
            try:
                png_bytes = await self.chromium_manager.convert_svg_to_png(svg_content, width, height, self.device_scale_factor)
                self.log.debug("SVG converted via CDP successfully")
            except Exception as e:  # noqa: BLE001
                self.log.error("CDP conversion failed: %s", e)
                return self.without_changes(svg)
            else:
                return self.IMAGE_PNG, png_bytes
        else:
            self.log.error("No ChromiumManager available, returning original SVG")
            return self.without_changes(svg)

    def ensure_mandatory_attributes(self, svg: Element) -> Element:
        # Ensure required XML namespace exists and non-empty
        if not svg.tag.startswith("{"):
            svg.tag = f"{{{SvgProcessor.SVG_NS}}}svg"
        return svg

    def without_changes(self, svg: Element) -> tuple[str, str | bytes]:
        return self.IMAGE_SVG, self.svg_to_string(svg)

    # ---------------- XML / SVG utilities ----------------

    def svg_from_string(self, content: str) -> Element | None:
        try:
            return det.fromstring(content)
        except det.ParseError as e:
            self.log.error("Failed to parse SVG content: %s", e)
            return None

    @staticmethod
    def svg_to_string(svg: Element) -> str:
        ET.register_namespace("", SvgProcessor.SVG_NS)  # NOSONAR
        return ET.tostring(svg, encoding="unicode")

    # ---------------- Dimension parsing/conversion ----------------

    def extract_svg_dimensions_as_px(self, svg: Element) -> tuple[int | None, int | None, Element]:
        """
        Read width/height, convert to px. If units are relative (vw/vh/%), use viewBox where possible.
        If viewBox present and one/both dimensions missing, set explicit px attributes on the SVG.
        """
        width, width_unit = self.get_svg_dimension(svg, "width")
        height, height_unit = self.get_svg_dimension(svg, "height")
        vb_width, vb_height = self.parse_viewbox(svg)

        width_px = self.calculate_dimension(width, width_unit, vb_width)
        height_px = self.calculate_dimension(height, height_unit, vb_height)

        if vb_width is not None and vb_height is not None:
            if width_px is None:
                width_px = math.ceil(vb_width)
            if height_px is None:
                height_px = math.ceil(vb_height)
            svg = self.replace_svg_size_attributes(svg, width_px, height_px)

        if width_px is None or height_px is None:
            return None, None, svg

        return width_px, height_px, svg

    @staticmethod
    def get_svg_dimension(svg: Element, dimension: str) -> tuple[str | None, str | None]:
        value = svg.attrib.get(dimension)
        if value is None:
            return None, None

        match = re.search(
            r"^(?P<value>-?\d+(?:\.\d+)?)(?P<unit>[a-z%]+)?$",
            value,
            flags=re.IGNORECASE,
        )
        if match:
            return match.group("value"), match.group("unit")
        return None, None

    @staticmethod
    def parse_viewbox(svg: Element) -> tuple[float | None, float | None]:
        viewbox = svg.attrib.get("viewBox")
        if viewbox is None:
            return None, None

        # SVG viewBox format: "min-x min-y width height" with spaces and/or commas.
        try:
            # Normalize commas to spaces and split on whitespace
            parts = viewbox.replace(",", " ").split()
            if len(parts) != SvgProcessor.VIEWBOX_PARTS_COUNT:
                return None, None
            vb_width = float(parts[2])
            vb_height = float(parts[3])
        except Exception:  # noqa: BLE001
            return None, None
        else:
            return vb_width, vb_height

    def calculate_dimension(
        self,
        value: str | None,
        unit: str | None,
        vb_dimension: float | None,
    ) -> int | None:
        if value is None:
            return None

        if unit in self.SPECIAL_UNITS:
            if vb_dimension is None:
                raise ValueError(f"{unit} units require a viewBox to be defined")
            return self.calculate_special_unit(value, unit, vb_dimension)

        return self.convert_to_px(value, unit)

    @staticmethod
    def replace_svg_size_attributes(svg: Element, width_px: int, height_px: int) -> Element:
        svg.set("width", f"{width_px}px")
        svg.set("height", f"{height_px}px")
        return svg

    def calculate_special_unit(self, value: str, unit: str | None, viewbox_dimension: float) -> int:
        try:
            val = float(value)
        except (ValueError, TypeError) as err:
            raise ValueError(f"could not convert string to float: '{value}'") from err

        if unit in self.SPECIAL_UNITS:
            return math.ceil((val / 100) * viewbox_dimension)

        fallback = self.convert_to_px(value, unit)
        if fallback is None:
            raise ValueError(f"Cannot convert unit '{unit}' to px")
        return fallback

    # ---------------- Generic helpers ----------------

    @staticmethod
    def to_base64(content: str | bytes) -> str:
        if isinstance(content, str):
            content = content.encode("utf-8")
        return base64.b64encode(content).decode("utf-8")

    def convert_to_px(self, value: str | None, unit: str | None) -> int | None:
        try:
            if value is None:
                # The try converts, this guard rejects a missing value.
                raise ValueError  # noqa: TRY301
            value_f64 = float(value)

            if unit in self.SPECIAL_UNITS:
                return None

            return math.ceil(value_f64 * self.get_px_conversion_ratio(unit))
        except ValueError:
            self.log.error("Invalid value for conversion: %s", value)
            return None

    @staticmethod
    def get_px_conversion_ratio(unit: str | None) -> float:
        """
        Convert CSS units to px at 96 DPI.
        Note: Device scale factor should NOT affect layout dimensions; it only controls rasterization DPI.
        """
        if not unit:
            return 1.0
        return {
            "px": 1.0,
            "pt": 4 / 3,
            "in": 96.0,
            "cm": 96 / 2.54,
            "mm": 96 / 2.54 / 10,
            "pc": 16.0,
            "ex": 8.0,
        }.get(unit, 1.0)

    @staticmethod
    def _get_attr_str(tag: Tag, name: str) -> str | None:
        val = tag.get(name)
        return val if isinstance(val, str) else None

    @staticmethod
    def _parse_float(value: str | None, default: float) -> float:
        try:
            return float(value) if value is not None else default
        except ValueError:
            return default
