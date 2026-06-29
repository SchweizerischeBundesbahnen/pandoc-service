from __future__ import annotations

import contextlib
import io
import logging
import os
import platform
import subprocess
import tempfile
import time
from http import HTTPStatus
from pathlib import Path
from typing import TYPE_CHECKING

import anyio
import starlette.datastructures
import uvicorn
from bs4 import BeautifulSoup
from fastapi import FastAPI, Request, Response
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse, PlainTextResponse, StreamingResponse
from prometheus_fastapi_instrumentator import Instrumentator

from app.schema import VersionSchema

from . import DocxColorPreProcess, DocxListLevelPreProcess, DocxParagraphPreProcess, DocxPostProcess, HtmlListsPreProcess, HtmlParagraphPreProcess, PptxPostProcess
from .chromium_manager import get_chromium_manager
from .metrics_server import MetricsServer, get_metrics_port, is_metrics_server_enabled
from .pandoc_metrics import get_pandoc_metrics
from .prometheus_metrics import (
    increment_conversion_failure,
    increment_conversion_success,
    increment_template_conversion,
    initialize_pandoc_info,
    observe_post_processing_duration,
    observe_request_body_size,
    observe_response_body_size,
    observe_subprocess_duration,
)
from .svg_processor import SvgProcessor

if TYPE_CHECKING:
    from collections.abc import AsyncGenerator, Awaitable, Callable


CUSTOM_REFERENCE_DOCX = "custom-reference.docx"
CUSTOM_REFERENCE_PPTX = "custom-reference.pptx"
PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_BASE_PATH = "/usr/local/share/pandoc/filters"

FILTERS = {
    "page_break": f"{FILTER_BASE_PATH}/pagebreak.lua",
    "page_orientation": f"{FILTER_BASE_PATH}/page_orientation.lua",
    "heading_levels": f"{FILTER_BASE_PATH}/heading_levels.lua",
    "inline_styles": f"{FILTER_BASE_PATH}/inline_styles.lua",
    "docx_colors_to_latex": f"{FILTER_BASE_PATH}/docx_colors_to_latex.lua",
    "docx_paragraphs_to_latex": f"{FILTER_BASE_PATH}/docx_paragraphs_to_latex.lua",
    "docx_lists_to_latex": f"{FILTER_BASE_PATH}/docx_lists_to_latex.lua",
    "html_lists": f"{FILTER_BASE_PATH}/html_lists.lua",
}

# List of allowed pandoc options for security
ALLOWED_PANDOC_OPTIONS = [
    f"--lua-filter={FILTERS['page_break']}",
    f"--lua-filter={FILTERS['page_orientation']}",
    f"--lua-filter={FILTERS['heading_levels']}",
    f"--lua-filter={FILTERS['inline_styles']}",
    f"--lua-filter={FILTERS['docx_colors_to_latex']}",
    f"--lua-filter={FILTERS['docx_paragraphs_to_latex']}",
    f"--lua-filter={FILTERS['docx_lists_to_latex']}",
    f"--lua-filter={FILTERS['html_lists']}",
    "--track-changes=all",
    "--reference-doc=",  # Prefix for reference-doc option
    "--pdf-engine=tectonic",
    "--toc",
]

# Target formats whose writer ultimately produces LaTeX (PDF goes through
# tectonic, latex is the raw .tex file). The DOCX color preprocessor only
# helps for these targets — for DOCX -> DOCX/HTML/etc. we leave the input
# alone.
_LATEX_TARGET_FORMATS = frozenset({"pdf", "latex"})

# Add other allowed formats as needed
ALLOWED_SOURCE_FORMATS = ["docx", "epub", "fb2", "html", "json", "latex", "markdown", "rtf", "textile"]
ALLOWED_TARGET_FORMATS = ["docx", "epub", "fb2", "html", "json", "latex", "markdown", "odt", "pdf", "plain", "pptx", "rtf", "textile"]

MIME_TYPES = {
    "html": "text/html",
    "html5": "text/html",
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "odt": "application/vnd.oasis.opendocument.text",
    "epub": "application/epub+zip",
    "markdown": "text/markdown",
    "md": "text/markdown",
    "latex": "application/x-latex",
    "tex": "application/x-tex",
    "rtf": "application/rtf",
    "txt": "text/plain",
    "json": "application/json",
    "xml": "application/xml",
}

DEFAULT_MIME_TYPE = "application/octet-stream"

FILE_EXTENSIONS = {
    "html": "html",
    "html5": "html",
    "pdf": "pdf",
    "docx": "docx",
    "odt": "odt",
    "epub": "epub",
    "markdown": "md",
    "md": "md",
    "latex": "tex",
    "tex": "tex",
    "rtf": "rtf",
    "txt": "txt",
    "json": "json",
    "xml": "xml",
    "asciidoc": "adoc",
    "rst": "rst",
    "org": "org",
    "revealjs": "html",
    "beamer": "pdf",
    "context": "tex",
    "textile": "textile",
    "dokuwiki": "txt",
    "mediawiki": "wiki",
    "man": "man",
    "ms": "ms",
    "pptx": "pptx",
    "plain": "txt",
}

# Build default options dynamically
DEFAULT_CONVERSION_OPTIONS = ["--track-changes=all", f"--lua-filter={FILTERS['page_break']}", f"--lua-filter={FILTERS['heading_levels']}"]

logger = logging.getLogger(__name__)


async def _start_chromium() -> None:
    """Start the persistent Chromium browser used to rasterize embedded SVGs.

    Best effort: if it fails to start, conversions still run (SVGs just pass
    through unrasterized) so a missing browser never takes the service down.
    """
    if not is_svg_conversion_enabled():
        return
    chromium_manager = get_chromium_manager()
    try:
        await chromium_manager.start()
        logger.info("Chromium started for SVG conversion (version: %s)", chromium_manager.get_version())
    except Exception as e:  # noqa: BLE001
        logger.exception("Failed to start Chromium for SVG conversion; SVG rasterization disabled: %s", e)


async def _stop_chromium() -> None:
    """Stop the persistent Chromium browser if it is running."""
    if not is_svg_conversion_enabled():
        return
    chromium_manager = get_chromium_manager()
    if chromium_manager.is_running():
        try:
            await chromium_manager.stop()
        except Exception as e:  # noqa: BLE001
            logger.exception("Error stopping Chromium: %s", e)


@contextlib.asynccontextmanager
async def lifespan(app_instance: FastAPI) -> AsyncGenerator[None]:  # noqa: ARG001
    """
    Manage the lifecycle of the metrics server.

    This ensures the metrics server is started when the FastAPI application
    starts and properly cleaned up on shutdown.

    The metrics server is started on a dedicated port (default: 9182) for
    security isolation from the main application API.
    """
    # Initialize metrics and cache pandoc version
    pandoc_metrics = get_pandoc_metrics()
    pandoc_version = get_pandoc_version()
    pandoc_metrics.set_pandoc_version(pandoc_version)
    logger.info("Pandoc version: %s", pandoc_version)

    # Initialize Prometheus info metric once at startup
    # Guard against repeated lifespan execution (e.g., uvicorn --reload, TestClient)
    service_version = os.environ.get("PANDOC_SERVICE_VERSION", "unknown")
    try:
        initialize_pandoc_info(pandoc_version or "unknown", service_version)
    except ValueError:
        logger.debug("Prometheus info metric already initialized (lifespan re-executed)")

    # Start metrics server if enabled
    metrics_server: MetricsServer | None = None
    if is_metrics_server_enabled():
        metrics_port = get_metrics_port()
        metrics_server = MetricsServer(port=metrics_port)
        try:
            await metrics_server.start()
        except Exception as e:  # noqa: BLE001
            logger.error("Failed to start metrics server: %s", e)
            metrics_server = None

    # Start the persistent Chromium browser used to rasterize embedded SVGs.
    await _start_chromium()

    yield  # Application runs here

    await _stop_chromium()

    # Stop metrics server
    if metrics_server:
        try:
            await metrics_server.stop()
        except Exception as e:  # noqa: BLE001
            logger.error("Error stopping metrics server: %s", e)


app = FastAPI(
    openapi_url="/static/openapi.json",
    docs_url="/api/docs",
    lifespan=lifespan,
)

# Initialize Prometheus Instrumentator for automatic HTTP metrics
# Note: We instrument but don't expose() - metrics are served on a dedicated port via metrics_server
Instrumentator(
    should_group_status_codes=False,
    should_ignore_untemplated=True,
    should_respect_env_var=True,
    should_instrument_requests_inprogress=True,
    env_var_name="ENABLE_METRICS",
    inprogress_name="http_requests_inprogress",
    inprogress_labels=True,
).instrument(app)


# Validate REQUEST_BODY_LIMIT_MB environment variable
def get_request_body_limit_mb() -> int:
    default_limit_mb = 500
    env_value = os.environ.get("REQUEST_BODY_LIMIT_MB", str(default_limit_mb))
    try:
        value = int(env_value)
        if value <= 0:
            logger.warning(f"REQUEST_BODY_LIMIT_MB value '{env_value}' is not positive. Using default {default_limit_mb} MB.")
            value = default_limit_mb
    except ValueError:
        logger.warning(f"REQUEST_BODY_LIMIT_MB value '{env_value}' is not a valid integer. Using default {default_limit_mb} MB.")
        value = default_limit_mb
    return value


env_data_limit = get_request_body_limit_mb()
data_limit = env_data_limit * 1024 * 1024  # Convert MB to bytes


# Set the maximum request body size to data_limit
@app.middleware("http")
async def check_request_size(request: Request, call_next: Callable[[Request], Awaitable[Response]]) -> Response:
    size = len(await request.body())

    if size > data_limit:
        return process_error(
            Exception(f"Body Size {size} > {data_limit}"),
            "Request Body too large",
            413,
        )

    response = await call_next(request)
    return response


# Replace standard request validation error handler to adhere to the PlainText format
@app.exception_handler(RequestValidationError)
async def handle_validation_error(request: Request, exc: RequestValidationError) -> PlainTextResponse:
    return process_error(
        exc,
        "Validation Error",
        422,
    )


def get_pandoc_version() -> str | None:
    """Get the pandoc version using subprocess."""
    try:
        result = subprocess.run(
            [PANDOC_PATH, "--version"],
            capture_output=True,
            text=True,
            check=True,
            shell=False,
        )
        # Extract version from the first line of output
        version_line = result.stdout.splitlines()[0]
        return version_line.split()[-1]  # Last word on the first line is the version
    except Exception as e:
        logger.error(f"Error getting pandoc version: {e}")
        return None


def get_tectonic_availability() -> str:
    try:
        subprocess.run(
            ["/usr/bin/tectonic", "--version"],
            capture_output=True,
            check=True,
        )
        return "available"
    except (FileNotFoundError, subprocess.CalledProcessError) as e:
        logger.warning(f"Tectonic check failed: {e}")
        return "unavailable"
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Tectonic check error: {e}")
        return "unknown"


def get_temp_directory_writability() -> str:
    try:
        with tempfile.NamedTemporaryFile("w") as probe_file:
            probe_file.write("ok")
            return "writable"
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Temp directory is not writable: {e}")
        return "unwritable"


@app.get(
    "/version",
    summary="Get service version information",
    description="Get version information for python, pandoc executable and pandoc service",
    responses={200: {"description": "Success", "content": {MIME_TYPES["txt"]: {}}}, 422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}}},
)
def version() -> VersionSchema:
    return VersionSchema(
        python=platform.python_version(),
        pandoc=get_pandoc_version(),
        pandocService=os.environ.get("PANDOC_SERVICE_VERSION"),
        timestamp=os.environ.get("PANDOC_SERVICE_BUILD_TIMESTAMP"),
        chromium=get_chromium_manager().get_version(),
    )


@app.get(
    "/health",
    summary="Health check endpoint",
    description="Returns service health status",
    operation_id="healthCheck",
    tags=["meta"],
    responses={
        200: {
            "description": "Service healthy",
            "content": {MIME_TYPES["json"]: {}},
        },
        503: {"description": "Service unavailable", "content": {MIME_TYPES["json"]: {}}},
    },
)
def health() -> JSONResponse:
    """
    Health check endpoint for monitoring.

    Returns:
        JSONResponse: {status, pandoc, tectonic, filesystem}
    """
    logger.debug("Health check endpoint called")

    # Basic health check - service is running
    # Note: "chromium" is informational only (values never match the unhealthy
    # trip words below) because SVG rasterization is a best-effort enhancement.
    health_status = {"status": "healthy", "pandoc": "available" if get_pandoc_version() else "unavailable", "tectonic": get_tectonic_availability(), "filesystem": get_temp_directory_writability(), "chromium": get_chromium_health()}

    # Optional: Add dependency checks here
    # Memory/resource status

    if any(key in ["unavailable", "unwritable", "unknown"] for key in health_status.values()):
        health_status["status"] = "unhealthy"
        return JSONResponse(health_status, status_code=503)

    return JSONResponse(health_status)


@app.get(
    "/docx-template",
    summary="Download DOCX template",
    description="Get the default DOCX template for document conversion",
    responses={
        200: {
            "description": "Success",
            "content": {MIME_TYPES["docx"]: {}},
        },
        422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}},
        500: {"description": "Internal server error while generating the template.", "content": {MIME_TYPES["txt"]: {}}},
    },
)
async def get_docx_template():  # type: ignore
    path = Path(CUSTOM_REFERENCE_DOCX)
    try:
        # ruff: noqa: S603
        proc = await anyio.run_process(
            [
                PANDOC_PATH,
                "-o",
                "custom-reference.docx",
                "--print-default-data-file",
                "reference.docx",
            ],
            check=True,
        )
        if proc.returncode != 0:
            return process_error(Exception(f"Process failed with return code {proc.returncode}"), "An internal error has occurred while generating the template", 500)

        async with await anyio.open_file(path, "rb") as f:
            doc_content = await f.read()

        response = StreamingResponse(
            io.BytesIO(doc_content),
            headers={"Content-Disposition": "attachment; filename=reference.docx"},
            media_type=MIME_TYPES["docx"],
        )
        return response
    finally:
        if Path.exists(path):
            Path.unlink(path)


@app.get(
    "/pptx-template",
    summary="Download PPTX template",
    description="Get the default PPTX template for presentation conversion",
    responses={
        200: {
            "description": "Success",
            "content": {MIME_TYPES["pptx"]: {}},
        },
        422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}},
        500: {"description": "Internal server error while generating the template.", "content": {MIME_TYPES["txt"]: {}}},
    },
)
async def get_pptx_template():  # type: ignore
    path = Path(CUSTOM_REFERENCE_PPTX)
    try:
        # ruff: noqa: S603
        proc = await anyio.run_process(
            [
                PANDOC_PATH,
                "-o",
                "custom-reference.pptx",
                "--print-default-data-file",
                "reference.pptx",
            ],
            check=True,
        )
        if proc.returncode != 0:
            return process_error(Exception(f"Process failed with return code {proc.returncode}"), "An internal error has occurred while generating the template", 500)

        async with await anyio.open_file(path, "rb") as f:
            pptx_content = await f.read()

        response = StreamingResponse(
            io.BytesIO(pptx_content),
            headers={"Content-Disposition": "attachment; filename=reference.pptx"},
            media_type=MIME_TYPES["pptx"],
        )
        return response
    finally:
        if Path.exists(path):
            Path.unlink(path)


def _validate_pandoc_options(options: list[str]) -> list[str]:
    """
    Validate pandoc options against the whitelist to prevent command injection.

    Args:
        options: List of pandoc options to validate

    Returns:
        List of validated options

    Raises:
        ValueError: If any option is not in the whitelist
    """
    validated_options = []
    for option in options:
        is_valid = False
        # Check exact match
        if option in ALLOWED_PANDOC_OPTIONS:
            is_valid = True
        # Check prefix match (for options like --reference-doc=filename.docx)
        else:
            for allowed_prefix in ALLOWED_PANDOC_OPTIONS:
                if allowed_prefix.endswith("=") and option.startswith(allowed_prefix):
                    is_valid = True
                    break

        if not is_valid:
            raise ValueError(f"Invalid pandoc option: {option}")

        validated_options.append(option)

    return validated_options


def _build_pandoc_command(  # noqa: PLR0913
    *,
    source_format: str,
    target_format: str,
    source_path: str,
    output_path: str,
    validated_options: list[str],
    apply_docx_latex_filters: bool,
    preserve_table_styles: bool = False,
) -> list[str]:
    """Build the pandoc CLI invocation for run_pandoc_conversion."""
    # Source format gains the +styles extension on the docx->latex path so the
    # synthetic character/paragraph styles the preprocessors injected surface as
    # custom-style attributes the docx_colors_to_latex / docx_paragraphs_to_latex
    # filters can pick up.
    pandoc_source_format = f"{source_format}+styles" if apply_docx_latex_filters else source_format
    cmd = [PANDOC_PATH, "-f", pandoc_source_format, "-t", target_format, "-o", output_path, source_path]

    # Convert inline CSS on HTML <span style="..."> into raw OOXML runs for
    # the DOCX writer. The filter emits RawInline("openxml", ...) nodes which
    # only render when the target writer is docx; for any other target
    # (markdown, html, pdf, pptx, ...) those nodes are silently dropped and
    # the styled-span text disappears entirely, so the filter must be gated
    # on both source and target.
    if source_format == "html" and target_format == "docx":
        cmd.append(f"--lua-filter={FILTERS['inline_styles']}")
        # Pairs with the HtmlListsPreProcess pass on the source bytes: the
        # preprocessor wraps orphan <ol>/<ul> with a sentinel <li>, and this
        # filter strips the marker paragraph that pandoc would otherwise emit
        # for those synthetic list items.
        cmd.append(f"--lua-filter={FILTERS['html_lists']}")
        # Opt-in: preserve CSS table cell styles (background-color, borders)
        # by rebuilding styled tables as raw OOXML via the Lua filter.
        if preserve_table_styles:
            cmd.extend(["-M", "preserve_table_styles=true"])

    # Companion filters to the DOCX color and paragraph-format preprocessors.
    # Both only emit raw LaTeX, so they are gated on the docx->latex path. Div
    # (paragraph) and Span (run color) scopes are independent, so order between
    # them does not matter.
    if apply_docx_latex_filters:
        cmd.append(f"--lua-filter={FILTERS['docx_colors_to_latex']}")
        cmd.append(f"--lua-filter={FILTERS['docx_paragraphs_to_latex']}")
        cmd.append(f"--lua-filter={FILTERS['docx_lists_to_latex']}")

    if validated_options:
        cmd.extend(validated_options)
    return cmd


def is_svg_conversion_enabled() -> bool:
    """Whether HTML SVG-to-PNG rasterization via headless Chromium is enabled (default on)."""
    return os.environ.get("ENABLE_SVG_CONVERSION", "true").strip().lower() not in {"false", "0", "no"}


def get_chromium_health() -> str:
    """Report Chromium status for the health endpoint (informational, never gates health)."""
    if not is_svg_conversion_enabled():
        return "disabled"
    return "available" if get_chromium_manager().health_check() else "stopped"


async def preprocess_html_svgs(source: str | bytes, scale_factor: float | None = None) -> str | bytes:
    """Rasterize SVGs embedded in HTML to PNG via headless Chromium before pandoc runs.

    Word's built-in SVG engine does not support draw.io's
    ``<switch>`` + ``foreignObject`` mechanism, so a raw SVG embedded in the
    DOCX renders a "Text is not SVG - cannot display" fallback instead of the
    diagram. Rendering the SVG to a PNG here (mirroring the weasyprint-service
    fix) sidesteps that.

    ``scale_factor`` is the device scale factor controlling rasterization
    density (image density). When None, SvgProcessor falls back to the
    DEVICE_SCALE_FACTOR env var (default 1.0). This mirrors weasyprint-service.

    Returns the source unchanged when SVG conversion is disabled, the browser is
    unavailable, the input contains no SVG, or processing fails (best effort:
    a missing rasterizer must never break a conversion).
    """
    if not is_svg_conversion_enabled():
        return source

    is_bytes = isinstance(source, bytes)
    html_str = source.decode("utf-8", errors="replace") if isinstance(source, bytes) else source

    # Cheap guard: skip the parse/serialize round-trip when there is no SVG.
    lowered = html_str.lower()
    if "data:image/svg+xml" not in lowered and "<svg" not in lowered:
        return source

    manager = get_chromium_manager()
    if not manager.is_running():
        logger.warning("SVG conversion requested but Chromium is not running; passing HTML through unchanged")
        return source

    try:
        soup = BeautifulSoup(html_str, "html.parser")
        processor = SvgProcessor(chromium_manager=manager, device_scale_factor=scale_factor)
        processed = await processor.process_svg(soup)
        result = str(processed)
        return result.encode("utf-8") if is_bytes else result
    except Exception as e:  # noqa: BLE001
        logger.exception("SVG preprocessing failed; passing HTML through unchanged: %s", e)
        return source


def run_pandoc_conversion(source_data: str | bytes, source_format: str, target_format: str, options: list[str] | None = None, preserve_table_styles: bool = False) -> bytes:
    """
    Run pandoc conversion using subprocess.

    Args:
        source_data: The data to convert (string or bytes)
        source_format: The source format
        target_format: The target format
        options: Additional pandoc options

    Returns:
        Converted output as bytes
    """
    if options is None:
        options = []

    # Sanitize format parameters to prevent shell injection
    if not source_format.isalnum() or not target_format.isalnum():
        raise ValueError("Format parameters must be alphanumeric")

    # Strict equality check against allowlist
    if source_format not in ALLOWED_SOURCE_FORMATS:
        raise ValueError(f"Invalid source format: {source_format}")

    if target_format not in ALLOWED_TARGET_FORMATS:
        raise ValueError(f"Invalid target format: {target_format}")

    # Validate all options against whitelist to prevent command injection
    validated_options = _validate_pandoc_options(options)

    # Normalize source_data to bytes once, so the rest of the function
    # works with a single type. The temp file below is opened in "wb"
    # mode and DocxColorPreProcess.preprocess expects bytes anyway.
    if isinstance(source_data, str):
        source_data = source_data.encode("utf-8")

    # Pandoc's DOCX reader drops direct run-level color formatting
    # (<w:color>, <w:shd>, <w:highlight>) before producing the AST, so a
    # post-reader Lua filter cannot recover those properties. For
    # targets that ultimately produce LaTeX (PDF, latex), rewrite the
    # colored runs in the source as references to synthetic character
    # styles, then ask pandoc to surface those style references via
    # docx+styles, and let the docx_colors_to_latex Lua filter emit the
    # matching \textcolor / \colorbox raw LaTeX.
    apply_docx_latex_filters = source_format == "docx" and target_format in _LATEX_TARGET_FORMATS
    if apply_docx_latex_filters:
        source_data = DocxColorPreProcess.preprocess(source_data)
        # Same docx->latex gate: pandoc's docx reader drops paragraph alignment
        # (<w:jc>) and collapses left indentation (<w:ind w:left>) into a single
        # BlockQuote before the AST. Rewrite those paragraphs as references to
        # synthetic paragraph styles that survive via docx+styles, and let the
        # docx_paragraphs_to_latex Lua filter emit the matching \leftskip /
        # \centering / \raggedleft. See app/DocxParagraphPreProcess.py.
        source_data = DocxParagraphPreProcess.preprocess(source_data)
        # Same gate: pandoc's docx reader flattens "irregular" lists (a deeper
        # level nested directly inside a shallower one) so the deeper item loses
        # its indentation. Tag each list paragraph's true <w:ilvl> so the
        # docx_lists_to_latex Lua filter can restore the indentation. See
        # app/DocxListLevelPreProcess.py.
        source_data = DocxListLevelPreProcess.preprocess(source_data)

    # html -> docx: rewrite orphan <ol>/<ul> directly nested inside another
    # list so pandoc's HTML reader doesn't synthesize an implicit list item
    # that the DOCX writer would render as a stray marker (e.g. "a.") above
    # the deeper item. See app/HtmlListsPreProcess.py and
    # filters/html_lists.lua for the full pipeline.
    # Also wrap each <p style="margin-left: ...; text-align: ..."> in a marker
    # <div> so the paragraph indent and/or alignment survive pandoc's HTML
    # reader (which drops <p>'s style attribute outright). See
    # app/HtmlParagraphPreProcess.py and the Div handler in
    # filters/inline_styles.lua for the full pipeline.
    if source_format == "html" and target_format == "docx":
        source_data = HtmlListsPreProcess.preprocess(source_data)
        source_data = HtmlParagraphPreProcess.preprocess(source_data)

    with tempfile.NamedTemporaryFile(mode="wb", delete=False) as source_file, tempfile.NamedTemporaryFile(delete=False) as output_file:
        try:
            # Write input data to temporary file
            source_file.write(source_data)
            source_file.flush()

            cmd = _build_pandoc_command(
                source_format=source_format,
                target_format=target_format,
                source_path=source_file.name,
                output_path=output_file.name,
                validated_options=validated_options,
                apply_docx_latex_filters=apply_docx_latex_filters,
                preserve_table_styles=preserve_table_styles,
            )

            # Run pandoc with validated parameters and measure duration
            subprocess_start_time = time.time()
            subprocess.run(cmd, check=True, shell=False, stdin=subprocess.PIPE)
            subprocess_duration = time.time() - subprocess_start_time
            observe_subprocess_duration(subprocess_duration)

            # Read output
            with Path(output_file.name).open("rb") as f:
                return f.read()
        finally:
            # Clean up temporary files
            if Path(source_file.name).exists():
                Path(source_file.name).unlink()
            if Path(output_file.name).exists():
                Path(output_file.name).unlink()


@app.post(
    "/convert/{source_format}/to/docx-with-template",
    summary="Convert to DOCX with a template",
    description="Converts a source document to DOCX format with an optional template file.",
    responses={
        200: {
            "description": "Success",
            "content": {MIME_TYPES["docx"]: {}},
        },
        400: {"description": "Bad request.", "content": {MIME_TYPES["txt"]: {}}},
        413: {"description": "Request body too large.", "content": {MIME_TYPES["txt"]: {}}},
        422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}},
    },
)
async def convert_docx_with_ref(  # noqa: PLR0913, C901
    request: Request,
    source_format: str,
    encoding: str | None = None,
    file_name: str = "converted-document.docx",
    paper_size: str | None = None,
    orientation: str | None = None,
    scale_factor: float | None = None,
    preserve_table_styles: bool = False,
) -> Response:
    temp_template_filename = None
    pandoc_metrics = get_pandoc_metrics()
    conversion_start_time = time.time()
    pandoc_metrics.record_conversion_start()

    try:
        form = await request.form(max_part_size=data_limit)  # NOSONAR False positive - max_part_size is valid parameter
        source_content = form.get("source")
        source = await get_docx_source_data(source_content, encoding)
        if not source:
            pandoc_metrics.record_conversion_failure()
            increment_conversion_failure(source_format, "docx")
            return process_error(Exception("No source file"), "No data or file provided using key 'source'", 400)

        # Record input size
        input_size = len(source) if isinstance(source, bytes) else len(source.encode("utf-8"))
        observe_request_body_size(input_size)

        # Optional docx template file
        docx_template_file = form.get("template")

        if isinstance(docx_template_file, str):
            pandoc_metrics.record_conversion_failure()
            increment_conversion_failure(source_format, "docx")
            return process_error(Exception("Docx template must be a File"), "Invalid template file", 400)

        has_template = bool(docx_template_file)
        if docx_template_file:
            temp_template_filename = f"ref_{int(time.time())}.docx"
            async with await anyio.open_file(temp_template_filename, "wb") as f:
                await f.write(await docx_template_file.read())

        # Build conversion options including template if provided
        options = DEFAULT_CONVERSION_OPTIONS.copy()

        # Add page orientation filter if not already present
        page_orientation_filter = f"--lua-filter={FILTERS['page_orientation']}"
        if page_orientation_filter not in options:
            options.append(page_orientation_filter)

        extended_options = form.get("options")
        if isinstance(extended_options, str):
            options.append(extended_options)

        if temp_template_filename is not None:
            options.append(f"--reference-doc={temp_template_filename}")

        # Rasterize any embedded SVGs to PNG so Word gets a usable image
        # instead of the draw.io "Text is not SVG - cannot display" fallback.
        if source_format == "html":
            source = await preprocess_html_svgs(source, scale_factor)

        # Convert using subprocess instead of pandoc module
        output = run_pandoc_conversion(source, source_format, "docx", options, preserve_table_styles=preserve_table_styles)

        response = postprocess_and_build_response(output, "docx", file_name, paper_size, orientation)

        # Record success metrics
        duration_seconds = time.time() - conversion_start_time
        pandoc_metrics.record_conversion_success(duration_seconds * 1000)
        increment_conversion_success(source_format, "docx", duration_seconds)
        if has_template:
            increment_template_conversion("docx")

        return response

    except Exception as e:
        pandoc_metrics.record_conversion_failure()
        increment_conversion_failure(source_format, "docx")
        return process_error(e, HTTPStatus.BAD_REQUEST.phrase, HTTPStatus.BAD_REQUEST.value)
    finally:
        if temp_template_filename is not None and Path(temp_template_filename).exists():
            Path(temp_template_filename).unlink()


@app.post(
    "/convert/{source_format}/to/pptx-with-template",
    summary="Convert to PPTX with a template",
    description="Converts a source document to PPTX format with an optional template file.",
    responses={
        200: {
            "description": "Success",
            "content": {MIME_TYPES["pptx"]: {}},
        },
        400: {"description": "Bad request.", "content": {MIME_TYPES["txt"]: {}}},
        413: {"description": "Request body too large.", "content": {MIME_TYPES["txt"]: {}}},
        422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}},
    },
)
async def convert_pptx_with_ref(  # noqa: PLR0913
    request: Request,
    source_format: str,
    encoding: str | None = None,
    file_name: str = "converted-document.pptx",
    slide_size: str | None = None,
    scale_factor: float | None = None,
) -> Response:
    temp_template_filename = None
    pandoc_metrics = get_pandoc_metrics()
    conversion_start_time = time.time()
    pandoc_metrics.record_conversion_start()

    try:
        form = await request.form(max_part_size=data_limit)  # NOSONAR False positive - max_part_size is valid parameter
        source_content = form.get("source")
        source = await get_docx_source_data(source_content, encoding)
        if not source:
            pandoc_metrics.record_conversion_failure()
            increment_conversion_failure(source_format, "pptx")
            return process_error(Exception("No source file"), "No data or file provided using key 'source'", 400)

        # Record input size
        input_size = len(source) if isinstance(source, bytes) else len(source.encode("utf-8"))
        observe_request_body_size(input_size)

        # Optional pptx template file
        pptx_template_file = form.get("template")

        if isinstance(pptx_template_file, str):
            pandoc_metrics.record_conversion_failure()
            increment_conversion_failure(source_format, "pptx")
            return process_error(Exception("PPTX template must be a File"), "Invalid template file", 400)

        has_template = bool(pptx_template_file)
        if pptx_template_file:
            temp_template_filename = f"ref_{int(time.time())}.pptx"
            async with await anyio.open_file(temp_template_filename, "wb") as f:
                await f.write(await pptx_template_file.read())

        # Build conversion options including template if provided
        options = DEFAULT_CONVERSION_OPTIONS.copy()

        extended_options = form.get("options")
        if isinstance(extended_options, str):
            options.append(extended_options)

        if temp_template_filename is not None:
            options.append(f"--reference-doc={temp_template_filename}")

        # Rasterize any embedded SVGs to PNG so the slide renderer gets a usable image.
        if source_format == "html":
            source = await preprocess_html_svgs(source, scale_factor)

        # Convert using subprocess instead of pandoc module
        output = run_pandoc_conversion(source, source_format, "pptx", options)

        response = postprocess_and_build_response(output, "pptx", file_name, slide_size, None)

        # Record success metrics
        duration_seconds = time.time() - conversion_start_time
        pandoc_metrics.record_conversion_success(duration_seconds * 1000)
        increment_conversion_success(source_format, "pptx", duration_seconds)
        if has_template:
            increment_template_conversion("pptx")

        return response

    except Exception as e:
        pandoc_metrics.record_conversion_failure()
        increment_conversion_failure(source_format, "pptx")
        return process_error(e, "Bad request", 400)
    finally:
        if temp_template_filename is not None and Path(temp_template_filename).exists():
            Path(temp_template_filename).unlink()


@app.post(
    "/convert/{source_format}/to/{target_format}",
    summary="Convert document between formats",
    description="Converts document from source format to target format",
    responses={
        200: {
            "description": "Success",
            "content": {DEFAULT_MIME_TYPE: {}},
        },
        400: {"description": "Bad request.", "content": {MIME_TYPES["txt"]: {}}},
        413: {"description": "Request body too large.", "content": {MIME_TYPES["txt"]: {}}},
        422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}},
    },
)
async def convert(  # noqa: PLR0913
    request: Request,
    source_format: str,
    target_format: str,
    encoding: str | None = None,
    file_name: str | None = None,
    paper_size: str | None = None,
    orientation: str | None = None,
    scale_factor: float | None = None,
    preserve_table_styles: bool = False,
) -> Response:
    pandoc_metrics = get_pandoc_metrics()
    conversion_start_time = time.time()
    pandoc_metrics.record_conversion_start()

    try:
        file_name = file_name if file_name else "converted-document." + FILE_EXTENSIONS.get(target_format, "docx")
        if source_format in {"txt", "markdown", "html"}:
            data = await request.body()
            source = data if not encoding else data.decode(encoding)
        else:
            form = await request.form(max_part_size=data_limit)  # NOSONAR False positive - max_part_size is valid parameter
            uploaded_file = form.get("source")

            try:
                source = await uploaded_file.read()  # type: ignore
            except AttributeError:
                pandoc_metrics.record_conversion_failure()
                increment_conversion_failure(source_format, target_format)
                return process_error(Exception("Expected file-like object"), "Invalid uploaded file", 400)

        # Record input size
        input_size = len(source) if isinstance(source, bytes) else len(source.encode("utf-8"))
        observe_request_body_size(input_size)

        options = DEFAULT_CONVERSION_OPTIONS.copy()

        if target_format == "pdf":
            options.append("--pdf-engine=tectonic")

        # Rasterize any embedded SVGs to PNG so renderers without full SVG
        # support (e.g. Word) get a usable image instead of a fallback warning.
        if source_format == "html":
            source = await preprocess_html_svgs(source, scale_factor)

        # Convert using subprocess instead of pandoc module
        output = run_pandoc_conversion(source, source_format, target_format, options, preserve_table_styles=preserve_table_styles)

        response = postprocess_and_build_response(output, target_format, file_name, paper_size, orientation)

        # Record success metrics
        duration_seconds = time.time() - conversion_start_time
        pandoc_metrics.record_conversion_success(duration_seconds * 1000)
        increment_conversion_success(source_format, target_format, duration_seconds)

        return response

    except Exception as e:
        pandoc_metrics.record_conversion_failure()
        increment_conversion_failure(source_format, target_format)
        return process_error(e, HTTPStatus.BAD_REQUEST.phrase, HTTPStatus.BAD_REQUEST.value)


async def get_docx_source_data(source_content: starlette.datastructures.UploadFile | str | None, encoding: str | None) -> bytes | str | None:
    if isinstance(source_content, starlette.datastructures.UploadFile):
        source_bytes = await source_content.read()
        if not source_bytes:
            return None
        return source_bytes if not encoding else source_bytes.decode(encoding)
    return source_content


def postprocess_and_build_response(output: bytes, target_format: str, file_name: str, paper_size: str | None = None, orientation: str | None = None) -> Response:
    if target_format == "docx":
        post_process_start = time.time()
        output = DocxPostProcess.process(output, paper_size, orientation)
        observe_post_processing_duration("docx", time.time() - post_process_start)
    elif target_format == "pptx":
        # For PPTX, paper_size parameter is repurposed as slide_size
        post_process_start = time.time()
        output = PptxPostProcess.process(output, paper_size)
        observe_post_processing_duration("pptx", time.time() - post_process_start)

    # Record final response size after post-processing
    observe_response_body_size(len(output))
    mime_type = MIME_TYPES.get(target_format, DEFAULT_MIME_TYPE)

    response = Response(output, media_type=mime_type, status_code=200)
    response.headers.append("Content-Disposition", "attachment; filename=" + file_name)
    response.headers.append("Python-Version", platform.python_version())
    response.headers.append("Pandoc-Version", (get_pandoc_version() or "unknown"))
    response.headers.append("Pandoc-Service-Version", os.environ.get("PANDOC_SERVICE_VERSION", "unknown"))
    return response


def process_error(e: Exception, err_msg: str, status: int) -> PlainTextResponse:
    sanitized_err_msg = err_msg.replace("\r\n", "").replace("\n", "")
    logger.exception(msg=sanitized_err_msg + ": " + str(e))
    return PlainTextResponse(
        content=err_msg + ": " + getattr(e, "message", repr(e)),
        status_code=status,
    )


def start_server(port: int) -> None:
    """Start the server on the specified port.

    Args:
        port: The port number to listen on
    """
    uvicorn.run(app=app, host="", port=port)
