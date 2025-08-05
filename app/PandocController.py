import io
import logging
import os
import platform
import subprocess
import tempfile
import time
from collections.abc import Awaitable, Callable
from pathlib import Path

import anyio
import starlette.datastructures
import uvicorn
from fastapi import FastAPI, Request, Response
from fastapi.exceptions import RequestValidationError
from fastapi.responses import PlainTextResponse, StreamingResponse

from app.schema import VersionSchema

from . import DocxPostProcess

CUSTOM_REFERENCE_DOCX = "custom-reference.docx"
PANDOC_PATH = "/usr/local/bin/pandoc"

# List of allowed pandoc options for security
ALLOWED_PANDOC_OPTIONS = [
    "--lua-filter=/usr/local/share/pandoc/filters/pagebreak.lua",
    "--track-changes=all",
    "--reference-doc=",  # Prefix for reference-doc option
    "--pdf-engine=tectonic",
]

# Add other allowed formats as needed
ALLOWED_SOURCE_FORMATS = ["docx", "epub", "fb2", "html", "json", "latex", "markdown", "rtf", "textile"]
ALLOWED_TARGET_FORMATS = ["docx", "epub", "fb2", "html", "json", "latex", "markdown", "odt", "pdf", "plain", "rtf", "textile"]

MIME_TYPES = {
    "html": "text/html",
    "html5": "text/html",
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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

DEFAULT_CONVERSION_OPTIONS = ["--lua-filter=/usr/local/share/pandoc/filters/pagebreak.lua", "--track-changes=all"]

app = FastAPI(
    openapi_url="/static/openapi.json",
    docs_url="/api/docs",
)
data_limit = 200 * 1024 * 1024  # 200MB;


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
        logging.error(f"Error getting pandoc version: {e}")
        return None


@app.get(
    "/version",
    summary="Get service version information",
    description="Get version information for python, pandoc executable and pandoc service",
    response_model=VersionSchema,
    responses={200: {"description": "Success", "content": {MIME_TYPES["txt"]: {}}}, 422: {"description": "Validation error.", "content": {MIME_TYPES["txt"]: {}}}},
)
def version() -> VersionSchema:
    return VersionSchema(
        python=platform.python_version(),
        pandoc=get_pandoc_version(),
        pandocService=os.environ.get("PANDOC_SERVICE_VERSION"),
        timestamp=os.environ.get("PANDOC_SERVICE_BUILD_TIMESTAMP"),
    )


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

        with path.open("rb") as f:
            doc_content = f.read()

        response = StreamingResponse(
            io.BytesIO(doc_content),
            headers={"Content-Disposition": "attachment; filename=reference.docx"},
            media_type=MIME_TYPES["docx"],
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


def run_pandoc_conversion(source_data: str | bytes, source_format: str, target_format: str, options: list[str] | None = None) -> bytes:
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

    with tempfile.NamedTemporaryFile(mode="wb", delete=False) as source_file, tempfile.NamedTemporaryFile(delete=False) as output_file:
        try:
            # Write input data to temporary file
            if isinstance(source_data, str):
                source_file.write(source_data.encode("utf-8"))
            else:
                source_file.write(source_data)
            source_file.flush()

            # Build pandoc command with validated options and safe parameters
            cmd = [PANDOC_PATH, "-f", source_format, "-t", target_format, "-o", output_file.name, source_file.name]

            # Add validated options separately
            if validated_options:
                cmd.extend(validated_options)

            # Run pandoc with validated parameters
            subprocess.run(cmd, check=True, shell=False, stdin=subprocess.PIPE)

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
async def convert_docx_with_ref(request: Request, source_format: str, encoding: str | None = None, file_name: str = "converted-document.docx"):  # type: ignore
    temp_template_filename = None
    try:
        form = await request.form()
        source_content = form.get("source")
        source = await get_docx_source_data(source_content, encoding)
        if not source:
            return process_error(Exception("No source file"), "No data or file provided using key 'source'", 400)

        # Optional docx template file
        docx_template_file = form.get("template")

        if isinstance(docx_template_file, str):
            return process_error(Exception("Docx template must be a File"), "Invalid template file", 400)

        if docx_template_file:
            temp_template_filename = f"ref_{int(time.time())}.docx"
            async with await anyio.open_file(temp_template_filename, "wb") as f:
                await f.write(await docx_template_file.read())

        # Build conversion options including template if provided
        options = DEFAULT_CONVERSION_OPTIONS.copy()
        if temp_template_filename is not None:
            options.append(f"--reference-doc={temp_template_filename}")

        # Convert using subprocess instead of pandoc module
        output = run_pandoc_conversion(source, source_format, "docx", options)

        return postprocess_and_build_response(output, "docx", file_name)

    except Exception as e:
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
async def convert(request: Request, source_format: str, target_format: str, encoding: str | None = None, file_name: str | None = None) -> Response:
    try:
        file_name = file_name if file_name else "converted-document." + FILE_EXTENSIONS.get(target_format, "docx")
        if source_format in {"txt", "markdown", "html"}:
            data = await request.body()
            source = data if not encoding else data.decode(encoding)
        else:
            form = await request.form()
            uploaded_file = form.get("source")

            try:
                source = await uploaded_file.read()  # type: ignore
            except AttributeError:
                return process_error(Exception("Expected file-like object"), "Invalid uploaded file", 400)

        options = DEFAULT_CONVERSION_OPTIONS.copy()
        if target_format == "pdf":
            options.append("--pdf-engine=tectonic")

        # Convert using subprocess instead of pandoc module
        output = run_pandoc_conversion(source, source_format, target_format, options)

        return postprocess_and_build_response(output, target_format, file_name)

    except Exception as e:
        return process_error(e, "Bad request", 400)


async def get_docx_source_data(source_content: starlette.datastructures.UploadFile | str | None, encoding: str | None) -> bytes | str | None:
    if isinstance(source_content, starlette.datastructures.UploadFile):
        source_bytes = await source_content.read()
        if not source_bytes:
            return None
        return source_bytes if not encoding else source_bytes.decode(encoding)
    return source_content


def postprocess_and_build_response(output: bytes, target_format: str, file_name: str) -> Response:
    if target_format == "docx":
        output = DocxPostProcess.replace_table_properties(output)
    mime_type = MIME_TYPES.get(target_format, DEFAULT_MIME_TYPE)

    response = Response(output, media_type=mime_type, status_code=200)
    response.headers.append("Content-Disposition", "attachment; filename=" + file_name)
    response.headers.append("Python-Version", platform.python_version())
    response.headers.append("Pandoc-Version", (get_pandoc_version() or "unknown"))
    response.headers.append("Pandoc-Service-Version", os.environ.get("PANDOC_SERVICE_VERSION", "unknown"))
    return response


def process_error(e: Exception, err_msg: str, status: int) -> PlainTextResponse:
    sanitized_err_msg = err_msg.replace("\r\n", "").replace("\n", "")
    logging.exception(msg=sanitized_err_msg + ": " + str(e))
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
