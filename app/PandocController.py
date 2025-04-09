import io
import logging
import os
import platform
import subprocess
import tempfile
import time
from pathlib import Path

from flask import Flask, Response, request, send_file
from gevent.pywsgi import WSGIServer  # type: ignore

from app import DocxPostProcess

CUSTOM_REFERENCE_DOCX = "custom-reference.docx"
PANDOC_PATH = "/usr/local/bin/pandoc"

# List of allowed pandoc options for security
ALLOWED_PANDOC_OPTIONS = [
    "--lua-filter=/usr/local/share/pandoc/filters/pagebreak.lua",
    "--track-changes=all",
    "--reference-doc=",  # Prefix for reference-doc option
]

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

app = Flask(__name__)

data_limit = 200 * 1024 * 1024  # 200MB;
app.config.update(
    MAX_CONTENT_LENGTH=data_limit,
    MAX_FORM_MEMORY_SIZE=data_limit,
)


def get_pandoc_version() -> str | None:
    """Get the pandoc version using subprocess."""
    try:
        result = subprocess.run(
            [PANDOC_PATH, "--version"],
            capture_output=True,
            text=True,
            check=True,
        )
        # Extract version from the first line of output
        version_line = result.stdout.splitlines()[0]
        return version_line.split()[-1]  # Last word on the first line is the version
    except Exception as e:
        logging.error(f"Error getting pandoc version: {e}")
        return None


@app.route("/version", methods=["GET"])
def version() -> dict[str, str | None]:
    return {
        "python": platform.python_version(),
        "pandoc": get_pandoc_version(),
        "pandocService": os.environ.get("PANDOC_SERVICE_VERSION"),
        "timestamp": os.environ.get("PANDOC_SERVICE_BUILD_TIMESTAMP"),
    }


@app.route("/docx-template", methods=["GET"])
def get_docx_template() -> Response:
    path = Path(CUSTOM_REFERENCE_DOCX)
    try:
        # ruff: noqa: S603
        try:
            subprocess.run(
                [
                    PANDOC_PATH,
                    "-o",
                    "custom-reference.docx",
                    "--print-default-data-file",
                    "reference.docx",
                ],
                check=True,
            )
        except subprocess.SubprocessError as e:
            logging.error(f"Error generating template: {e}")
            return Response(f"Error generating template: {e}", status=500)

        with path.open("rb") as f:
            doc_content = f.read()

        return send_file(
            io.BytesIO(doc_content),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="reference.docx",
        )
    finally:
        if Path.exists(path):
            Path.unlink(path)


ALLOWED_FORMATS = ["markdown", "html", "docx", "pdf", "latex", "textile", "plain"]  # Add other allowed formats as needed


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

    if source_format not in ALLOWED_FORMATS or target_format not in ALLOWED_FORMATS:
        raise ValueError("Invalid format specified.")

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

            # Build pandoc command with validated options
            cmd = [PANDOC_PATH, "-f", source_format, "-t", target_format, "-o", output_file.name, source_file.name] + validated_options

            # Run pandoc
            subprocess.run(cmd, check=True)

            # Read output
            with Path(output_file.name).open("rb") as f:
                return f.read()
        finally:
            # Clean up temporary files
            if Path(source_file.name).exists():
                Path(source_file.name).unlink()
            if Path(output_file.name).exists():
                Path(output_file.name).unlink()


@app.route("/convert/<source_format>/to/<target_format>", methods=["POST"])
def convert(source_format: str, target_format: str) -> Response:
    try:
        encoding = request.args.get("encoding")
        file_name = request.args.get(
            "file_name",
            default=("converted-document." + FILE_EXTENSIONS.get(target_format, "docx")),
        )

        source = request.get_data() if not encoding else request.get_data().decode(encoding)

        # Convert using subprocess instead of pandoc module
        output = run_pandoc_conversion(source, source_format, target_format, DEFAULT_CONVERSION_OPTIONS)

        return postprocess_and_build_response(output, target_format, file_name)

    except Exception as e:
        return process_error(e, "Bad request", 400)


@app.route("/convert/<source_format>/to/docx-with-template", methods=["POST"])
def convert_docx_with_ref(source_format: str) -> Response:
    temp_template_filename = None
    try:
        encoding = request.args.get("encoding")
        file_name = request.args.get(
            "file_name",
            default="converted-document.docx",
        )

        source_text = request.form.get("source")  # first try to get it as a form text data
        if source_text is not None:
            source = source_text
        else:
            source_file = request.files.get("source")  # then we attempt to get it as a file
            if not source_file:
                return process_error(Exception("No source file"), "No data or file provided using key 'source'", 400)
            source = source_file.read() if not encoding else source_file.read().decode(encoding)

        # Optional docx template file
        docx_template_file = request.files.get("template")
        if docx_template_file:
            temp_template_filename = f"ref_{int(time.time())}.docx"
            with Path(temp_template_filename).open("wb") as f:
                f.write(docx_template_file.read())

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


def postprocess_and_build_response(output: bytes, target_format: str, file_name: str) -> Response:
    if target_format == "docx":
        output = DocxPostProcess.replace_table_properties(output)
    mime_type = MIME_TYPES.get(target_format, DEFAULT_MIME_TYPE)

    response = Response(output, mimetype=mime_type, status=200)
    response.headers.add("Content-Disposition", "attachment; filename=" + file_name)
    response.headers.add("Python-Version", platform.python_version())
    response.headers.add("Pandoc-Version", get_pandoc_version() or "unknown")
    response.headers.add("Pandoc-Service-Version", os.environ.get("PANDOC_SERVICE_VERSION"))
    return response


def process_error(e: Exception, err_msg: str, status: int) -> Response:
    sanitized_err_msg = err_msg.replace("\r\n", "").replace("\n", "")
    logging.exception(msg=sanitized_err_msg + ": " + str(e))
    return Response(
        err_msg + ": " + getattr(e, "message", repr(e)),
        mimetype="plain/text",
        status=status,
    )


def start_server(port: int) -> None:
    http_server = WSGIServer(("", port), app)
    http_server.serve_forever()
