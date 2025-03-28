import io
import logging
import os
import platform
import subprocess
import time
from pathlib import Path

import pandoc  # type: ignore
from flask import Flask, Response, request, send_file
from gevent.pywsgi import WSGIServer  # type: ignore

from app import DocxPostProcess

CUSTOM_REFERENCE_DOCX = "custom-reference.docx"

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


@app.route("/version", methods=["GET"])
def version() -> dict[str, str | None]:
    pandoc_config = pandoc.configure(auto=True, read=True)
    return {
        "python": platform.python_version(),
        "pandoc": pandoc_config.get("version"),
        "pandocService": os.environ.get("PANDOC_SERVICE_VERSION"),
        "timestamp": os.environ.get("PANDOC_SERVICE_BUILD_TIMESTAMP"),
    }


@app.route("/docx-template", methods=["GET"])
def get_docx_template() -> Response:
    path = Path(CUSTOM_REFERENCE_DOCX)
    try:
        # ruff: noqa: S603
        subprocess.run(
            [
                "/usr/local/bin/pandoc",
                "-o",
                "custom-reference.docx",
                "--print-default-data-file",
                "reference.docx",
            ],
            check=True,
        )

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


@app.route("/convert/<source_format>/to/<target_format>", methods=["POST"])
def convert(source_format: str, target_format: str) -> Response:
    try:
        encoding = request.args.get("encoding")
        file_name = request.args.get(
            "file_name",
            default=("converted-document." + FILE_EXTENSIONS.get(target_format, "docx")),
        )

        source = request.get_data() if not encoding else request.get_data().decode(encoding)
        doc = pandoc.read(source, format=source_format)
        output = pandoc.write(doc, format=target_format, options=DEFAULT_CONVERSION_OPTIONS)

        return postprocess_and_build_response(output, target_format, file_name)

    except AssertionError as e:
        return process_error(e, "Assertion error, check the request body", 400)
    except (UnicodeDecodeError, LookupError) as e:
        return process_error(e, "Cannot decode request body", 400)
    except Exception as e:
        return process_error(e, f"Unexpected error due converting to {target_format}", 500)


@app.route("/convert/<source_format>/to/docx-with-template", methods=["POST"])
def convert_docx_with_ref(source_format: str) -> Response:
    temp_template_filename = None
    try:
        encoding = request.args.get("encoding")
        file_name = request.args.get(
            "file_name",
            default="converted-document.docx",
        )

        source = request.form.get("source")  # first try to get it as a form text data
        if not source:
            source_file = request.files.get("source")  # then we attempt to get it as a file
            if not source_file:
                return process_error(Exception("No source file"), "No data or file provided using key 'source'", 400)
            source = source_file.read() if not encoding else source_file.read().decode(encoding)

        doc = pandoc.read(source, format=source_format)

        # Optional docx template file
        docx_template_file = request.files.get("template")
        if docx_template_file:
            temp_template_filename = f"ref_{int(time.time())}.docx"
            with Path(temp_template_filename).open("wb") as f:
                f.write(docx_template_file.read())

        output = pandoc.write(
            doc,
            format="docx",
            options=DEFAULT_CONVERSION_OPTIONS + ([f"--reference-doc={temp_template_filename}"] if temp_template_filename is not None else []),
        )

        return postprocess_and_build_response(output, "docx", file_name)

    except AssertionError as e:
        return process_error(e, "Assertion error, check the request data", 400)
    except (UnicodeDecodeError, LookupError) as e:
        return process_error(e, "Cannot decode source content", 400)
    except Exception as e:
        return process_error(e, "Unexpected error due converting to docx", 500)
    finally:
        if temp_template_filename is not None:
            Path.unlink(Path(temp_template_filename))


def postprocess_and_build_response(output: bytes, target_format: str, file_name: str) -> Response:
    if target_format == "docx":
        output = DocxPostProcess.replace_table_properties(output)
    mime_type = MIME_TYPES.get(target_format, DEFAULT_MIME_TYPE)

    pandoc_config = pandoc.configure(auto=True, read=True)
    response = Response(output, mimetype=mime_type, status=200)
    response.headers.add("Content-Disposition", "attachment; filename=" + file_name)
    response.headers.add("Python-Version", platform.python_version())
    response.headers.add("Pandoc-Version", pandoc_config.get("version"))
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
