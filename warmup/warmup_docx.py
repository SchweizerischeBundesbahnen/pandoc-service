"""Build-time check that the tectonic cache holds what a conversion of DOCX to PDF needs (see the Dockerfile).

The warm-up compiles warmup.md with the latex template of pandoc. A DOCX goes another way: through the docx
filters, the fonts of the image and the size of the body of the document, each of which asks tectonic for
files of its own. This script converts warmup.md to DOCX in each body size of Word, and converts each DOCX
to PDF as the service does, with tectonic limited to its cache. A file the cache misses fails the build,
which names it: add the construct which needs it to warmup.md. The script writes the file "done" into its
folder when every conversion succeeded.

Usage: python warmup_docx.py FOLDER
"""

from __future__ import annotations

import io
import re
import subprocess
import sys
import zipfile
from pathlib import Path

from app.pandoc_controller import DEFAULT_CONVERSION_OPTIONS, PANDOC_PATH, run_pandoc_conversion

# The size of the body in half-points: LaTeX loads a file for each, size10.clo, size11.clo and size12.clo.
BODY_SIZES = (20, 22, 24)
_DEFAULT_SIZE = re.compile(r'(<w:rPrDefault>.*?<w:sz w:val=")\d+(")', re.DOTALL)


def with_body_size(docx: bytes, half_points: int) -> bytes:
    """The DOCX with its default run size, the size of the body, set to *half_points*."""
    source = zipfile.ZipFile(io.BytesIO(docx))
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "word/styles.xml":
                data = _DEFAULT_SIZE.sub(rf"\g<1>{half_points}\g<2>", data.decode("utf-8"), count=1).encode("utf-8")
            target.writestr(item, data)
    return output.getvalue()


def main(folder: Path) -> None:
    # The arguments are the paths of the build, no input of a request
    docx = subprocess.run([PANDOC_PATH, str(folder / "warmup.md"), "-t", "docx", "-o", "-", f"--resource-path={folder}"], capture_output=True, check=True).stdout  # noqa: S603
    for half_points in BODY_SIZES:
        run_pandoc_conversion(with_body_size(docx, half_points), "docx", "pdf", [*DEFAULT_CONVERSION_OPTIONS, "--pdf-engine=tectonic"])
    (folder / "done").touch()


if __name__ == "__main__":
    main(Path(sys.argv[1]))
