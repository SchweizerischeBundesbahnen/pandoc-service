"""Run pandoc in the container of the tests: the pandoc and the filters of the image, whatever the machine has."""

from __future__ import annotations

import io
import tarfile
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docker.models.containers import Container

PANDOC = "/usr/local/bin/pandoc"
FILTERS = "/usr/local/share/pandoc/filters"
_FOLDER = "/tmp/pandoc-in-container"


def filter_path(name: str) -> str:
    """The path of a filter in the image, by its path under filters/ of the repository: ``docx/docx_tables_to_latex.lua``."""
    return f"{FILTERS}/{name}"


def run_pandoc(container: Container, arguments: list[str], source: bytes, name: str = "in") -> bytes:
    """Run pandoc with the arguments on the source, as a file of the given name. Returns what pandoc writes to stdout.

    The source goes in as a tar archive: as an argument of a shell it would meet the limit of the length of one.
    """
    archive = io.BytesIO()
    with tarfile.open(fileobj=archive, mode="w") as tar:
        info = tarfile.TarInfo(name)
        info.size = len(source)
        tar.addfile(info, io.BytesIO(source))
    container.exec_run(["mkdir", "-p", _FOLDER])
    assert container.put_archive(_FOLDER, archive.getvalue()), "the source did not reach the container"
    exit_code, (stdout, stderr) = container.exec_run([PANDOC, *arguments, "-o", "-", f"{_FOLDER}/{name}"], demux=True)
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {(stderr or b'').decode()}"
    return stdout or b""
