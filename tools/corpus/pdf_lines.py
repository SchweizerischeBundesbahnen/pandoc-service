"""The lines of text of a PDF, as the corpus compares a conversion with its reference.

A line is a run of characters on one baseline, without a gap wider than 0.8 times
the font size: the cells of a table row and the parts of a line a tab separates
are lines of their own. The same rule applies to both PDFs, so
their lines compare however each producer splits its text into objects.
"""

from __future__ import annotations

import ctypes
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

if TYPE_CHECKING:
    from pathlib import Path

# A line shorter than this is a page number, a bullet or a stray letter: too common to match by its text
MIN_LENGTH = 4
# A glyph without a character, as the letters of an equation come out: a line with one has no text to match by
UNREADABLE = "\ufffd"
LINE_BREAKS = frozenset({0x0A, 0x0D, 0xFFFE, 0})


@dataclass(frozen=True)
class Line:
    """A line of a page: its text, where its baseline starts, in points from the top left, and its width."""

    page: int
    x: float
    y: float
    width: float
    text: str


@dataclass(frozen=True)
class Word:
    """A word of a page: its text and where its baseline starts, in points from the top left."""

    page: int
    x: float
    y: float
    text: str


@dataclass
class _Run:
    page: int
    x: float
    y: float
    right: float
    size: float
    characters: list[str] = field(default_factory=list)
    # Where each character starts, a space where the run has none
    lefts: list[float | None] = field(default_factory=list)

    def line(self) -> Line | None:
        text = normalized("".join(self.characters))
        if len(text) < MIN_LENGTH or UNREADABLE in text:
            return None
        return Line(self.page, round(self.x, 2), round(self.y, 2), round(self.right - self.x, 2), text)

    def words(self) -> list[Word]:
        """The words of the run, split at its spaces and on both sides of a leader of dots."""
        words, characters, start = [], "", None
        for character, left in [*_leaders_apart(list(zip(self.characters, self.lefts, strict=True))), (" ", None)]:
            if character.isspace():
                text = normalized(characters)
                if text and start is not None:
                    words.append(Word(self.page, round(start, 2), round(self.y, 2), text))
                characters, start = "", None
                continue
            characters += character
            start = left if start is None else start
        return words


def _leaders_apart(characters: list[tuple[str, float | None]]) -> list[tuple[str, float | None]]:
    """The characters of a run with a space before and after each leader, four dots or more."""
    text = "".join(character for character, _left in characters)
    apart: list[tuple[str, float | None]] = []
    end = 0
    for leader in re.finditer(r"\.{4,}", text):
        apart += [*characters[end : leader.start()], (" ", None), *characters[leader.start() : leader.end()], (" ", None)]
        end = leader.end()
    return apart + characters[end:]


def normalized(text: str) -> str:
    """The text as the comparison reads it: one space between words, a dot leader as three dots."""
    # pdfium hands a character outside the BMP over as two UTF-16 surrogates
    text = text.encode("utf-16", "surrogatepass").decode("utf-16", "replace")
    # A leader of dots is as long as its line leaves room for: its dots do not count. It stands apart from the
    # text on both sides, whether or not a reader sees a space in the gap Word leaves before the page
    text = re.sub(r"(?:\s?\.){4,}", " ... ", text)
    text = re.sub(r"\s+", " ", text).strip()
    # XeTeX maps the hyphen of Aptos to U+2010, Word to U+002D: the same glyph
    return text.replace("\u2010", "-")


def _continues(run: _Run, baseline: float, left: float, size: float) -> bool:
    return abs(baseline - run.y) < max(1.0, 0.3 * size) and left >= run.right - 1 and left - run.right < 0.8 * max(size, run.size)


def _page_runs(page: pdfium.PdfPage, number: int) -> list[_Run]:
    """The runs of characters of one page, in the order of its characters."""
    textpage = page.get_textpage()
    height = page.get_height()
    origin_x, origin_y = ctypes.c_double(), ctypes.c_double()
    runs: list[_Run] = []
    current: _Run | None = None
    for index in range(textpage.count_chars()):
        code = pdfium_c.FPDFText_GetUnicode(textpage.raw, index)
        if code in LINE_BREAKS:
            continue
        character = chr(code)
        if character.isspace():
            # A space carries no place of its own that both producers agree on
            if current is not None:
                current.characters.append(" ")
                current.lefts.append(None)
            continue
        pdfium_c.FPDFText_GetCharOrigin(textpage.raw, index, origin_x, origin_y)
        left, _bottom, right, _top = textpage.get_charbox(index)
        # Quartz writes its text at a font size of 1 and scales it: the height of the loose box is the size both agree on
        _left, bottom, _right, top = textpage.get_charbox(index, loose=True)
        size = top - bottom
        baseline = height - origin_y.value
        if current is not None and _continues(current, baseline, left, size):
            current.characters.append(character)
            current.lefts.append(left)
            current.right = max(current.right, right)
            continue
        current = _Run(number, left, baseline, right, size, [character], [left])
        runs.append(current)
    return runs


def page_lines(page: pdfium.PdfPage, number: int) -> list[Line]:
    """The lines of one page, in the order of its characters."""
    return [line for line in (run.line() for run in _page_runs(page, number)) if line is not None]


def page_words(page: pdfium.PdfPage, number: int) -> list[Word]:
    """The words of one page, in the order of its characters."""
    return [word for run in _page_runs(page, number) for word in run.words()]


def document_lines(path: Path) -> list[list[Line]]:
    """The lines of each page of a PDF."""
    document = pdfium.PdfDocument(path)
    try:
        return [page_lines(document[number], number) for number in range(len(document))]
    finally:
        document.close()


def document_words(path: Path) -> list[list[Word]]:
    """The words of each page of a PDF."""
    document = pdfium.PdfDocument(path)
    try:
        return [page_words(document[number], number) for number in range(len(document))]
    finally:
        document.close()
