"""Make the free-font corpus: each DOCX of the corpus with the fonts of Word replaced by open fonts of the image.

    python -m tools.corpus.free_fonts corpus/docx corpus/docx-free

Word and the image then set the documents in the same fonts, so the score of this corpus shows the layout alone.
The theme, the styles and every run name their fonts anew; the equations take Latin Modern Math.
Word prints the references with these fonts installed: see corpus/README.md.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

# The fonts of Word in the corpus, by name in lower case, and the open font of the image for each
FREE_FONTS = {
    "aptos": "Liberation Sans",
    "aptos display": "Liberation Sans",
    "arial": "Liberation Sans",
    "segoe ui": "DejaVu Sans",
    "verdana": "DejaVu Sans",
    "calibri": "Carlito",
    "times new roman": "Liberation Serif",
    "cambria": "Liberation Serif",
    "georgia": "DejaVu Serif",
    "consolas": "DejaVu Sans Mono",
    "courier new": "Liberation Mono",
}
MATH_FONT = "Latin Modern Math"
# The attributes which name the font of Latin text: of a run (w:rFonts), of the theme (a:latin)
_FONT_ATTRIBUTE = re.compile(r'((?:w:ascii|w:hAnsi|w:cs)="|<a:latin typeface=")([^"]*)(")')
_MATH_PROPERTIES = f'<m:mathPr><m:mathFont m:val="{MATH_FONT}"/></m:mathPr>'


def _free(match: re.Match[str]) -> str:
    name = match.group(2)
    return match.group(1) + FREE_FONTS.get(name.strip().casefold(), name) + match.group(3)


def free_part(name: str, xml: str) -> str:
    """The XML of one part of a DOCX, its fonts replaced; the settings name the font of the equations."""
    xml = _FONT_ATTRIBUTE.sub(_free, xml)
    if name == "word/settings.xml" and "<m:mathPr" not in xml:
        # m:mathPr goes before w:themeFontLang in the order of the schema, or before the end
        anchor = next((tag for tag in ("<w:attachedSchema", "<w:themeFontLang", "<w:clrSchemeMapping", "<w:decimalSymbol", "</w:settings>") if tag in xml), None)
        if anchor is not None:
            xml = xml.replace(anchor, _MATH_PROPERTIES + anchor, 1)
    return xml


def free_docx(source: Path, target: Path) -> None:
    """Write TARGET as SOURCE with free fonts, every other part as it is."""
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as copy:
        for item in original.infolist():
            data = original.read(item)
            if item.filename.endswith(".xml"):
                data = free_part(item.filename, data.decode("utf-8")).encode("utf-8")
            copy.writestr(item, data)


def main(arguments: list[str]) -> int:
    if len(arguments) != 2:  # noqa: PLR2004 - the source and the target
        print(__doc__, file=sys.stderr)  # noqa: T201 - a command line tool
        return 2
    source, target = Path(arguments[0]), Path(arguments[1])
    target.mkdir(parents=True, exist_ok=True)
    for docx in sorted(source.glob("*.docx")):
        free_docx(docx, target / docx.name)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
