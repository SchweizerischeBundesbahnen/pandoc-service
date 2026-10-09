"""Resolve the left and right indents a paragraph is laid out with.

A list item states no indent of its own: it names a numbering level, and the level states the indent.
A paragraph style can state one too, or name a numbering of its own. Word reads each side from the
nearest source which states it: the paragraph's own properties, then its numbering level, then its
style and the styles that one is based on, then the document defaults.
"""

from typing import Any

from app.docx_numbers import whole_number

SCHEMA = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # NOSONAR
_W = f"{{{SCHEMA}}}"

# Each side with the name a newer document may give it instead.
_SIDES = {"left": ("left", "start"), "right": ("right", "end")}


class ParagraphIndents:
    """The indents of paragraphs in one document, in twips, read from its styles and numbering."""

    def __init__(self, styles: Any, numbering: Any) -> None:
        self._styles = {style.get(f"{_W}styleId"): style for style in styles.findall(f"{_W}style")} if styles is not None else {}
        self._default_style = next(
            (style_id for style_id, style in self._styles.items() if style.get(f"{_W}type") == "paragraph" and style.get(f"{_W}default") in ("1", "true")),
            None,
        )
        self._defaults = styles.find(f"{_W}docDefaults/{_W}pPrDefault/{_W}pPr") if styles is not None else None
        self._numbering = numbering

    def width_taken(self, p: Any) -> int:
        """The part of the text width the paragraph's indents take, in twips: left, right and a first line indent."""
        sources = self._sources(p)
        left = _side(sources, "left")
        right = _side(sources, "right")
        # A first line set in from the left takes room from a picture which starts it; a hanging one gives none
        first_line = _first_line(sources)
        return max(left, 0) + max(right, 0) + max(first_line, 0)

    def _sources(self, p: Any) -> list[Any]:
        """The pPr elements that can state an indent for the paragraph, nearest first."""
        own = p.find(f"{_W}pPr")
        style_chain = self._style_chain(own)
        num_id, level = _numbering_reference([own, *style_chain])
        sources = [own, self._level_properties(num_id, level), *style_chain, self._defaults]
        return [source for source in sources if source is not None]

    def _style_chain(self, own: Any) -> list[Any]:
        reference = own.find(f"{_W}pStyle") if own is not None else None
        style_id = reference.get(f"{_W}val") if reference is not None else self._default_style
        chain: list[Any] = []
        seen: set[str] = set()
        while style_id is not None and style_id not in seen:
            seen.add(style_id)
            style = self._styles.get(style_id)
            if style is None:
                break
            properties = style.find(f"{_W}pPr")
            if properties is not None:
                chain.append(properties)
            based_on = style.find(f"{_W}basedOn")
            style_id = based_on.get(f"{_W}val") if based_on is not None else None
        return chain

    def _level_properties(self, num_id: str | None, level: str) -> Any:
        """The pPr of a numbering level: an override of the numbering instance first, else its abstract numbering's."""
        if self._numbering is None or num_id is None or num_id == "0":
            return None
        num = next((element for element in self._numbering.findall(f"{_W}num") if element.get(f"{_W}numId") == num_id), None)
        if num is None:
            return None
        for override in num.findall(f"{_W}lvlOverride"):
            if override.get(f"{_W}ilvl") == level:
                properties = override.find(f"{_W}lvl/{_W}pPr")
                if properties is not None:
                    return properties
        abstract_reference = num.find(f"{_W}abstractNumId")
        abstract_id = abstract_reference.get(f"{_W}val") if abstract_reference is not None else None
        abstract = next((element for element in self._numbering.findall(f"{_W}abstractNum") if element.get(f"{_W}abstractNumId") == abstract_id), None)
        if abstract is None:
            return None
        lvl = next((element for element in abstract.findall(f"{_W}lvl") if element.get(f"{_W}ilvl") == level), None)
        return lvl.find(f"{_W}pPr") if lvl is not None else None


def _numbering_reference(properties: list[Any]) -> tuple[str | None, str]:
    """The numbering and level the paragraph names, itself or through its style, nearest first."""
    return _nearest_numbering_value(properties, "numId"), _nearest_numbering_value(properties, "ilvl") or "0"


def _nearest_numbering_value(properties: list[Any], name: str) -> str | None:
    """The value of numPr/<name> in the nearest properties which state it."""
    for element in properties:
        reference = element.find(f"{_W}numPr/{_W}{name}") if element is not None else None
        value = reference.get(f"{_W}val") if reference is not None else None
        if value is not None:
            return value
    return None


def _side(sources: list[Any], side: str) -> int:
    for source in sources:
        ind = source.find(f"{_W}ind")
        if ind is None:
            continue
        for name in _SIDES[side]:
            value = _twips(ind.get(f"{_W}{name}"))
            if value is not None:
                return value
    return 0


def _first_line(sources: list[Any]) -> int:
    """The first line indent. The nearest source stating it or a hanging indent decides, as the two exclude each other."""
    for source in sources:
        ind = source.find(f"{_W}ind")
        if ind is None:
            continue
        if ind.get(f"{_W}hanging") is not None:
            return 0
        value = _twips(ind.get(f"{_W}firstLine"))
        if value is not None:
            return value
    return 0


def _twips(value: str | None) -> int | None:
    return whole_number(value, signed=True)
