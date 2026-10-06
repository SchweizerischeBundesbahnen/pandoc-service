"""Parse a whole HTML document, also one larger than libxml2's default limits.

lxml's default HTML parser stops at libxml2's resource limits: about 10 MB for
one value, such as the ``data:`` URI of an embedded image, and about 20 MB for
the whole input. The HTML parser recovers instead of raising, so the tree just
ends there, and a pre-processor that writes it back drops the rest of the
document. ``huge_tree`` lifts those limits. The request size limit of the
service still bounds the input.

The cost of a parse grows with the input, not with ``huge_tree``: about 3 ms
and 1 MB of memory per MB parsed, and about 4 times the input at the peak when
the tree is written back.
"""

from __future__ import annotations

from lxml import html  # type: ignore[import-untyped]


def parse_document(source: bytes) -> html.HtmlElement:
    """Parse ``source`` as an HTML document, to its end.

    A new parser per call keeps calls independent of each other; creating one is cheap.
    """
    return html.document_fromstring(source, parser=html.HTMLParser(huge_tree=True))
