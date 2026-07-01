"""Integration tests for ``filters/docx_lists_to_latex.lua``.

Runs the real ``pandoc`` binary inside the pandoc-service container with the
filter on a native AST that mimics what the docx reader produces for a Polarion
"irregular" list once ``DocxListLevelPreProcess`` has tagged each item with its
true ``<w:ilvl>`` (a leading ``\\uE000<level>\\uE001`` sentinel ``Str``).
Pandoc-only — no DOCX fixture or tectonic needed.
"""

from __future__ import annotations

from tests.test_container import TestParameters

PANDOC_PATH = "/usr/local/bin/pandoc"
FILTER_PATH = "/usr/local/share/pandoc/filters/docx_lists_to_latex.lua"

OPEN = "\uE000"
CLOSE = "\uE001"


def _tag(level: int) -> str:
    return f"{OPEN}{level}{CLOSE}"


def _native_to_latex(container, native: str) -> str:
    container.exec_run(["sh", "-c", "mkdir -p /tmp/test"])
    container.exec_run(["sh", "-c", f"cat > /tmp/test/in.native << 'HEREDOC_EOF'\n{native}\nHEREDOC_EOF"])
    exit_code, output = container.exec_run(
        ["sh", "-c", f"{PANDOC_PATH} -f native -t latex --lua-filter={FILTER_PATH} /tmp/test/in.native"],
    )
    assert exit_code == 0, f"pandoc failed (exit {exit_code}): {output.decode()}"
    return output.decode("utf-8")


# The flattened AST pandoc produces for the irregular ordered list, with the
# preprocessor's level sentinels prepended: Level 1 (lvl0), then a LowerRoman
# sublist holding "Level 3" (lvl2), then a LowerAlpha sublist holding
# "Level 2" (lvl1) — all as depth-2 siblings under Level 1's item.
def _ordered_irregular() -> str:
    return (
        "[ OrderedList (1, Decimal, Period)\n"
        f'  [ [ Plain [Str "{_tag(0)}Level", Space, Str "1"]\n'
        f'    , OrderedList (1, LowerRoman, Period) [ [ Plain [Str "{_tag(2)}Level", Space, Str "3"] ] ]\n'
        f'    , OrderedList (1, LowerAlpha, Period) [ [ Plain [Str "{_tag(1)}Level", Space, Str "2"] ] ]\n'
        "    ] ] ]\n"
    )


def _bulleted_irregular() -> str:
    return f'[ BulletList\n  [ [ Plain [Str "{_tag(0)}B1"]\n    , BulletList [ [ Plain [Str "{_tag(2)}B3"] ] ]\n    , BulletList [ [ Plain [Str "{_tag(1)}B2"] ] ]\n    ] ] ]\n'


def test_ordered_irregular_pushes_only_the_deeper_sublist(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, _ordered_irregular())
    # Exactly one marker-less wrapper level — for "Level 3" (lvl2 at depth2).
    assert latex.count("\\begin{enumerate}\\item[]") == 1, latex
    # "Level 2" (lvl1 at depth2) is NOT wrapped.
    assert "\\begin{itemize}" not in latex
    # Sentinels are gone and content is intact.
    assert OPEN not in latex and CLOSE not in latex, "sentinel leaked into output"
    assert "Level 3" in latex and "Level 2" in latex and "Level 1" in latex


def test_bulleted_irregular_uses_itemize_wrapper(test_parameters: TestParameters):
    latex = _native_to_latex(test_parameters.container, _bulleted_irregular())
    assert latex.count("\\begin{itemize}\\item[]") == 1, latex
    assert OPEN not in latex and CLOSE not in latex
    assert "B3" in latex and "B2" in latex


def test_wellformed_contiguous_list_is_not_wrapped(test_parameters: TestParameters):
    """A properly nested 0->1->2 list needs no push; only the sentinels are stripped."""
    native = (
        "[ OrderedList (1, Decimal, Period)\n"
        f'  [ [ Plain [Str "{_tag(0)}L1"]\n'
        "    , OrderedList (1, LowerAlpha, Period)\n"
        f'      [ [ Plain [Str "{_tag(1)}L2"]\n'
        f'        , OrderedList (1, LowerRoman, Period) [ [ Plain [Str "{_tag(2)}L3"] ] ]\n'
        "        ] ] ] ] ]\n"
    )
    latex = _native_to_latex(test_parameters.container, native)
    assert "\\item[]" not in latex, latex
    assert OPEN not in latex and CLOSE not in latex
    assert "L1" in latex and "L2" in latex and "L3" in latex


def test_untagged_list_is_left_alone(test_parameters: TestParameters):
    """A plain list with no sentinels must pass through unchanged (no wrapper)."""
    native = '[ OrderedList (1, Decimal, Period) [ [ Plain [Str "plain"] ] ] ]\n'
    latex = _native_to_latex(test_parameters.container, native)
    assert "\\item[]" not in latex
    assert "plain" in latex
