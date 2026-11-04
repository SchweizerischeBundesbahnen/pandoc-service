"""Integration tests for ``filters/html_nowrap.lua``.

Converts HTML to DOCX through the pandoc-service container, which applies the
filter to every HTML to DOCX conversion, and reads the text of each paragraph
of the produced ``document.xml``. Inside a white-space: nowrap element, a space
and the gap between a picture and its text must be a non-breaking space, the
only character Word does not break a line at.
"""

from __future__ import annotations

from tests.test_container import TestParameters
from tests.test_html_whitespace_filter import _png, _texts


def test_status_icon_keeps_its_label(test_parameters: TestParameters):
    """Polarion's status of a work item: an icon followed by its name, which Word must not put on the next line."""
    body = f'<p>Status: <span class="polarion-JSEnumOption" style="white-space:nowrap"><img src="{_png()}" alt="">Draft</span></p>'
    assert _texts(test_parameters, body) == ["Status: [img]\u00a0Draft"]


def test_picture_and_text_without_nowrap_stay_apart(test_parameters: TestParameters):
    assert _texts(test_parameters, f'<p><span class="polarion-JSEnumOption"><img src="{_png()}" alt="">Draft</span></p>') == ["[img]Draft"]


def test_spaces_inside_nowrap_do_not_break(test_parameters: TestParameters):
    assert _texts(test_parameters, '<p><span style="white-space: nowrap">two  words</span> free</p>') == ["two\u00a0words free"]


def test_text_before_picture_keeps_it(test_parameters: TestParameters):
    assert _texts(test_parameters, f'<p><span style="white-space:nowrap">EL-101<img src="{_png()}"></span></p>') == ["EL-101\u00a0[img]"]


def test_space_next_to_picture_becomes_one_non_breaking_space(test_parameters: TestParameters):
    assert _texts(test_parameters, f'<p><span style="white-space:nowrap"><img src="{_png()}"> Draft</span></p>') == ["[img]\u00a0Draft"]


def test_picture_and_text_in_separate_styled_spans_keep_together(test_parameters: TestParameters):
    """The gap crosses element boundaries, and inline_styles turns the colored span into raw OOXML afterwards."""
    body = f'<p><span style="white-space:nowrap"><span><img src="{_png()}"></span><span style="color:#000000;">EL-101</span></span></p>'
    assert _texts(test_parameters, body) == ["[img]\u00a0EL-101"]


def test_picture_outside_nowrap_stays_apart_from_text_inside(test_parameters: TestParameters):
    """A gap gets a non-breaking space only when both of its sides are inside a nowrap element."""
    body = f'<p><img src="{_png()}"><span style="white-space:nowrap">Draft</span></p>'
    assert _texts(test_parameters, body) == ["[img]Draft"]


def test_span_setting_whitespace_back_to_normal_breaks_again(test_parameters: TestParameters):
    body = '<p><span style="white-space:nowrap">two words <span style="white-space:normal">free text</span></span></p>'
    assert _texts(test_parameters, body) == ["two\u00a0words\u00a0free text"]


def test_nowrap_div_applies_to_its_paragraphs(test_parameters: TestParameters):
    body = '<div style="white-space:nowrap"><p>one two</p><div style="white-space:normal"><p>three four</p></div></div>'
    assert _texts(test_parameters, body) == ["one\u00a0two", "three four"]
