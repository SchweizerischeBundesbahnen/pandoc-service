"""Unit tests for :mod:`app.docx_paragraph_indent`."""

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from app.docx_paragraph_indent import ParagraphIndents

W = nsdecls("w")


def _styles(*styles: str, defaults: str = "") -> object:
    return parse_xml(f"<w:styles {W}><w:docDefaults><w:pPrDefault><w:pPr>{defaults}</w:pPr></w:pPrDefault></w:docDefaults>{''.join(styles)}</w:styles>")


def _numbering(level_ind: str = '<w:ind w:left="1440" w:hanging="360"/>', override: str = "") -> object:
    return parse_xml(
        f"<w:numbering {W}>"
        f'<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
        f'<w:lvl w:ilvl="1"><w:pPr>{level_ind}</w:pPr></w:lvl></w:abstractNum>'
        f'<w:num w:numId="3"><w:abstractNumId w:val="7"/>{override}</w:num>'
        f"</w:numbering>"
    )


def _paragraph(properties: str) -> object:
    return parse_xml(f"<w:p {W}><w:pPr>{properties}</w:pPr></w:p>")


def _list_item(level: int, extra: str = "") -> object:
    return _paragraph(f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="3"/></w:numPr>{extra}')


def test_a_paragraph_without_indents_takes_nothing():
    assert ParagraphIndents(_styles(), None).width_taken(_paragraph("")) == 0


def test_the_indents_of_the_paragraph_itself_count_on_both_sides():
    assert ParagraphIndents(_styles(), None).width_taken(_paragraph('<w:ind w:left="600" w:right="400"/>')) == 1000


def test_a_newer_document_names_the_sides_start_and_end():
    assert ParagraphIndents(_styles(), None).width_taken(_paragraph('<w:ind w:start="600" w:end="400"/>')) == 1000


def test_a_list_item_is_indented_by_its_numbering_level():
    indents = ParagraphIndents(_styles(), _numbering())

    assert indents.width_taken(_list_item(0)) == 720
    assert indents.width_taken(_list_item(1)) == 1440


def test_an_override_of_the_numbering_instance_wins_over_its_abstract_numbering():
    override = '<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:pPr><w:ind w:left="2000"/></w:pPr></w:lvl></w:lvlOverride>'

    assert ParagraphIndents(_styles(), _numbering(override=override)).width_taken(_list_item(1)) == 2000


def test_the_paragraph_own_indent_wins_over_its_numbering_level_side_by_side():
    assert ParagraphIndents(_styles(), _numbering()).width_taken(_list_item(1, '<w:ind w:right="300"/>')) == 1440 + 300


def test_numbering_zero_is_no_numbering():
    paragraph = _paragraph('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="0"/></w:numPr>')

    assert ParagraphIndents(_styles(), _numbering()).width_taken(paragraph) == 0


def test_a_style_states_the_indent_through_the_style_it_is_based_on():
    styles = _styles(
        '<w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:ind w:left="500"/></w:pPr></w:style>',
        '<w:style w:type="paragraph" w:styleId="Quote"><w:basedOn w:val="Base"/><w:pPr><w:ind w:right="500"/></w:pPr></w:style>',
    )

    assert ParagraphIndents(styles, None).width_taken(_paragraph('<w:pStyle w:val="Quote"/>')) == 1000


def test_a_style_can_name_the_numbering_and_the_paragraph_its_level():
    styles = _styles('<w:style w:type="paragraph" w:styleId="ListBullet"><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>')

    paragraph = _paragraph('<w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="1"/></w:numPr>')

    assert ParagraphIndents(styles, _numbering()).width_taken(paragraph) == 1440


def test_a_paragraph_without_a_style_takes_the_default_one():
    styles = _styles('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:ind w:left="250"/></w:pPr></w:style>')

    assert ParagraphIndents(styles, None).width_taken(_paragraph("")) == 250


def test_the_document_defaults_come_last():
    assert ParagraphIndents(_styles(defaults='<w:ind w:left="100"/>'), None).width_taken(_paragraph("")) == 100


def test_a_first_line_indent_takes_room_and_a_hanging_one_does_not():
    indents = ParagraphIndents(_styles(), None)

    assert indents.width_taken(_paragraph('<w:ind w:left="400" w:firstLine="300"/>')) == 700
    assert indents.width_taken(_paragraph('<w:ind w:left="400" w:hanging="300"/>')) == 400


def test_a_negative_indent_gives_no_room():
    assert ParagraphIndents(_styles(), None).width_taken(_paragraph('<w:ind w:left="-400" w:right="200"/>')) == 200


def test_a_numbering_which_names_a_missing_abstract_numbering_states_nothing():
    numbering = parse_xml(f'<w:numbering {W}><w:num w:numId="3"><w:abstractNumId w:val="99"/></w:num></w:numbering>')

    assert ParagraphIndents(_styles(), numbering).width_taken(_list_item(0)) == 0


def test_an_indent_which_is_not_ascii_digits_states_none():
    assert ParagraphIndents(_styles(), None).width_taken(_paragraph('<w:ind w:left="²" w:right="300"/>')) == 300
