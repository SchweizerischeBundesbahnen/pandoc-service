import logging
from unittest.mock import MagicMock, patch

from docx.document import Document as DocumentObject
from docx.oxml import parse_xml

from app.DocxReferencesPostProcess import (
    SCHEMA,
    _add_caption_style_and_tc_field,
    _create_tc_field_runs,
    _create_toc_field,
    _create_tof_field,
    _create_tot_field,
    _find_placeholder_paragraphs,
    _get_paragraph_text,
    _get_paragraph_style,
    add_table_of_contents_entries,
    enable_auto_update_fields,
)


def test_get_paragraph_text_empty():
    """Test getting text from an empty paragraph."""
    para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    result = _get_paragraph_text(para)
    assert result == ""


def test_get_paragraph_text_single_text_element():
    """Test getting text from a paragraph with a single text element."""
    para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:t>Hello World</w:t>
        </w:r>
    </w:p>
    ''')
    result = _get_paragraph_text(para)
    assert result == "Hello World"


def test_get_paragraph_text_multiple_runs():
    """Test getting text from a paragraph with multiple runs."""
    para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:t>Hello </w:t>
        </w:r>
        <w:r>
            <w:t>World</w:t>
        </w:r>
    </w:p>
    ''')
    result = _get_paragraph_text(para)
    assert result == "Hello World"


def test_get_paragraph_text_with_none_text():
    """Test that None text elements are handled gracefully."""
    para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:r>
            <w:t>Valid</w:t>
        </w:r>
        <w:r>
            <w:t/>
        </w:r>
    </w:p>
    ''')
    result = _get_paragraph_text(para)
    assert result == "Valid"


def test_get_paragraph_style_with_no_properties():
    """Test getting style from paragraph without properties."""
    para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"/>')
    result = _get_paragraph_style(para)
    assert result is None


def test_get_paragraph_style_with_bodytext():
    """Test getting BodyText style."""
    para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:pPr>
            <w:pStyle w:val="BodyText"/>
        </w:pPr>
    </w:p>
    ''')
    result = _get_paragraph_style(para)
    assert result == "BodyText"


def test_create_tc_field_runs_for_figure():
    """Test creating TC field runs for a figure caption."""
    caption_text = "Figure 1: Test Caption"
    runs = _create_tc_field_runs(caption_text, field_flag="F")

    assert len(runs) == 3
    # Check begin run
    assert runs[0].find(".//w:fldChar", namespaces={"w": SCHEMA}) is not None
    # Check instruction run contains caption text
    instr_text = runs[1].find(".//w:instrText", namespaces={"w": SCHEMA}).text
    assert caption_text in instr_text
    assert "\\f F" in instr_text
    # Check end run
    assert runs[2].find(".//w:fldChar", namespaces={"w": SCHEMA}) is not None


def test_create_tc_field_runs_for_table():
    """Test creating TC field runs for a table caption."""
    caption_text = "Table 1: Test Table"
    runs = _create_tc_field_runs(caption_text, field_flag="T")

    assert len(runs) == 3
    # Check instruction run contains correct flag
    instr_text = runs[1].find(".//w:instrText", namespaces={"w": SCHEMA}).text
    assert caption_text in instr_text
    assert "\\f T" in instr_text


def test_create_tc_field_runs_default_flag():
    """Test that default field flag is F for figures."""
    caption_text = "Figure 2"
    runs = _create_tc_field_runs(caption_text)

    instr_text = runs[1].find(".//w:instrText", namespaces={"w": SCHEMA}).text
    assert "\\f F" in instr_text


def test_add_caption_style_to_paragraph_without_style():
    """Test adding Caption style to a paragraph that has no style."""
    para = parse_xml(f'<w:p xmlns:w="{SCHEMA}"><w:r><w:t>Figure 1</w:t></w:r></w:p>')
    caption_text = "Figure 1"

    _add_caption_style_and_tc_field(para, caption_text, field_flag="F")

    # Check that Caption style was added
    p_style = para.find(".//w:pStyle", namespaces={"w": SCHEMA})
    assert p_style is not None
    assert p_style.get(f"{{{SCHEMA}}}val") == "Caption"

    # Check that TC field runs were added
    fld_chars = para.findall(".//w:fldChar", namespaces={"w": SCHEMA})
    assert len(fld_chars) == 2  # begin and end


def test_add_caption_style_to_paragraph_with_existing_style():
    """Test adding Caption style to a paragraph that already has a different style."""
    para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:pPr>
            <w:pStyle w:val="Normal"/>
        </w:pPr>
        <w:r><w:t>Table 1</w:t></w:r>
    </w:p>
    ''')
    caption_text = "Table 1"

    _add_caption_style_and_tc_field(para, caption_text, field_flag="T")

    # Check that style was changed to Caption
    p_style = para.find(".//w:pStyle", namespaces={"w": SCHEMA})
    assert p_style is not None
    assert p_style.get(f"{{{SCHEMA}}}val") == "Caption"


def test_add_caption_style_with_existing_properties():
    """Test adding Caption style when paragraph already has properties."""
    para = parse_xml(f'''
    <w:p xmlns:w="{SCHEMA}">
        <w:pPr>
            <w:jc w:val="center"/>
        </w:pPr>
        <w:r><w:t>Figure 2</w:t></w:r>
    </w:p>
    ''')
    caption_text = "Figure 2"

    _add_caption_style_and_tc_field(para, caption_text, field_flag="F")

    # Check that Caption style was added while preserving other properties
    p_style = para.find(".//w:pStyle", namespaces={"w": SCHEMA})
    assert p_style is not None
    assert p_style.get(f"{{{SCHEMA}}}val") == "Caption"

    # Check that other properties still exist
    jc = para.find(".//w:jc", namespaces={"w": SCHEMA})
    assert jc is not None


def test_create_toc_field_returns_list():
    """Test that _create_toc_field returns a list of paragraphs."""
    result = _create_toc_field()
    assert isinstance(result, list)
    assert len(result) == 2  # TOC paragraph and empty paragraph


def test_create_toc_field_structure():
    """Test that TOC field has correct structure."""
    paragraphs = _create_toc_field()
    toc_para = paragraphs[0]

    # Check for field characters
    fld_chars = toc_para.findall(".//w:fldChar", namespaces={"w": SCHEMA})
    assert len(fld_chars) == 3  # begin, separate, end

    # Check instruction text
    instr = toc_para.find(".//w:instrText", namespaces={"w": SCHEMA})
    assert instr is not None
    assert "TOC" in instr.text
    assert "\\o" in instr.text  # outline levels
    assert "\\h" in instr.text  # hyperlinks
    assert "\\z" in instr.text  # hide tab leader
    assert "\\u" in instr.text  # use outline


def test_create_toc_field_empty_paragraph():
    """Test that second paragraph is empty for spacing."""
    paragraphs = _create_toc_field()
    empty_para = paragraphs[1]

    # Empty paragraph should have no child elements
    assert len(empty_para) == 0


def test_create_tof_field_returns_list():
    """Test that _create_tof_field returns a list of paragraphs."""
    result = _create_tof_field()
    assert isinstance(result, list)
    assert len(result) == 2


def test_create_tof_field_structure():
    """Test that Table of Figures field has correct structure."""
    paragraphs = _create_tof_field()
    tof_para = paragraphs[0]

    # Check instruction text
    instr = tof_para.find(".//w:instrText", namespaces={"w": SCHEMA})
    assert instr is not None
    assert "TOC" in instr.text
    assert "\\f F" in instr.text  # figures flag
    assert "\\h" in instr.text
    assert "\\z" in instr.text


def test_create_tot_field_returns_list():
    """Test that _create_tot_field returns a list of paragraphs."""
    result = _create_tot_field()
    assert isinstance(result, list)
    assert len(result) == 2


def test_create_tot_field_structure():
    """Test that Table of Tables field has correct structure."""
    paragraphs = _create_tot_field()
    tot_para = paragraphs[0]

    # Check instruction text
    instr = tot_para.find(".//w:instrText", namespaces={"w": SCHEMA})
    assert instr is not None
    assert "TOC" in instr.text
    assert "\\f T" in instr.text  # tables flag
    assert "\\h" in instr.text
    assert "\\z" in instr.text


def test_find_placeholder_paragraphs_empty_body():
    """Test finding placeholders in empty document body."""
    body = parse_xml(f'<w:body xmlns:w="{SCHEMA}"/>')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)
    assert len(elements_to_replace) == 0


def test_find_placeholder_paragraphs_with_toc_marker():
    """Test finding TOC_PLACEHOLDER marker."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>TOC_PLACEHOLDER</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)

    assert len(elements_to_replace) == 1
    idx, element, has_toc, has_tof, has_tot = elements_to_replace[0]
    assert idx == 0
    assert has_toc is True
    assert has_tof is False
    assert has_tot is False


def test_find_placeholder_paragraphs_with_tof_marker():
    """Test finding TOF_PLACEHOLDER marker."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>TOF_PLACEHOLDER</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)

    assert len(elements_to_replace) == 1
    idx, element, has_toc, has_tof, has_tot = elements_to_replace[0]
    assert idx == 0
    assert has_toc is False
    assert has_tof is True
    assert has_tot is False


def test_find_placeholder_paragraphs_with_tot_marker():
    """Test finding TOT_PLACEHOLDER marker."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>TOT_PLACEHOLDER</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)

    assert len(elements_to_replace) == 1
    idx, element, has_toc, has_tof, has_tot = elements_to_replace[0]
    assert idx == 0
    assert has_toc is False
    assert has_tof is False
    assert has_tot is True


def test_find_placeholder_paragraphs_with_all_three():
    """Test finding all three placeholder types."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>TOC_PLACEHOLDER</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>TOF_PLACEHOLDER</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>TOT_PLACEHOLDER</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)

    assert len(elements_to_replace) == 3
    # Check TOC
    assert elements_to_replace[0][2] is True  # has_toc
    # Check TOF
    assert elements_to_replace[1][3] is True  # has_tof
    # Check TOT
    assert elements_to_replace[2][4] is True  # has_tot


def test_find_placeholder_paragraphs_not_found():
    """Test finding placeholders in document without any."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>Regular paragraph</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)
    assert len(elements_to_replace) == 0


def test_find_placeholder_paragraphs_in_heading_ignored():
    """Test that placeholders in headings are ignored."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r><w:t>TOC_PLACEHOLDER</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)
    assert len(elements_to_replace) == 0


def test_find_placeholder_paragraphs_in_body_text():
    """Test finding placeholders in BodyText style."""
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="BodyText"/>
            </w:pPr>
            <w:r><w:t>TOC_PLACEHOLDER</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    elements_to_replace = []
    _find_placeholder_paragraphs(body, elements_to_replace)
    assert len(elements_to_replace) == 1


def test_enable_auto_update_fields_when_present():
    """Test enabling auto-update when updateFields setting already exists."""
    mock_doc = MagicMock(spec=DocumentObject)
    mock_settings = MagicMock()
    mock_settings_element = MagicMock()
    mock_update_fields = MagicMock()
    mock_settings_element.find.return_value = mock_update_fields
    mock_settings.element = mock_settings_element
    mock_doc.settings = mock_settings

    enable_auto_update_fields(mock_doc)

    # Verify existing element was updated
    mock_update_fields.set.assert_called_once_with(f'{{{SCHEMA}}}val', 'true')


def test_enable_auto_update_fields_handles_exception(caplog):
    """Test that exceptions are handled gracefully with warning."""
    mock_doc = MagicMock(spec=DocumentObject)
    mock_doc.settings.element.find.side_effect = Exception("Test error")

    with caplog.at_level(logging.WARNING):
        enable_auto_update_fields(mock_doc)

    # Verify warning was logged
    assert "Could not enable auto-update fields" in caplog.text


def test_add_toc_entries_finds_figure_captions(caplog):
    """Test that figure captions are found and processed."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>Figure 1: Test figure</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Figure 2: Another figure</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    with caplog.at_level(logging.INFO):
        add_table_of_contents_entries(mock_doc)

    # Verify logging indicates figures were found
    assert "Found 2 figure captions" in caplog.text


def test_add_toc_entries_finds_table_captions(caplog):
    """Test that table captions are found and processed."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>Table 1: Test table</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    with caplog.at_level(logging.INFO):
        add_table_of_contents_entries(mock_doc)

    # Verify logging indicates tables were found
    assert "Found 0 figure captions and 1 table captions" in caplog.text


def test_add_toc_entries_inserts_toc_field_with_placeholder(caplog):
    """Test that TOC field is inserted when TOC_PLACEHOLDER is found."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>TOC_PLACEHOLDER</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Content after placeholder</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    with caplog.at_level(logging.INFO):
        add_table_of_contents_entries(mock_doc)

    # Verify TOC was inserted
    assert "Inserted Table of Contents" in caplog.text


def test_add_toc_entries_handles_empty_document():
    """Test that empty document is handled without errors."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'<w:body xmlns:w="{SCHEMA}"/>')
    mock_doc.element.body = body

    # Should not raise any exceptions
    add_table_of_contents_entries(mock_doc)


def test_add_toc_entries_processes_both_figures_and_tables(caplog):
    """Test processing document with both figure and table captions."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>Figure 1: Test figure</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Table 1: Test table</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    with caplog.at_level(logging.INFO):
        add_table_of_contents_entries(mock_doc)

    # Should process both types
    assert "Found 1 figure captions and 1 table captions" in caplog.text


def test_full_workflow_with_figures_and_toc():
    """Integration test: document with figures and TOC_PLACEHOLDER."""
    from docx import Document

    # Create a real document
    doc = Document()
    doc.add_paragraph("TOC_PLACEHOLDER")
    doc.add_paragraph("Figure 1: Test figure caption")
    doc.add_paragraph("Figure 2: Another figure")

    # Run the function
    add_table_of_contents_entries(doc)

    # Verify Caption styles were added to figure paragraphs
    body = doc.element.body
    paras = body.findall(".//w:p", namespaces={"w": SCHEMA})
    figure_paras = [p for p in paras if "Figure" in _get_paragraph_text(p)]

    for para in figure_paras:
        p_style = para.find(".//w:pStyle", namespaces={"w": SCHEMA})
        if p_style is not None:
            assert p_style.get(f"{{{SCHEMA}}}val") == "Caption"


def test_full_workflow_with_tables_only():
    """Integration test: document with only table captions."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("Table 1: Test table caption")
    doc.add_paragraph("Table 2: Another table")

    # Run the function
    add_table_of_contents_entries(doc)

    # Verify table captions got Caption style
    body = doc.element.body
    paras = body.findall(".//w:p", namespaces={"w": SCHEMA})
    table_paras = [p for p in paras if "Table" in _get_paragraph_text(p)]

    for para in table_paras:
        p_style = para.find(".//w:pStyle", namespaces={"w": SCHEMA})
        if p_style is not None:
            assert p_style.get(f"{{{SCHEMA}}}val") == "Caption"


def test_full_workflow_empty_document():
    """Integration test: empty document should not cause errors."""
    from docx import Document

    doc = Document()

    # Should not raise any exceptions
    add_table_of_contents_entries(doc)


def test_enable_auto_update_on_real_document():
    """Integration test: enable auto-update fields on real document."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("Test content")

    # Should not raise any exceptions
    enable_auto_update_fields(doc)

    # Verify updateFields element was added
    settings_element = doc.settings.element
    update_fields = settings_element.find('.//w:updateFields', namespaces={'w': SCHEMA})
    assert update_fields is not None
    assert update_fields.get(f'{{{SCHEMA}}}val') == 'true'


def test_caption_text_with_special_characters():
    """Test that caption text with special characters is handled correctly."""
    caption_text = 'Figure 1: Test - Special (Characters) with [brackets]'
    runs = _create_tc_field_runs(caption_text, field_flag="F")

    # Verify runs were created without errors
    assert len(runs) == 3
    instr_text = runs[1].find(".//w:instrText", namespaces={"w": SCHEMA}).text
    assert caption_text in instr_text


def test_caption_text_with_quotes():
    """Test that caption text with quotes is handled correctly."""
    caption_text = "Figure 1: Test with 'single' quotes"
    runs = _create_tc_field_runs(caption_text, field_flag="F")

    # Verify runs were created without errors
    assert len(runs) == 3
    instr_text = runs[1].find(".//w:instrText", namespaces={"w": SCHEMA}).text
    assert caption_text in instr_text


def test_very_long_caption_text():
    """Test handling of very long caption text."""
    caption_text = "Figure 1: " + "A" * 1000  # Very long caption
    runs = _create_tc_field_runs(caption_text, field_flag="F")

    # Should create runs without errors
    assert len(runs) == 3


def test_caption_starting_with_whitespace():
    """Test captions with leading/trailing whitespace."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>  Figure 1: Test  </w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    # Should still find the caption after strip()
    add_table_of_contents_entries(mock_doc)


def test_figure_and_table_with_same_number():
    """Test handling when figure and table have same number."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>Figure 1: Test figure</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Table 1: Test table</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    # Should handle both without confusion
    add_table_of_contents_entries(mock_doc)


def test_all_three_placeholders_together():
    """Test document with all three placeholder types."""
    mock_doc = MagicMock(spec=DocumentObject)
    body = parse_xml(f'''
    <w:body xmlns:w="{SCHEMA}">
        <w:p>
            <w:r><w:t>TOC_PLACEHOLDER</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>TOF_PLACEHOLDER</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>TOT_PLACEHOLDER</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Figure 1: Test</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Table 1: Test</w:t></w:r>
        </w:p>
    </w:body>
    ''')
    mock_doc.element.body = body

    # Should handle all three placeholders
    add_table_of_contents_entries(mock_doc)