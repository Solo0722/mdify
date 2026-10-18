"""Independent fixtures written by Microsoft Office, not our XML test builders."""

from pathlib import Path

from mdify import convert_docx, convert_pptx, convert_xlsx

FIXTURES = Path(__file__).parent / 'fixtures'


def test_word_fixture():
    md, raw = convert_docx(FIXTURES / 'office.docx')
    assert '# **Fixture heading**' in md
    assert '- First bullet\n- Second bullet' in md
    assert md.endswith('Content control preserved')
    assert raw.splitlines() == ['Fixture heading', 'First bullet', 'Second bullet', 'Content control preserved']


def test_excel_fixture():
    md, raw = convert_xlsx(FIXTURES / 'office.xlsx')
    assert md == '## Values\n\n| Date | Total |\n| --- | --- |\n| 2024-01-01 | 5 |'
    assert raw == 'Values\nDate Total\n2024-01-01 5'


def test_powerpoint_fixture():
    md, raw = convert_pptx(FIXTURES / 'office.pptx')
    assert md.index('Slide 1: Results (hidden)') < md.index('Slide 2: Overview') < md.index('Slide 3')
    assert '| Merged header |  | Third |' in md
    assert '- First bullet\n    - Nested bullet\n- Last bullet' in md
    assert 'Grouped left\n\nGrouped right' in md
    assert '### Speaker notes\n\nSpeaker note: synthetic fixture.' in md
    assert md.endswith('(Empty slide)')
    assert raw.splitlines() == ['Results', 'Merged header', 'Third', 'A|B', 'C', 'D', 'Overview',
                               'First bullet', 'Nested bullet', 'Last bullet', 'Grouped left',
                               'Grouped right', 'Speaker note: synthetic fixture.']


def test_powerpoint_note_opt_out_changes_both_outputs():
    md, raw = convert_pptx(FIXTURES / 'office.pptx', include_notes=False)
    assert 'Speaker note' not in md and 'Speaker note' not in raw
