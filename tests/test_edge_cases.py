"""Boundary cases for formatting, package errors, and output safety."""

import io
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

from mdify import convert_docx, convert_html, convert_xlsx
from mdify.cli import main, run_convert
from mdify.convert_tabular import convert_csv
from test_regressions import W, S, docx, package, rels, xlsx


@pytest.mark.parametrize('source, expected', [
    ('<p><code>`x`</code></p>', '`` `x` ``'),
    ('<p><code> x </code></p>', '`  x  `'),
    ('<ol start="10"><li>Parent<ul><li>Child</li></ul></li></ol>', '10. Parent\n    - Child'),
    ('<ol><li value="5">Five</li><li>Six</li></ol>', '5. Five\n6. Six'),
    ('<blockquote><p>A</p><p>B</p></blockquote>', '> A\n>\n> B'),
    ('<p>First<br/>Second</p>', 'First  \nSecond'),
    ('\ufeff<p>Hello</p>', 'Hello'),
    ('<p>&amp;amp; ~~literal~~</p>', r'&amp;amp; \~\~literal\~\~'),
    ('<p><a href="https://example.com/a b(c)">[label]</a></p>', r'[\[label\]](https://example.com/a%20b%28c%29)'),
])
def test_html_edge_cases(source, expected):
    assert convert_html(source)[0] == expected


def test_html_rowspan_does_not_shift_values():
    md, _ = convert_html('<table><tr><th>A</th><th>B</th></tr>'
                         '<tr><td rowspan="2">Shared</td><td>First</td></tr><tr><td>Second</td></tr></table>')
    assert md.endswith('| Shared | First |\n|  | Second |')


def test_code_in_list_keeps_blank_lines_and_indentation():
    md, _ = convert_html('<ul><li>Code<pre>x\n\n\n  y\n</pre></li></ul>')
    assert md == '- Code\n\n  ```\n  x\n\n\n    y\n  ```'


def test_csv_explicit_delimiter_and_multiline_cell():
    md, raw = convert_csv('A^B\n"line1\nline2"^a|b\n', delimiter='^')
    assert '| line1<br>line2 | a\\|b |' in md
    assert 'line1\nline2' in raw
    with pytest.raises(ValueError):
        convert_csv('a,b', delimiter='two')


def test_docx_breaks_nested_controls_and_numbering_restart():
    body = '<w:p><w:sdt><w:sdtContent><w:r><w:t>Before</w:t><w:br/><w:t>After</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    assert convert_docx(docx(body)) == ('Before  \nAfter', 'Before\nAfter')
    nums = f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="a">' + ''.join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl>' for i in [0, 1]
    ) + '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="a"/></w:num></w:numbering>'
    paragraphs = ''.join(f'<w:p><w:pPr><w:numPr><w:ilvl w:val="{depth}"/><w:numId w:val="1"/></w:numPr>'
                         f'</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'
                         for depth, text in [(0, 'A'), (1, 'a'), (1, 'b'), (0, 'B'), (1, 'c')])
    md, _ = convert_docx(docx(paragraphs, extras={'word/numbering.xml': nums}))
    assert md == '1. A\n    1. a\n    2. b\n2. B\n    1. c'


def test_docx_horizontal_and_vertical_merged_cells():
    def cell(text, props=''):
        return f'<w:tc><w:tcPr>{props}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>'
    table = '<w:tbl><w:tr>' + cell('Both', '<w:gridSpan w:val="2"/>') + cell('Third') + '</w:tr><w:tr>'
    table += cell('A', '<w:vMerge w:val="restart"/>') + cell('B') + cell('C') + '</w:tr><w:tr>'
    table += cell('', '<w:vMerge/>') + cell('D') + cell('E') + '</w:tr></w:tbl>'
    md, raw = convert_docx(docx(table))
    assert '| Both |  | Third |' in md
    assert md.endswith('| A | B | C |\n|  | D | E |')
    assert raw.count('Both') == 1


@pytest.mark.parametrize('serial, date1904, format_id, expected', [
    ('0', True, 14, '1904-01-01'), ('59', False, 14, '1900-02-28'),
    ('60', False, 14, '1900-02-29'), ('61', False, 14, '1900-03-01'),
    ('0.5', False, 21, '12:00:00'), ('1.5', False, 46, '36:00:00'),
])
def test_excel_dates_and_times(serial, date1904, format_id, expected):
    styles = f'<styleSheet xmlns="{S}"><cellXfs><xf numFmtId="{format_id}"/></cellXfs></styleSheet>'
    md, _ = convert_xlsx(xlsx(f'<row r="1"><c r="A1" s="0"><v>{serial}</v></c></row>',
                             extras={'xl/styles.xml': styles}, date1904=date1904))
    assert expected in md


def test_sparse_far_row_and_column_output_limit():
    md, _ = convert_xlsx(xlsx('<row r="1048576"><c r="A1048576"><v>7</v></c></row>'))
    assert md == '## Data\n\n| 7 |\n| --- |'
    rows = ''.join(f'<row r="{i}"><c r="XFD{i}"><v>1</v></c></row>' for i in range(1, 63))
    with pytest.raises(ValueError, match='output limit'):
        convert_xlsx(xlsx(rows))


def test_missing_sheet_and_invalid_shared_string_fail():
    with pytest.raises(ValueError, match='Missing Office package part'):
        convert_xlsx(xlsx('', target='worksheets/missing.xml'))
    with pytest.raises(ValueError, match='shared string'):
        convert_xlsx(xlsx('<row r="1"><c r="A1" t="s"><v>99</v></c></row>'))


def test_strict_word_namespace_is_supported():
    payload = docx('<w:p><w:r><w:t>Strict text</w:t></w:r></w:p>')
    parts = {}
    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        for name in archive.namelist():
            parts[name] = archive.read(name).replace(W.encode(), b'http://purl.oclc.org/ooxml/wordprocessingml/main')
    assert convert_docx(package(parts)) == ('Strict text', 'Strict text')


@pytest.mark.parametrize('encoding', ['utf-8-sig', 'utf-16', 'utf-32', 'cp1252'])
def test_byte_decoding_file_stdin_equivalence(tmp_path, encoding):
    payload = '<p>Caf\u00e9</p>'.encode(encoding)
    path = tmp_path / 'text.html'
    path.write_bytes(payload)
    for source in [str(path), '-']:
        out = io.StringIO()
        assert run_convert(source, file_type='html', stdin_stream=io.BytesIO(payload), stdout_stream=out) == 0
        assert out.getvalue() == 'Caf\u00e9\n'


def test_output_failure_preserves_existing_destination(tmp_path, monkeypatch):
    source, dest = tmp_path / 'source.html', tmp_path / 'out.md'
    source.write_text('<p>New</p>', encoding='utf-8')
    dest.write_text('Existing', encoding='utf-8')
    def fail(*args):
        raise OSError('test write failure')
    monkeypatch.setattr('mdify.cli.os.replace', fail)
    err = io.StringIO()
    assert run_convert(str(source), str(dest), stderr_stream=err) == 1
    assert dest.read_text(encoding='utf-8') == 'Existing'
    assert not list(tmp_path.glob('.mdify-*'))


def test_hard_link_source_collision(tmp_path):
    source, dest = tmp_path / 'source.html', tmp_path / 'out.md'
    source.write_text('<p>Original</p>', encoding='utf-8')
    try:
        os.link(source, dest)
    except OSError:
        pytest.skip('Hard links unavailable on this filesystem')
    assert run_convert(str(source), str(dest), stderr_stream=io.StringIO()) == 1
    assert source.read_text(encoding='utf-8') == '<p>Original</p>'


def test_cli_options_and_binary_stdin_validation(tmp_path, monkeypatch):
    monkeypatch.setattr('sys.stdin', io.StringIO('a^b\n1^2'))
    out = io.StringIO()
    monkeypatch.setattr('sys.stdout', out)
    assert main(['convert', '-', '--type', 'csv', '--delimiter', '^']) == 0
    assert '| 1 | 2 |' in out.getvalue()
    err = io.StringIO()
    assert run_convert('-', file_type='pptx', stdin_stream=io.StringIO('zip'), stderr_stream=err) == 1
    assert 'binary stdin' in err.getvalue()
    assert run_convert('-', file_type='html', delimiter=',', stderr_stream=io.StringIO()) == 1
    assert run_convert('-', file_type='csv', include_notes=False, stderr_stream=io.StringIO()) == 1


def test_cli_unicode_redirected_output():
    result = subprocess.run([sys.executable, '-m', 'mdify', 'convert', '-', '--type', 'html'],
                            input='<p>\u9053\u8def Caf\u00e9</p>'.encode('utf-8'), capture_output=True,
                            env={**os.environ, 'PYTHONIOENCODING': 'ascii'}, check=True)
    assert result.stdout.decode('utf-8').strip() == '\u9053\u8def Caf\u00e9'
