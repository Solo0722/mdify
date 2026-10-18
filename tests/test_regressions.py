"""Content-preservation regressions discovered during the converter audit."""

import io
import zipfile

import pytest

from mdify.cli import run_convert
from mdify.convert_docx import convert_docx
from mdify.convert_html import convert_html
from mdify.convert_tabular import convert_csv, convert_xlsx

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def package(parts):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return stream.getvalue()


def rels(*items):
    return f'<Relationships xmlns="{PKG}">' + "".join(
        f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{target}"/>'
        for rid, kind, target in items
    ) + '</Relationships>'


def docx(body, *, main="word/document.xml", extras=None):
    parts = {
        "_rels/.rels": rels(("doc", "officeDocument", main)),
        main: f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>',
    }
    parts.update(extras or {})
    return package(parts)


def xlsx(rows, *, target="worksheets/sheet1.xml", extras=None, date1904=False):
    parts = {
        "_rels/.rels": rels(("wb", "officeDocument", "xl/workbook.xml")),
        "xl/workbook.xml": f'<workbook xmlns="{S}" xmlns:r="{R}">'
        f'<workbookPr date1904="{int(date1904)}"/><sheets>'
        '<sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": rels(("sheet", "worksheet", target)),
        "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{S}"><sheetData>{rows}</sheetData></worksheet>',
    }
    parts.update(extras or {})
    return package(parts)


def test_html_inline_spaces_and_literal_markdown():
    assert convert_html('<div><span>Hello</span> <span>world</span></div>')[0] == 'Hello world'
    assert convert_html('<p># literal</p><p>*literal*</p>')[0] == '\\# literal\n\n\\*literal\\*'


def test_html_code_preserved_and_fences_safe():
    code = 'a   \n\n\n\nb\n```\n'
    md, _ = convert_html('<pre><code>' + code + '</code></pre>')
    assert md == '````\n' + code + '````'
    assert convert_html('<code>a`b</code>')[0] == '``a`b``'


def test_html_blockquote_and_list_continuations():
    assert convert_html('<blockquote><p>Quote</p></blockquote>')[0] == '> Quote'
    assert convert_html('<ul><li><b>First<li>Second</ul>')[0] == '- **First**\n- Second'
    assert convert_html('<ul><li><p>First</p><p>Second</p></li><li>Next</li></ul>')[0] == '- First\n\n  Second\n- Next'


def test_html_table_spans_and_raw_boundaries():
    md, raw = convert_html('<table><tr><th colspan="2">Both</th><th>Third</th></tr>'
                           '<tr><td>A</td><td>B</td><td>C</td></tr></table>')
    assert '| Both |  | Third |' in md
    assert 'BothThird' not in raw and 'ABC' not in raw


def test_csv_single_column_does_not_guess_letters():
    assert convert_csv('hello\nworld\n')[0] == '| hello |\n| --- |\n| world |'


def test_docx_relationships_and_content_controls():
    body = '<w:sdt><w:sdtContent><w:p><w:r><w:t>Preserved</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    assert convert_docx(docx(body, main='word/document2.xml')) == ('Preserved', 'Preserved')


def test_docx_style_numbering_and_start_override():
    body = '<w:p><w:pPr><w:pStyle w:val="CustomList"/></w:pPr><w:r><w:t>Fifth</w:t></w:r></w:p>'
    extras = {
        'word/_rels/document.xml.rels': rels(('styles', 'styles', 'styles.xml'), ('nums', 'numbering', 'numbering.xml')),
        'word/styles.xml': f'<w:styles xmlns:w="{W}"><w:style w:styleId="CustomList"><w:pPr>'
        '<w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>',
        'word/numbering.xml': f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0">'
        '<w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0">'
        '<w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>',
    }
    assert convert_docx(docx(body, extras=extras))[0] == '5. Fifth'


@pytest.mark.parametrize('converter', [convert_docx, convert_xlsx])
def test_invalid_office_package_is_an_error(converter):
    with pytest.raises(ValueError):
        converter(package({'unrelated.txt': 'Not a document'}))


def test_xlsx_uncached_formula_and_relative_relationship():
    rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Total</t></is></c></row>'
    rows += '<row r="2"><c r="A2"><f>SUM(B2:C2)</f></c></row>'
    md, raw = convert_xlsx(xlsx(rows, target='worksheets/../worksheets/sheet1.xml'))
    assert '=SUM(B2:C2)' in md and '=SUM(B2:C2)' in raw


def test_xlsx_date_format():
    extras = {'xl/styles.xml': f'<styleSheet xmlns="{S}"><cellXfs><xf numFmtId="0"/>'
              '<xf numFmtId="14"/></cellXfs></styleSheet>'}
    md, _ = convert_xlsx(xlsx('<row r="1"><c r="A1" s="1"><v>45292</v></c></row>', extras=extras))
    assert '2024-01-01' in md


def test_cli_same_input_output_preserves_source(tmp_path):
    path = tmp_path / 'source.html'
    path.write_text('<p>Original</p>', encoding='utf-8')
    err = io.StringIO()
    assert run_convert(str(path), str(path), stderr_stream=err) == 1
    assert path.read_text(encoding='utf-8') == '<p>Original</p>'
    assert 'same' in err.getvalue().lower()


@pytest.mark.parametrize('fmt, source', [('html', '<p>caf\u00e9</p>'), ('csv', 'name\ncaf\u00e9\n')])
def test_cli_file_stdin_encoding_equivalence(tmp_path, fmt, source):
    payload = source.encode('cp1252')
    path = tmp_path / ('source.' + fmt)
    path.write_bytes(payload)
    file_out, stdin_out = io.StringIO(), io.StringIO()
    assert run_convert(str(path), stdout_stream=file_out) == 0
    assert run_convert('-', file_type=fmt, stdin_stream=io.BytesIO(payload), stdout_stream=stdin_out) == 0
    assert file_out.getvalue() == stdin_out.getvalue()
    assert 'caf\u00e9' in stdin_out.getvalue()
