"""PPTX tests use relationship order and exact content, not ZIP filename order."""

import io
import zipfile

import pytest

from mdify import convert_pptx, convert_pptx_file
from mdify.cli import run_convert

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def rels(*items):
    return f'<Relationships xmlns="{PKG}">' + ''.join(
        f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{target}"'
        + (' TargetMode="External"' if target.startswith('https:') else '') + '/>'
        for rid, kind, target in items) + '</Relationships>'


def paragraph(text, properties='', run_properties=''):
    return f'<a:p>{properties}<a:r>{run_properties}<a:t>{text}</a:t></a:r></a:p>'


def shape(text='', *, ph='', paragraphs=None):
    return '<p:sp><p:nvSpPr><p:cNvPr id="1" name="Text"/><p:cNvSpPr/><p:nvPr>' + ph + \
        '</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' + \
        (paragraphs if paragraphs is not None else paragraph(text)) + '</p:txBody></p:sp>'


def slide(shapes, attrs=''):
    return f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" {attrs}><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld></p:sld>'


def presentation(slides, *, order=None, extras=None):
    order = order if order is not None else list(range(1, len(slides) + 1))
    parts = {
        '_rels/.rels': rels(('main', 'officeDocument', 'ppt/presentation.xml')),
        'ppt/presentation.xml': f'<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:sldIdLst>' +
        ''.join(f'<p:sldId id="{255+i}" r:id="s{i}"/>' for i in order) + '</p:sldIdLst></p:presentation>',
        'ppt/_rels/presentation.xml.rels': rels(*[(f's{i}', 'slide', f'slides/slide{i}.xml') for i in order]),
    }
    parts.update({f'ppt/slides/slide{i}.xml': xml for i, xml in enumerate(slides, 1)})
    parts.update(extras or {})
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w') as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
    return stream.getvalue()


def test_slide_order_title_first_groups_and_raw_text():
    title = shape('Overview', ph='<p:ph type="title"/>')
    group = '<p:grpSp>' + shape('Group A') + '<p:grpSp>' + shape('Group B') + '</p:grpSp></p:grpSp>'
    payload = presentation([slide(shape('Body') + title + group), slide(shape('Second file'))], order=[2, 1])
    md, raw = convert_pptx(payload)
    assert md.index('Second file') < md.index('## Slide 2: Overview') < md.index('Body') < md.index('Group A') < md.index('Group B')
    assert raw.splitlines() == ['Second file', 'Overview', 'Body', 'Group A', 'Group B']
    for value in ['Second file', 'Overview', 'Body', 'Group A', 'Group B']:
        assert md.count(value) == 1


def test_inherited_bullets_overrides_numbering_and_formatting():
    ph = '<p:ph idx="1"/>'
    paragraphs = paragraph('Inherited') + paragraph('Nested', '<a:pPr lvl="1"/>') + \
        paragraph('Plain', '<a:pPr><a:buNone/></a:pPr>') + \
        paragraph('Five', '<a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr>') + \
        paragraph('Six', '<a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr>')
    layout = f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>' + \
        shape('', ph='<p:ph type="body" idx="1"/>') + '</p:spTree></p:cSld></p:sldLayout>'
    master = f'<p:sldMaster xmlns:p="{P}" xmlns:a="{A}"><p:txStyles><p:bodyStyle>' + \
        '<a:lvl1pPr><a:buChar char="*"/><a:defRPr b="1"/></a:lvl1pPr>' + \
        '<a:lvl2pPr><a:buChar char="*"/></a:lvl2pPr></p:bodyStyle></p:txStyles></p:sldMaster>'
    extras = {
        'ppt/slides/_rels/slide1.xml.rels': rels(('layout', 'slideLayout', '../slideLayouts/layout.xml')),
        'ppt/slideLayouts/layout.xml': layout,
        'ppt/slideLayouts/_rels/layout.xml.rels': rels(('master', 'slideMaster', '../slideMasters/master.xml')),
        'ppt/slideMasters/master.xml': master,
    }
    md, _ = convert_pptx(presentation([slide(shape(ph=ph, paragraphs=paragraphs))], extras=extras))
    assert '- **Inherited**\n    - Nested' in md
    assert '\n\n**Plain**\n\n5. **Five**\n6. **Six**' in md


def test_notes_only_body_hidden_and_blank_slides():
    notes = f'<p:notes xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>' + \
        shape('Speaker note', ph='<p:ph type="body"/>') + shape('99', ph='<p:ph type="sldNum"/>') + \
        shape('Footer', ph='<p:ph type="ftr"/>') + '</p:spTree></p:cSld></p:notes>'
    extras = {'ppt/slides/_rels/slide1.xml.rels': rels(('notes', 'notesSlide', '../notesSlides/notes1.xml')),
              'ppt/notesSlides/notes1.xml': notes}
    payload = presentation([slide('', 'show="0"')], extras=extras)
    md, raw = convert_pptx(payload)
    assert '## Slide 1 (hidden)' in md and '(Empty slide)' in md
    assert '### Speaker notes\n\nSpeaker note' in md
    assert raw == 'Speaker note' and '99' not in md and 'Footer' not in md
    md, raw = convert_pptx(payload, include_notes=False)
    assert 'Speaker note' not in md and raw == ''


def test_tables_images_charts_and_links():
    def cell(text, attrs=''):
        return f'<a:tc {attrs}><a:txBody>{paragraph(text)}</a:txBody></a:tc>'
    table = '<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tr>' + \
        cell('Merged', 'gridSpan="2"') + cell('', 'hMerge="1"') + cell('Third') + \
        '</a:tr><a:tr>' + cell('A|B') + cell('C') + cell('D') + '</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    image = '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture" descr="A diagram"/></p:nvPicPr></p:pic>'
    chart = '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>'
    linked = shape(paragraphs=paragraph('External', run_properties='<a:rPr b="1"><a:hlinkClick r:id="url"/></a:rPr>') +
                   paragraph('Next', run_properties='<a:rPr><a:hlinkClick r:id="next"/></a:rPr>'))
    extras = {'ppt/slides/_rels/slide1.xml.rels': rels(('url', 'hyperlink', 'https://example.com/a(b)'), ('next', 'slide', 'slide2.xml'))}
    md, raw = convert_pptx(presentation([slide(table + image + chart + linked), slide('')], extras=extras))
    assert '| Merged |  | Third |' in md and r'A\|B' in md
    assert '[Image: A diagram]' in md and '[Chart omitted]' in md
    assert '[**External**](https://example.com/a%28b%29)' in md
    assert '[Next](#slide-2)' in md and '<a id="slide-2"></a>' in md
    assert 'A diagram' in raw and 'Chart omitted' not in raw


def test_pptx_file_stdin_stats_and_no_notes(tmp_path):
    payload = presentation([slide(shape('Caf\u00e9'))])
    path = tmp_path / 'input.PPTX'
    path.write_bytes(payload)
    assert convert_pptx_file(path) == convert_pptx(payload)
    outputs = []
    for source in [str(path), '-']:
        out, err = io.StringIO(), io.StringIO()
        assert run_convert(source, file_type='pptx', stdin_stream=io.BytesIO(payload),
                           stdout_stream=out, stderr_stream=err, show_stats=True, include_notes=False) == 0
        outputs.append(out.getvalue())
        assert 'Statistics:' in err.getvalue() and 'Statistics:' not in out.getvalue()
    assert outputs[0] == outputs[1]


@pytest.mark.parametrize('payload', [b'not a zip', presentation([slide('')], extras={'ppt/slides/slide1.xml': '<broken>'}),
                                   presentation([slide('')], extras={'ppt/_rels/presentation.xml.rels': rels()})])
def test_bad_pptx_fails_without_output(payload):
    out, err = io.StringIO(), io.StringIO()
    assert run_convert('-', file_type='pptx', stdin_stream=io.BytesIO(payload), stdout_stream=out, stderr_stream=err) == 1
    assert out.getvalue() == '' and 'Error converting' in err.getvalue()


def test_empty_presentation_and_caller_owned_zip():
    payload = presentation([])
    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        assert convert_pptx(archive) == ('(Empty presentation)', '')
        assert archive.fp is not None
