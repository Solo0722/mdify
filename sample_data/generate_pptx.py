"""Generate a small, editable PowerPoint sample using only the standard library."""

from pathlib import Path
from xml.sax.saxutils import escape
import zipfile

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _rels(items):
    return f'<Relationships xmlns="{PKG}">' + ''.join(
        f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{target}"/>'
        for rid, kind, target in items) + '</Relationships>'


def _paragraph(text, bullet=False):
    props = '<a:pPr><a:buChar char="&#8226;"/></a:pPr>' if bullet else '<a:pPr><a:buNone/></a:pPr>'
    return f'<a:p>{props}<a:r><a:rPr lang="en-US"/><a:t>{escape(text)}</a:t></a:r></a:p>'


def _shape(sid, text, *, title=False, bullet=False, y=1500000):
    ph = '<p:ph type="title"/>' if title else '<p:ph type="body" idx="1"/>'
    paragraphs = ''.join(_paragraph(line, bullet) for line in text.split('\n'))
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Text {sid}"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr>' \
        f'</p:nvSpPr><p:spPr><a:xfrm><a:off x="600000" y="{y}"/><a:ext cx="7900000" cy="1400000"/></a:xfrm>' \
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>' + paragraphs + '</p:txBody></p:sp>'


def _tree(shapes=''):
    return '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' \
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/>' \
        '<a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' + shapes + '</p:spTree>'


def _xml(kind, inner, attrs=''):
    return f'<p:{kind} xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" {attrs}>{inner}</p:{kind}>'


def generate_sample_pptx(output_path: Path) -> None:
    color_map = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' \
                'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    overrides = '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
    slide1 = _xml('sld', '<p:cSld>' + _tree(
        _shape(2, 'Sample PPTX Presentation', title=True, y=200000) +
        _shape(3, 'Slide text and lists\nSpeaker notes are included by default', bullet=True)) + '</p:cSld>' + overrides)
    cells = lambda values: '<a:tr h="450000">' + ''.join(
        '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>' + _paragraph(value) + '</a:txBody><a:tcPr/></a:tc>' for value in values
    ) + '</a:tr>'
    table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Results table"/><p:cNvGraphicFramePr/><p:nvPr/>' \
        '</p:nvGraphicFramePr><p:xfrm><a:off x="600000" y="1800000"/><a:ext cx="7000000" cy="900000"/></p:xfrm>' \
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>' \
        '<a:tblGrid><a:gridCol w="3500000"/><a:gridCol w="3500000"/></a:tblGrid>' + \
        cells(['Feature', 'Status']) + cells(['Text and tables', 'Supported']) + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    slide2 = _xml('sld', '<p:cSld>' + _tree(_shape(2, 'Results', title=True, y=200000) + table) + '</p:cSld>' + overrides)
    master = _xml('sldMaster', '<p:cSld>' + _tree() + '</p:cSld>' + color_map +
                  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>'
                  '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:titleStyle>'
                  '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2200"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>')
    layout = _xml('sldLayout', '<p:cSld name="Blank">' + _tree() + '</p:cSld>' + overrides, 'type="blank" preserve="1"')
    colors = {'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '1F2937', 'lt2': 'F3F4F6', 'accent1': '2563EB',
              'accent2': '059669', 'accent3': 'D97706', 'accent4': '7C3AED', 'accent5': 'DC2626',
              'accent6': '0891B2', 'hlink': '0563C1', 'folHlink': '954F72'}
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    font = '<a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/>'
    theme = f'<a:theme xmlns:a="{A}" name="Sample"><a:themeElements><a:clrScheme name="Sample">' + ''.join(
        f'<a:{key}><a:srgbClr val="{value}"/></a:{key}>' for key, value in colors.items()) + \
        '</a:clrScheme><a:fontScheme name="Sample"><a:majorFont>' + font + '</a:majorFont><a:minorFont>' + font + \
        '</a:minorFont></a:fontScheme><a:fmtScheme name="Sample"><a:fillStyleLst>' + fill * 3 + \
        '</a:fillStyleLst><a:lnStyleLst>' + ('<a:ln w="9525">' + fill + '<a:prstDash val="solid"/></a:ln>') * 3 + \
        '</a:lnStyleLst><a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + \
        '</a:effectStyleLst><a:bgFillStyleLst>' + fill * 3 + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'
    parts = {
        '_rels/.rels': _rels([('main', 'officeDocument', 'ppt/presentation.xml')]),
        'ppt/presentation.xml': _xml('presentation', '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/>'
             '</p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="notesMaster"/></p:notesMasterIdLst>'
             '<p:sldIdLst><p:sldId id="256" r:id="slide1"/><p:sldId id="257" r:id="slide2"/></p:sldIdLst>'
             '<p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/>'),
        'ppt/_rels/presentation.xml.rels': _rels([('master', 'slideMaster', 'slideMasters/master.xml'),
             ('notesMaster', 'notesMaster', 'notesMasters/master.xml'), ('slide1', 'slide', 'slides/slide1.xml'),
             ('slide2', 'slide', 'slides/slide2.xml')]),
        'ppt/slides/slide1.xml': slide1,
        'ppt/slides/slide2.xml': slide2,
        'ppt/slides/_rels/slide1.xml.rels': _rels([('layout', 'slideLayout', '../slideLayouts/layout.xml'),
                                                  ('notes', 'notesSlide', '../notesSlides/notes1.xml')]),
        'ppt/slides/_rels/slide2.xml.rels': _rels([('layout', 'slideLayout', '../slideLayouts/layout.xml')]),
        'ppt/slideMasters/master.xml': master,
        'ppt/slideMasters/_rels/master.xml.rels': _rels([('layout', 'slideLayout', '../slideLayouts/layout.xml'), ('theme', 'theme', '../theme/theme1.xml')]),
        'ppt/slideLayouts/layout.xml': layout,
        'ppt/slideLayouts/_rels/layout.xml.rels': _rels([('master', 'slideMaster', '../slideMasters/master.xml')]),
        'ppt/theme/theme1.xml': theme,
        'ppt/notesSlides/notes1.xml': _xml('notes', '<p:cSld>' + _tree(_shape(2, 'This is a sample speaker note.')) + '</p:cSld>' + overrides),
        'ppt/notesSlides/_rels/notes1.xml.rels': _rels([('slide', 'slide', '../slides/slide1.xml'), ('master', 'notesMaster', '../notesMasters/master.xml')]),
        'ppt/notesMasters/master.xml': _xml('notesMaster', '<p:cSld>' + _tree() + '</p:cSld>' + color_map),
        'ppt/notesMasters/_rels/master.xml.rels': _rels([('theme', 'theme', '../theme/theme1.xml')]),
    }
    types = {'ppt/presentation.xml': 'presentation.main', 'ppt/theme/theme1.xml': None}
    for name in parts:
        if name.startswith('ppt/slides/') and name.endswith('.xml'): types[name] = 'slide'
        elif name.startswith('ppt/slideMasters/') and name.endswith('.xml'): types[name] = 'slideMaster'
        elif name.startswith('ppt/slideLayouts/') and name.endswith('.xml'): types[name] = 'slideLayout'
        elif name.startswith('ppt/notesSlides/') and name.endswith('.xml'): types[name] = 'notesSlide'
        elif name.startswith('ppt/notesMasters/') and name.endswith('.xml'): types[name] = 'notesMaster'
    parts['[Content_Types].xml'] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
        '<Default Extension="xml" ContentType="application/xml"/>' + ''.join(
            f'<Override PartName="/{name}" ContentType="application/vnd.openxmlformats-officedocument.' +
            (f'presentationml.{kind}+xml' if kind else 'theme+xml') + '"/>' for name, kind in types.items()) + '</Types>'
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, xml in parts.items():
            archive.writestr(name, xml)
