"""Relationship-aware DOCX to Markdown conversion with no runtime dependencies."""
from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass
from pathlib import Path

from mdify._markdown import inline, join_blocks, list_item, pipe_table
from mdify._ooxml import R, main_part, open_package, read_xml, related_part, relationships

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = R


def _qn(tag):
    return f'{{{W_NS}}}{tag}'


def _value(parent, name, default=None):
    elem = parent.find(_qn(name)) if parent is not None else None
    return elem.get(_qn('val'), default) if elem is not None else default


def _int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


@dataclass
class _Run:
    text: str
    bold: bool = False
    italic: bool = False
    link_url: str | None = None


class DocxParser:
    def __init__(self, archive):
        self.zip = archive
        self.part, self.root = main_part(archive, _qn('document'))
        self.relationships = relationships(archive, self.part)
        base = posixpath.dirname(self.part)
        styles = related_part(archive, self.relationships, 'styles', base + '/styles.xml')
        self.styles = {}
        self.default_style = None
        self.defaults = None
        if styles:
            root = read_xml(archive, styles)
            self.defaults = root.find(_qn('docDefaults'))
            for style in root.findall(_qn('style')):
                self.styles[style.get(_qn('styleId'))] = style
                if style.get(_qn('type')) == 'paragraph' and style.get(_qn('default')) in {'1', 'true'}:
                    self.default_style = style.get(_qn('styleId'))
        nums = related_part(archive, self.relationships, 'numbering', base + '/numbering.xml')
        self.numbering = self._numbering(read_xml(archive, nums)) if nums else {}
        self.list_counters = {}
        self.raw_text_parts = []

    def _style_chain(self, style_id):
        chain, seen = [], set()
        while style_id in self.styles:
            if style_id in seen:
                raise ValueError('Circular DOCX style inheritance')
            seen.add(style_id)
            style = self.styles[style_id]
            chain.append(style)
            style_id = _value(style, 'basedOn')
        return list(reversed(chain))

    def _numbering(self, root):
        abstract, result = {}, {}
        for definition in root.findall(_qn('abstractNum')):
            abstract[definition.get(_qn('abstractNumId'))] = {
                _int(level.get(_qn('ilvl'))): level for level in definition.findall(_qn('lvl'))
            }
        for num in root.findall(_qn('num')):
            levels = abstract.get(_value(num, 'abstractNumId'), {})
            overrides = {_int(o.get(_qn('ilvl'))): o for o in num.findall(_qn('lvlOverride'))}
            for depth in levels.keys() | overrides.keys():
                level = levels.get(depth)
                override = overrides.get(depth)
                if override is not None and override.find(_qn('lvl')) is not None:
                    level = override.find(_qn('lvl'))
                start = _int(_value(override, 'startOverride', _value(level, 'start', '1')), 1)
                # lvlRestart is one-based: 0 means never restart; absent means previous level.
                restart = _int(_value(level, 'lvlRestart'), depth) - 1
                result[(num.get(_qn('numId')), depth)] = (_value(level, 'numFmt', 'decimal'), start, restart)
        return result

    def _paragraph_properties(self, paragraph):
        direct = paragraph.find(_qn('pPr'))
        style_id = _value(direct, 'pStyle', self.default_style)
        chain = self._style_chain(style_id)
        layers = [style.find(_qn('pPr')) for style in chain] + [direct]
        if self.defaults is not None:
            layers.insert(0, self.defaults.find(f'{_qn("pPrDefault")}/{_qn("pPr")}'))
        props = {'num_id': None, 'level': 0, 'heading': None}
        for style in chain:
            name = _value(style, 'name', '')
            match = re.fullmatch(r'heading\s*([1-9])', name, re.I)
            if match:
                props['heading'] = min(6, int(match[1]))
        match = re.fullmatch(r'heading\s*([1-9])', style_id or '', re.I)
        if match:
            props['heading'] = min(6, int(match[1]))
        for layer in layers:
            if layer is None:
                continue
            outline = _value(layer, 'outlineLvl')
            if outline is not None:
                props['heading'] = min(6, _int(outline) + 1) if 0 <= _int(outline) <= 8 else None
            num = layer.find(_qn('numPr'))
            if num is not None:
                props['num_id'] = _value(num, 'numId', props['num_id'])
                props['level'] = max(0, min(8, _int(_value(num, 'ilvl', props['level']))))
        if props['num_id'] == '0':
            props['num_id'] = None
        return props, chain

    def _run_format(self, run, paragraph_styles):
        rpr = run.find(_qn('rPr'))
        layers = []
        if self.defaults is not None:
            layers.append(self.defaults.find(f'{_qn("rPrDefault")}/{_qn("rPr")}'))
        layers += [s.find(_qn('rPr')) for s in paragraph_styles]
        layers += [s.find(_qn('rPr')) for s in self._style_chain(_value(rpr, 'rStyle'))]
        values = {'b': False, 'i': False}
        for props in layers:
            for key in values:
                elem = props.find(_qn(key)) if props is not None else None
                # Word style b/i properties are toggles; direct formatting is absolute.
                if elem is not None and elem.get(_qn('val'), 'true').lower() not in {'0', 'false', 'off'}:
                    values[key] = not values[key]
        for key in values:
            elem = rpr.find(_qn(key)) if rpr is not None else None
            if elem is not None:
                values[key] = elem.get(_qn('val'), 'true').lower() not in {'0', 'false', 'off'}
        return values['b'], values['i']

    def _extract_runs(self, parent, styles, destination=None):
        runs = []
        for child in parent:
            if child.tag == _qn('r'):
                text = ''
                for elem in child:
                    if elem.tag == _qn('t'):
                        text += elem.text or ''
                    elif elem.tag in {_qn('br'), _qn('cr')}:
                        text += '\n'
                    elif elem.tag == _qn('tab'):
                        text += '\t'
                    elif elem.tag == _qn('noBreakHyphen'):
                        text += '\u2011'
                if text:
                    bold, italic = self._run_format(child, styles)
                    runs.append(_Run(text, bold, italic, destination))
            elif child.tag == _qn('hyperlink'):
                rel = self.relationships.get(child.get(f'{{{R}}}id'))
                anchor = child.get(_qn('anchor'))
                target = rel.target if rel and rel.kind == 'hyperlink' else '#' + anchor if anchor else None
                runs.extend(self._extract_runs(child, styles, target))
            elif child.tag in {_qn('sdt'), _qn('sdtContent'), _qn('ins'), _qn('moveTo'), _qn('smartTag'), _qn('customXml')}:
                runs.extend(self._extract_runs(child, styles, destination))
        return runs

    def _parse_paragraph(self, paragraph):
        props, styles = self._paragraph_properties(paragraph)
        runs = self._extract_runs(paragraph, styles)
        merged = []
        for run in runs:
            if merged and (run.bold, run.italic, run.link_url) == (merged[-1].bold, merged[-1].italic, merged[-1].link_url):
                merged[-1].text += run.text
            else:
                merged.append(run)
        raw = ''.join(r.text for r in merged)
        if not raw.strip():
            return '', False
        self.raw_text_parts.append(raw)
        md = ''.join('  \n'.join(inline(line, r.bold, r.italic, r.link_url) for line in r.text.split('\n')) for r in merged).strip()
        if props['heading']:
            return '#' * props['heading'] + ' ' + md, False
        num_id, level = props['num_id'], props['level']
        if num_id is not None:
            fmt, start, _ = self.numbering.get((num_id, level), ('bullet', 1, level - 1))
            for key in list(self.list_counters):
                if key[0] == num_id and key[1] > level:
                    restart = self.numbering.get(key, ('decimal', 1, key[1] - 1))[2]
                    if restart >= level:
                        del self.list_counters[key]
            key = (num_id, level)
            count = self.list_counters.get(key, start - 1) + 1
            self.list_counters[key] = count
            if fmt == 'none':
                return md, False
            marker = '-' if fmt == 'bullet' else f'{count}.'
            indent = sum(max(4, len(str(self.list_counters.get((num_id, depth), 0))) + 2)
                         for depth in range(level))
            return list_item(md, marker, indent), True
        return md, False

    def _blocks(self, parent):
        blocks = []
        for child in parent:
            if child.tag == _qn('p'):
                blocks.append(self._parse_paragraph(child))
            elif child.tag == _qn('tbl'):
                blocks.append((self._parse_table(child), False))
            elif child.tag in {_qn('sdt'), _qn('sdtContent'), _qn('ins'), _qn('moveTo'), _qn('customXml')}:
                blocks.extend(self._blocks(child))
        return blocks

    def _parse_table(self, table):
        rows = []
        for row in table.findall(_qn('tr')):
            cells = [''] * max(0, min(16384, _int(_value(row.find(_qn('trPr')), 'gridBefore'))))
            for cell in row.findall(_qn('tc')):
                props = cell.find(_qn('tcPr'))
                span = max(1, min(16384, _int(_value(props, 'gridSpan', '1'), 1)))
                cells.append(join_blocks(self._blocks(cell)))
                cells.extend([''] * (span - 1))
                if len(cells) > 16384:
                    raise ValueError('DOCX table exceeds 16,384 columns')
            rows.append(cells)
        return pipe_table(rows)

    def convert(self):
        body = self.root.find(_qn('body'))
        if body is None:
            raise ValueError('DOCX document has no body')
        md = join_blocks(self._blocks(body))
        return md, '\n'.join(self.raw_text_parts).strip()


def convert_docx(docx_bytes_or_file) -> tuple[str, str]:
    with open_package(docx_bytes_or_file) as archive:
        return DocxParser(archive).convert()


def convert_docx_file(file_path: str | Path) -> tuple[str, str]:
    return convert_docx(file_path)
