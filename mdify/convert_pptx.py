"""PPTX content extraction: slides, text, tables, notes, and visual descriptions.

Uses stored shape order after the title, not a geometric reading-order guess.
"""
from __future__ import annotations

from pathlib import Path

from mdify._markdown import escape_text, inline, join_blocks, list_item, pipe_table
from mdify._ooxml import R, main_part, open_package, read_xml, related_part, relationships

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'p': P, 'a': A, 'r': R}


def _find(node, path):
    return node.find(path, NS) if node is not None else None


def _integer(value, default=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def _placeholder(shape):
    return _find(shape, 'p:nvSpPr/p:nvPr/p:ph')


def _shapes(root):
    tree = _find(root, 'p:cSld/p:spTree')
    def walk(parent):
        if parent is not None:
            for child in parent:
                if child.tag == f'{{{P}}}grpSp':
                    yield from walk(child)
                elif child.tag in {f'{{{P}}}{name}' for name in ('sp', 'pic', 'graphicFrame', 'cxnSp')}:
                    yield child
                elif child.tag.endswith('}AlternateContent'):
                    # Prefer the DrawingML representation; never emit Choice and Fallback twice.
                    choice = next((c for c in child if c.tag.endswith('}Choice')), None)
                    if choice is None:
                        choice = next((c for c in child if c.tag.endswith('}Fallback')), None)
                    yield from walk(choice)
    return list(walk(tree))


class PptxParser:
    def __init__(self, archive, *, include_notes=True):
        self.zip = archive
        self.part, self.root = main_part(archive, f'{{{P}}}presentation')
        self.include_notes = include_notes
        self.raw = []
        self.cache = {}
        rels = relationships(archive, self.part)
        ids = _find(self.root, 'p:sldIdLst')
        self.slides = []
        if ids is not None:
            for slide_id in ids:
                rel = rels.get(slide_id.get(f'{{{R}}}id'))
                if rel is None or rel.kind != 'slide' or rel.external:
                    raise ValueError('Missing or invalid PPTX slide relationship')
                self.slides.append(rel.target)
        self.slide_numbers = {part: i + 1 for i, part in enumerate(self.slides)}

    def _part(self, path):
        if path not in self.cache:
            self.cache[path] = (read_xml(self.zip, path), relationships(self.zip, path))
        return self.cache[path]

    def _inheritance(self, slide_rels):
        layout_path = related_part(self.zip, slide_rels, 'slideLayout')
        layout, master = None, None
        if layout_path:
            layout, layout_rels = self._part(layout_path)
            master_path = related_part(self.zip, layout_rels, 'slideMaster')
            if master_path:
                master, _ = self._part(master_path)
        return layout, master

    def _matching_placeholders(self, shape, layout, master):
        ph = _placeholder(shape)
        if ph is None:
            return None, None, None
        index = ph.get('idx', '0')
        layout_shape = next((s for s in _shapes(layout) if _placeholder(s) is not None and
                             _placeholder(s).get('idx', '0') == index), None)
        layout_ph = _placeholder(layout_shape) if layout_shape is not None else None
        kind = ph.get('type') or (layout_ph.get('type', 'obj') if layout_ph is not None else 'obj')
        master_kind = 'title' if kind == 'ctrTitle' else 'body' if kind in {'obj', 'subTitle'} else kind
        master_shape = next((s for s in _shapes(master) if _placeholder(s) is not None and
                             _placeholder(s).get('type', 'obj') == master_kind), None)
        return layout_shape, master_shape, kind

    def _paragraph_styles(self, shape, paragraph, layout, master):
        direct = _find(paragraph, 'a:pPr')
        level = max(0, min(8, _integer(direct.get('lvl') if direct is not None else None)))
        layout_shape, master_shape, kind = self._matching_placeholders(shape, layout, master)
        style_name = 'titleStyle' if kind in {'title', 'ctrTitle'} else 'bodyStyle' if kind in {'body', 'obj', 'subTitle'} else 'otherStyle'
        containers = [_find(self.root, 'p:defaultTextStyle'), _find(master, f'p:txStyles/p:{style_name}')]
        for source in (master_shape, layout_shape, shape):
            containers.append(_find(source, 'p:txBody/a:lstStyle'))
        layers = []
        for container in containers:
            if container is not None:
                layers.extend([_find(container, 'a:defPPr'), _find(container, f'a:lvl{level + 1}pPr')])
        layers.append(direct)
        return level, [p for p in layers if p is not None]

    def _destination(self, props, rels):
        hlink = _find(props, 'a:hlinkClick')
        if hlink is None:
            return None
        rel = rels.get(hlink.get(f'{{{R}}}id'))
        if rel is None:
            return None
        if rel.kind == 'slide' and rel.target in self.slide_numbers:
            # Each slide has a stable explicit anchor independent of its title.
            return '#slide-' + str(self.slide_numbers[rel.target])
        if rel.kind == 'hyperlink':
            return rel.target
        return None

    def _paragraph(self, paragraph, shape, rels, layout, master, counters):
        level, layers = self._paragraph_styles(shape, paragraph, layout, master)
        base = {'b': False, 'i': False}
        for layer in layers:
            props = _find(layer, 'a:defRPr')
            if props is not None:
                for key in base:
                    if key in props.attrib:
                        base[key] = props.get(key) in {'1', 'true', 'on'}
        runs = []
        for child in paragraph:
            if child.tag == f'{{{A}}}br':
                runs.append(['\n', False, False, None])
            elif child.tag in {f'{{{A}}}r', f'{{{A}}}fld'}:
                props = _find(child, 'a:rPr')
                values = dict(base)
                if props is not None:
                    for key in values:
                        if key in props.attrib:
                            values[key] = props.get(key) in {'1', 'true', 'on'}
                text = ''.join(t.text or '' for t in child.findall('a:t', NS))
                destination = self._destination(props, rels)
                data = [text, values['b'], values['i'], destination]
                if runs and runs[-1][1:] == data[1:]:
                    runs[-1][0] += text
                else:
                    runs.append(data)
        raw = ''.join(run[0] for run in runs)
        if not raw.strip():
            return '', False
        self.raw.append(raw)
        text = ''.join('  \n'.join(inline(line, bold, italic, url) for line in value.split('\n'))
                       for value, bold, italic, url in runs).strip()
        bullet = None
        for layer in layers:
            for name in ('buNone', 'buChar', 'buAutoNum', 'buBlip'):
                elem = _find(layer, 'a:' + name)
                if elem is not None:
                    bullet = (name, elem)
                    break
        if bullet is None or bullet[0] == 'buNone':
            counters.clear()
            return text, False
        if bullet[0] == 'buAutoNum':
            elem = bullet[1]
            style = elem.get('type', 'arabicPeriod')
            start = _integer(elem.get('startAt'), 1)
            for depth in list(counters):
                if depth > level:
                    del counters[depth]
            prior = counters.get(level)
            number = prior[1] + 1 if prior and prior[0] == style and prior[2] == start else start
            counters[level] = (style, number, start)
            marker = f'{number}.'
        else:
            counters.pop(level, None)
            marker = '-'
        indent = sum(max(4, len(str(counters[depth][1])) + 2) if depth in counters else 4
                     for depth in range(level))
        return list_item(text, marker, indent), True

    def _text(self, shape, rels, layout=None, master=None, body=None):
        body = body if body is not None else _find(shape, 'p:txBody')
        if body is None:
            return ''
        counters = {}
        return join_blocks([self._paragraph(p, shape, rels, layout, master, counters)
                            for p in body.findall('a:p', NS)])

    def _table(self, table, rels):
        rows = []
        for row in table.findall('a:tr', NS):
            cells = []
            for cell in row.findall('a:tc', NS):
                # DrawingML includes continuation cells for both horizontal and vertical merges.
                merged = cell.get('hMerge') in {'1', 'true'} or cell.get('vMerge') in {'1', 'true'}
                text = self._text(cell, rels, body=_find(cell, 'a:txBody'))
                cells.append('' if merged and not text else text)
            rows.append(cells)
        return pipe_table(rows)

    def _description(self, shape):
        props = next((e for e in shape.iter() if e.tag == f'{{{P}}}cNvPr'), None)
        if props is None:
            return ''
        return props.get('descr') or props.get('title') or ''

    def _visual_marker(self, label, shape):
        description = self._description(shape)
        if description:
            self.raw.append(description)
        return '[' + label + (': ' + escape_text(description) if description else '') + ']'

    def _shape(self, shape, rels, layout=None, master=None):
        if shape.tag == f'{{{P}}}sp':
            return self._text(shape, rels, layout, master)
        if shape.tag == f'{{{P}}}pic':
            media = any(e.tag.rsplit('}', 1)[-1] in {'videoFile', 'audioFile', 'media'} for e in shape.iter())
            return self._visual_marker('Media omitted' if media else 'Image', shape)
        if shape.tag == f'{{{P}}}graphicFrame':
            table = _find(shape, 'a:graphic/a:graphicData/a:tbl')
            if table is not None:
                return self._table(table, rels)
            data = _find(shape, 'a:graphic/a:graphicData')
            uri = data.get('uri', '') if data is not None else ''
            label = 'Chart omitted' if 'chart' in uri else 'SmartArt omitted' if 'diagram' in uri else 'Embedded object omitted'
            return self._visual_marker(label, shape)
        return ''

    def _notes(self, slide_rels):
        path = related_part(self.zip, slide_rels, 'notesSlide')
        if not path:
            return ''
        root, rels = self._part(path)
        if root.tag != f'{{{P}}}notes':
            raise ValueError(f'Expected notes XML in {path}')
        blocks = []
        for shape in _shapes(root):
            ph = _placeholder(shape)
            if ph is not None and ph.get('type') != 'body':
                continue  # Slide image, page number, date, header and footer are not speaker notes.
            text = self._shape(shape, rels)
            if text:
                blocks.append(text)
        return '\n\n'.join(blocks)

    def convert(self):
        sections = []
        for number, path in enumerate(self.slides, 1):
            slide, rels = self._part(path)
            if slide.tag != f'{{{P}}}sld':
                raise ValueError(f'Expected slide XML in {path}')
            layout, master = self._inheritance(rels)
            shapes = _shapes(slide)
            title_shape = next((s for s in shapes if self._matching_placeholders(s, layout, master)[2]
                                in {'title', 'ctrTitle'}), None)
            title = self._text(title_shape, rels, layout, master) if title_shape is not None else ''
            # Titles are a single heading; exclude list markers inherited from unusual templates.
            title = title.replace('  \n', ' ').replace('\n', ' ')
            heading = f'## Slide {number}' + (': ' + title if title else '')
            if slide.get('show') in {'0', 'false', 'off'}:
                heading += ' (hidden)'
            blocks = []
            for shape in shapes:
                if shape is title_shape:
                    continue
                ph = _placeholder(shape)
                if ph is not None and ph.get('type') in {'dt', 'ftr', 'hdr', 'sldNum'}:
                    continue
                text = self._shape(shape, rels, layout, master)
                if text:
                    blocks.append(text)
            if not title and not blocks:
                blocks.append('(Empty slide)')
            if self.include_notes:
                notes = self._notes(rels)
                if notes:
                    blocks.append('### Speaker notes\n\n' + notes)
            sections.append(f'<a id="slide-{number}"></a>\n\n' + heading + ('\n\n' + '\n\n'.join(blocks) if blocks else ''))
        return '\n\n'.join(sections) or '(Empty presentation)', '\n'.join(self.raw).strip()


def convert_pptx(pptx_bytes_or_file, *, include_notes=True) -> tuple[str, str]:
    with open_package(pptx_bytes_or_file) as archive:
        return PptxParser(archive, include_notes=include_notes).convert()


def convert_pptx_file(file_path: str | Path, *, include_notes=True) -> tuple[str, str]:
    return convert_pptx(file_path, include_notes=include_notes)
