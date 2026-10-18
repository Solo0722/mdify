"""CSV and XLSX to Markdown tables, preserving values and package relationships."""
from __future__ import annotations

import csv
import io
import posixpath
import re
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from mdify._markdown import MAX_TABLE_CELLS, escape_text, pipe_table
from mdify._ooxml import (R, main_part, open_package, read_xml, related_part,
                          relationships)
from mdify._text import decode_text

SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = R


def _qn(tag):
    return f'{{{SS_NS}}}{tag}'


def _col_letters_to_index(letters):
    value = 0
    for letter in letters.upper():
        value = value * 26 + ord(letter) - ord('A') + 1
    return value - 1


def _format_pipe_table(rows):
    return pipe_table([[escape_text(str(cell)) for cell in row] for row in rows])


def convert_csv(csv_content: str, *, delimiter: str | None = None) -> tuple[str, str]:
    csv_content = csv_content.lstrip('\ufeff')
    if delimiter is not None and (len(delimiter) != 1 or delimiter in '\r\n\x00"'):
        raise ValueError('CSV delimiter must be one non-newline character other than a quote')
    if not csv_content.strip():
        return '', ''
    options = {'delimiter': delimiter} if delimiter else {}
    dialect = csv.excel
    if delimiter is None:
        try:
            dialect = csv.Sniffer().sniff(csv_content[:4096], delimiters=',;\t|')
        except csv.Error:
            pass  # Never guess an arbitrary letter as a delimiter.
    reader = csv.reader(io.StringIO(csv_content, newline=''), dialect=dialect, **options)
    rows = [row for row in reader if all(cell.strip() for cell in row)]
    return _format_pipe_table(rows), '\n'.join(' '.join(cell.strip() for cell in row) for row in rows)


def convert_csv_file(file_path: str | Path, *, delimiter=None, encoding=None) -> tuple[str, str]:
    return convert_csv(decode_text(Path(file_path).read_bytes(), encoding), delimiter=delimiter)


_BUILTIN_FORMATS = {
    14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yy h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0',
}


def _date_kind(format_code):
    code = format_code.split(';')[0].lower()
    code = re.sub(r'"[^\"]*"|\\.|_.|\*.', '', code)
    elapsed = bool(re.search(r'\[[hms]+\]', code))
    code = re.sub(r'\[[^]]*\]', '', code)
    date = bool(re.search('[yd]', code))
    time = elapsed or bool(re.search('[hs]', code))
    date = date or ('m' in code and not time)
    return date, time, elapsed


def _format_date(value, code, date1904):
    date, time, elapsed = _date_kind(code)
    if not (date or time):
        return value
    try:
        serial = Decimal(value)
        if not serial.is_finite():
            return value
        seconds = int((serial * 86400).to_integral_value())
        if elapsed:
            sign = '-' if seconds < 0 else ''
            hours, rest = divmod(abs(seconds), 3600)
            minutes, seconds = divmod(rest, 60)
            return f'{sign}{hours}:{minutes:02}:{seconds:02}'
        # Excel retains the fictional 1900-02-29 for serial 60.
        fictional = not date1904 and 60 <= serial < 61
        base = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 31)
        if not date1904 and serial >= 60:
            seconds -= 86400
        result = base + timedelta(seconds=seconds)
        date_text = '1900-02-29' if fictional else result.strftime('%Y-%m-%d')
        time_text = result.strftime('%H:%M:%S')
        return date_text + ' ' + time_text if date and time else date_text if date else time_text
    except (InvalidOperation, ValueError, OverflowError):
        return value


class XlsxParser:
    def __init__(self, archive):
        self.zip = archive
        self.part, self.root = main_part(archive, _qn('workbook'))
        self.relationships = relationships(archive, self.part)
        base = posixpath.dirname(self.part)
        strings = related_part(archive, self.relationships, 'sharedStrings', base + '/sharedStrings.xml')
        self.shared_strings = []
        if strings:
            for item in read_xml(archive, strings).findall(_qn('si')):
                self.shared_strings.append(self._string(item))
        self.formats = []
        styles = related_part(archive, self.relationships, 'styles', base + '/styles.xml')
        if styles:
            style_root = read_xml(archive, styles)
            formats = dict(_BUILTIN_FORMATS)
            for num_format in style_root.findall(f'{_qn("numFmts")}/{_qn("numFmt")}'):
                formats[int(num_format.get('numFmtId'))] = num_format.get('formatCode', '')
            for style in style_root.findall(f'{_qn("cellXfs")}/{_qn("xf")}'):
                self.formats.append(formats.get(int(style.get('numFmtId', '0')), ''))
        props = self.root.find(_qn('workbookPr'))
        self.date1904 = props is not None and props.get('date1904') in {'1', 'true'}

    @staticmethod
    def _string(node):
        # Phonetic annotations are not part of the displayed cell value.
        parts = []
        for child in node:
            if child.tag == _qn('t'):
                parts.append(child.text or '')
            elif child.tag == _qn('r'):
                parts.extend(t.text or '' for t in child.findall(_qn('t')))
        return ''.join(parts)

    def _extract_cell_value(self, cell):
        value = cell.find(_qn('v'))
        raw = value.text if value is not None and value.text is not None else ''
        formula = cell.find(_qn('f'))
        if formula is not None and not raw:
            return '=' + formula.text if formula.text else '[Formula result unavailable]'
        kind = cell.get('t')
        if kind == 's':
            try:
                index = int(raw)
                if not 0 <= index < len(self.shared_strings):
                    raise ValueError()
                return self.shared_strings[index]
            except ValueError as exc:
                raise ValueError(f'Invalid shared string index {raw!r} in cell {cell.get("r", "?")}') from exc
        if kind == 'inlineStr':
            item = cell.find(_qn('is'))
            return self._string(item) if item is not None else ''
        if kind == 'b':
            if raw not in {'0', '1', ''}:
                raise ValueError(f'Invalid Boolean value {raw!r}')
            return 'TRUE' if raw == '1' else 'FALSE' if raw == '0' else ''
        if raw and kind in {None, 'n'} and 's' in cell.attrib:
            style = int(cell.get('s'))
            if not 0 <= style < len(self.formats):
                raise ValueError(f'Invalid style index {style}')
            return _format_date(raw, self.formats[style], self.date1904)
        return raw

    def _parse_worksheet(self, part):
        root = read_xml(self.zip, part)
        if root.tag != _qn('worksheet'):
            raise ValueError(f'Expected worksheet XML in {part}')
        data = root.find(_qn('sheetData'))
        if data is None:
            return []
        rows, row_index, max_col = {}, 0, -1
        for row in data.findall(_qn('row')):
            row_index = int(row.get('r', row_index + 1))
            col_index = 0
            for cell in row.findall(_qn('c')):
                coord = cell.get('r')
                actual_row = row_index
                if coord:
                    match = re.fullmatch(r'([A-Za-z]+)([1-9]\d*)', coord)
                    if not match:
                        raise ValueError(f'Invalid cell coordinate: {coord}')
                    col_index = _col_letters_to_index(match[1])
                    actual_row = int(match[2])
                if not (1 <= actual_row <= 1048576 and 0 <= col_index < 16384):
                    raise ValueError(f'Cell coordinate outside XLSX limits: {coord}')
                value = self._extract_cell_value(cell)
                if value:
                    rows.setdefault(actual_row, {})[col_index] = value
                    max_col = max(max_col, col_index)
                col_index += 1
        # Only materialize populated rows, not the rectangle from A1 to the last cell.
        if len(rows) * (max_col + 1) > MAX_TABLE_CELLS:
            raise ValueError(f'Worksheet exceeds the {MAX_TABLE_CELLS:,}-cell output limit')
        return [[values.get(c, '') for c in range(max_col + 1)] for _, values in sorted(rows.items())]

    def convert(self):
        sections, raw = [], []
        sheets = self.root.find(_qn('sheets'))
        if sheets is None:
            raise ValueError('Workbook has no sheets list')
        for sheet in sheets.findall(_qn('sheet')):
            name = sheet.get('name', 'Sheet')
            rel = self.relationships.get(sheet.get(f'{{{R}}}id'))
            if rel is None or rel.external:
                raise ValueError(f'Missing worksheet relationship for {name}')
            raw.append(name)
            if rel.kind != 'worksheet':
                sections.append(f'## {escape_text(name)}\n\n[Unsupported {escape_text(rel.kind)} sheet]')
                continue
            rows = self._parse_worksheet(rel.target)
            sections.append(f'## {escape_text(name)}\n\n' + (_format_pipe_table(rows) or '(Empty sheet)'))
            raw.extend(' '.join(row) for row in rows)
        return '\n\n'.join(sections), '\n'.join(raw).strip()


def convert_xlsx(xlsx_bytes_or_file) -> tuple[str, str]:
    with open_package(xlsx_bytes_or_file) as archive:
        return XlsxParser(archive).convert()


def convert_xlsx_file(file_path: str | Path) -> tuple[str, str]:
    return convert_xlsx(file_path)
