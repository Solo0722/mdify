"""HTML to Markdown using a small, dependency-free tree renderer."""
from __future__ import annotations

import re
from html.parser import HTMLParser
from pathlib import Path

from mdify._markdown import code_block, code_span, escape_text, link, list_item, pipe_table
from mdify._text import decode_text

VOID_TAGS = set('area base br col embed hr img input link meta param source track wbr'.split())
IGNORED_TAGS = {'script', 'style', 'noscript', 'head'}
BLOCK_TAGS = set('p div section article header footer main aside nav h1 h2 h3 h4 h5 h6 pre blockquote ul ol table hr dl dt dd'.split())


class HtmlNode:
    def __init__(self, tag, attrs=None, parent=None):
        self.tag, self.attrs, self.parent = tag, attrs or {}, parent
        self.children = []


class HtmlTreeBuilder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = self.current = HtmlNode('root')
        self.ignored = []

    def _close_ancestor(self, tags, boundaries=()):
        node = self.current
        while node.parent:
            if node.tag in tags:
                self.current = node.parent
                return
            if node.tag in boundaries:
                return
            node = node.parent

    def handle_starttag(self, tag, attrs):
        if tag in IGNORED_TAGS:
            self.ignored.append(tag)
            return
        if self.ignored:
            return
        if tag == 'li':
            self._close_ancestor({'li'}, {'ul', 'ol'})
        if tag in {'td', 'th'}:
            self._close_ancestor({'td', 'th'}, {'tr', 'table'})
        if tag == 'tr':
            self._close_ancestor({'tr'}, {'table'})
        if tag in BLOCK_TAGS:
            self._close_ancestor({'p'}, {'li', 'td', 'th', 'blockquote'})
        node = HtmlNode(tag, dict(attrs), self.current)
        self.current.children.append(node)
        if tag not in VOID_TAGS:
            self.current = node

    def handle_endtag(self, tag):
        if self.ignored:
            if tag in self.ignored:
                while self.ignored:
                    if self.ignored.pop() == tag:
                        break
            return
        if tag not in VOID_TAGS:
            self._close_ancestor({tag})

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID_TAGS:
            self.handle_endtag(tag)

    def handle_data(self, data):
        if not self.ignored:
            self.current.children.append(data)


def _raw(node):
    if isinstance(node, str):
        return node
    if node.tag == 'img':
        return node.attrs.get('alt') or ''
    if node.tag == 'br':
        return '\n'
    result = ''.join(_raw(child) for child in node.children)
    if node.tag in {'td', 'th'}:
        return result + '\t'
    if node.tag in BLOCK_TAGS | {'li', 'tr'}:
        return result + '\n'
    return result


def _literal(node):
    """Text of code nodes, with no Markdown or prose whitespace processing."""
    if isinstance(node, str):
        return node
    if node.tag == 'br':
        return '\n'
    return ''.join(_literal(c) for c in node.children)


def _wrap(text, marker):
    match = re.fullmatch(r'(\s*)(.*?)(\s*)', text, re.S)
    if not match or not match[2]:
        return text
    return match[1] + marker + match[2] + marker + match[3]


def _children(node):
    """Join inline runs normally; separate block children without touching code."""
    chunks, pending = [], ''
    for child in node.children:
        if isinstance(child, HtmlNode) and child.tag in BLOCK_TAGS:
            if pending.strip():
                chunks.append(pending.strip())
            pending = ''
            rendered = _render(child)
            if rendered:
                chunks.append(rendered)
        else:
            pending += _render(child)
    if chunks:
        if pending.strip():
            chunks.append(pending.strip())
        return '\n\n'.join(chunks)
    return pending


def _positive_int(value, default=1, maximum=1000):
    try:
        return max(1, min(maximum, int(value)))
    except (ValueError, TypeError):
        return default


def _table(node):
    rows, spans = [], {}
    row_nodes = []
    def collect(parent):
        for child in parent.children:
            if isinstance(child, HtmlNode):
                if child.tag == 'tr':
                    row_nodes.append(child)
                elif child.tag in {'thead', 'tbody', 'tfoot'}:
                    collect(child)
    collect(node)
    for row_node in row_nodes:
        cells, col = [], 0
        next_spans = {c: left - 1 for c, left in spans.items() if left > 1}
        for cell in row_node.children:
            if not isinstance(cell, HtmlNode) or cell.tag not in {'td', 'th'}:
                continue
            while col in spans:
                cells.append('')
                col += 1
            width = _positive_int(cell.attrs.get('colspan'))
            height = _positive_int(cell.attrs.get('rowspan'))
            if col + width > 16384:
                raise ValueError('HTML table exceeds 16,384 columns')
            cells.append(_children(cell).strip())
            cells.extend([''] * (width - 1))
            if height > 1:
                for c in range(col, col + width):
                    next_spans[c] = height - 1
            col += width
        if spans:
            cells.extend([''] * max(0, max(spans) + 1 - len(cells)))
        spans = next_spans
        if cells:
            rows.append(cells)
    caption = next((c for c in node.children if isinstance(c, HtmlNode) and c.tag == 'caption'), None)
    table = pipe_table(rows)
    return (_children(caption).strip() + '\n\n' if caption else '') + table


def _render(node):
    if isinstance(node, str):
        return escape_text(re.sub(r'\s+', ' ', node))
    tag = node.tag
    if tag in IGNORED_TAGS:
        return ''
    if tag == 'pre':
        language = ''
        for child in node.children:
            if isinstance(child, HtmlNode) and child.tag == 'code':
                match = re.search(r'(?:lang-|language-)([\w+-]+)', child.attrs.get('class') or '')
                if match:
                    language = match[1]
        return code_block(_literal(node), language)
    if tag == 'code':
        return code_span(_literal(node))
    if tag == 'br':
        return '  \n'
    if tag == 'hr':
        return '---'
    if tag == 'table':
        return _table(node)
    if tag in {'ul', 'ol'}:
        items = []
        try:
            index = int(node.attrs.get('start') or '1')
        except ValueError:
            index = 1
        for child in node.children:
            if not isinstance(child, HtmlNode) or child.tag != 'li':
                continue
            try:
                index = int(child.attrs.get('value', index))
            except (ValueError, TypeError):
                pass
            content = ''
            pending = ''
            for item in child.children:
                if isinstance(item, HtmlNode) and item.tag in {'ul', 'ol'}:
                    if pending.strip():
                        content += ('\n\n' if content else '') + pending.strip()
                    pending = ''
                    nested = _render(item)
                    content += ('\n' if content else '') + nested
                elif isinstance(item, HtmlNode) and item.tag in BLOCK_TAGS:
                    if pending.strip():
                        content += ('\n\n' if content else '') + pending.strip()
                    pending = ''
                    content += ('\n\n' if content else '') + _render(item)
                else:
                    pending += _render(item)
            if pending.strip():
                content += ('\n\n' if content else '') + pending.strip()
            items.append(list_item(content, f'{index}.' if tag == 'ol' else '-'))
            index += 1
        return '\n'.join(items)
    inner = _children(node)
    if tag in {'b', 'strong'}:
        return _wrap(inner, '**')
    if tag in {'i', 'em'}:
        return _wrap(inner, '*')
    if tag in {'s', 'del'}:
        return _wrap(inner, '~~')
    if tag == 'img':
        return '!' + link(escape_text(node.attrs.get('alt') or ''), node.attrs.get('src') or '')
    if tag == 'a':
        href = node.attrs.get('href')
        return link(inner.strip() or escape_text(href), href) if href else inner
    if tag == 'blockquote':
        return '\n'.join('> ' + line if line else '>' for line in inner.strip().split('\n'))
    if re.fullmatch('h[1-6]', tag):
        return '#' * int(tag[1]) + ' ' + inner.strip()
    if tag in BLOCK_TAGS:
        return inner.strip()
    return inner


def convert_html(html_content: str) -> tuple[str, str]:
    builder = HtmlTreeBuilder()
    builder.feed(html_content.lstrip('\ufeff'))
    builder.close()
    return _children(builder.root).strip(), _raw(builder.root).strip()


def convert_html_file(file_path: str | Path, *, encoding=None) -> tuple[str, str]:
    return convert_html(decode_text(Path(file_path).read_bytes(), encoding, html=True))
