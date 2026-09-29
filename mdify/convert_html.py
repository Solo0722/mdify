"""HTML to Markdown converter using standard library html.parser."""

from __future__ import annotations

import html
import re
from html.parser import HTMLParser
from pathlib import Path
from typing import List, Optional, Union

VOID_TAGS = {
    "area",
    "base",
    "br",
    "col",
    "embed",
    "hr",
    "img",
    "input",
    "link",
    "meta",
    "param",
    "source",
    "track",
    "wbr",
}

IGNORED_TAGS = {"script", "style", "noscript", "head"}

BLOCK_CONTAINERS = {
    "root",
    "html",
    "body",
    "div",
    "section",
    "article",
    "header",
    "footer",
    "main",
    "aside",
    "nav",
    "table",
    "thead",
    "tbody",
    "tfoot",
    "tr",
    "ul",
    "ol",
}


class HtmlNode:
    """Lightweight DOM node for HTML AST representation."""

    def __init__(
        self,
        tag: str,
        attrs: Optional[dict[str, Optional[str]]] = None,
        parent: Optional[HtmlNode] = None,
    ):
        self.tag = tag.lower()
        self.attrs: dict[str, Optional[str]] = attrs or {}
        self.parent = parent
        self.children: List[Union[HtmlNode, str]] = []

    def add_child(self, child: Union[HtmlNode, str]) -> None:
        self.children.append(child)


class HtmlTreeBuilder(HTMLParser):
    """HTMLParser that builds a clean HtmlNode tree, ignoring script/style."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = HtmlNode("root")
        self.current = self.root
        self.ignore_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        tag_lower = tag.lower()
        if tag_lower in IGNORED_TAGS:
            self.ignore_depth += 1
            return

        if self.ignore_depth > 0:
            return

        attr_dict = {k.lower(): v for k, v in attrs}

        # Auto-closing tags for standard HTML lists/tables/paragraphs
        if tag_lower == "p" and self.current.tag == "p":
            if self.current.parent:
                self.current = self.current.parent

        if tag_lower == "li" and self.current.tag == "li":
            if self.current.parent:
                self.current = self.current.parent

        if tag_lower in ("td", "th") and self.current.tag in ("td", "th"):
            if self.current.parent:
                self.current = self.current.parent

        if tag_lower == "tr" and self.current.tag in ("tr", "td", "th"):
            while self.current.tag in ("td", "th", "tr") and self.current.parent:
                self.current = self.current.parent

        node = HtmlNode(tag_lower, attr_dict, parent=self.current)
        self.current.add_child(node)

        if tag_lower not in VOID_TAGS:
            self.current = node

    def handle_endtag(self, tag: str) -> None:
        tag_lower = tag.lower()
        if tag_lower in IGNORED_TAGS:
            if self.ignore_depth > 0:
                self.ignore_depth -= 1
            return

        if self.ignore_depth > 0:
            return

        if tag_lower in VOID_TAGS:
            return

        # Pop up until we match the tag or reach root
        walker = self.current
        found = False
        while walker is not None and walker.tag != "root":
            if walker.tag == tag_lower:
                found = True
                break
            walker = walker.parent

        if found:
            while self.current is not None and self.current.tag != tag_lower:
                if self.current.parent is not None:
                    self.current = self.current.parent
                else:
                    break
            if self.current is not None and self.current.parent is not None:
                self.current = self.current.parent

    def handle_data(self, data: str) -> None:
        if self.ignore_depth > 0:
            return
        if data:
            self.current.add_child(data)


def _wrap_inline_formatting(rendered: str, marker: str) -> str:
    """Wrap rendered text with Markdown inline marker (** or * or `).

    Preserves leading and trailing whitespace outside the markers.
    """
    if not rendered:
        return ""

    match = re.match(r"^(\s*)(.*?)(\s*)$", rendered, re.DOTALL)
    if not match:
        return f"{marker}{rendered}{marker}"

    leading_space, core, trailing_space = match.groups()
    if not core:
        return rendered

    return f"{leading_space}{marker}{core}{marker}{trailing_space}"


def _render_node_to_markdown(node: Union[HtmlNode, str], list_depth: int = 0) -> str:
    """Recursively render an HtmlNode or text string to Markdown."""
    if isinstance(node, str):
        return node

    tag = node.tag

    if tag in IGNORED_TAGS:
        return ""

    # Helper to render children, skipping whitespace text between block elements
    def render_children(parent_node: HtmlNode) -> str:
        parts = []
        is_block_parent = parent_node.tag in BLOCK_CONTAINERS
        for child in parent_node.children:
            if is_block_parent and isinstance(child, str) and not child.strip():
                continue
            parts.append(_render_node_to_markdown(child, list_depth))
        return "".join(parts)

    # Headings h1-h6
    if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
        level = int(tag[1])
        inner = render_children(node).strip()
        if not inner:
            return ""
        return f"\n\n{'#' * level} {inner}\n\n"

    # Paragraphs and divisions
    if tag in ("p", "div", "section", "article"):
        inner = render_children(node).strip()
        if not inner:
            return ""
        return f"\n\n{inner}\n\n"

    # Line break and horizontal rule
    if tag == "br":
        return "  \n"
    if tag == "hr":
        return "\n\n---\n\n"

    # Bold
    if tag in ("b", "strong"):
        inner = render_children(node)
        return _wrap_inline_formatting(inner, "**")

    # Italic
    if tag in ("i", "em"):
        inner = render_children(node)
        return _wrap_inline_formatting(inner, "*")

    # Code block <pre>
    if tag == "pre":
        lang = ""
        for child in node.children:
            if isinstance(child, HtmlNode) and child.tag == "code":
                cls = child.attrs.get("class", "") or ""
                m = re.search(r"(?:lang-|language-)(\w+)", cls)
                if m:
                    lang = m.group(1)

        raw_code = _extract_raw_text(node)
        raw_code = raw_code.rstrip("\r\n")
        return f"\n\n```{lang}\n{raw_code}\n```\n\n"

    # Inline code <code> (when not inside <pre>)
    if tag == "code":
        inner = render_children(node)
        return _wrap_inline_formatting(inner, "`")

    # Image <img>
    if tag == "img":
        src = node.attrs.get("src", "") or ""
        alt = node.attrs.get("alt", "") or ""
        return f"![{alt}]({src})"

    # Link <a>
    if tag == "a":
        href = node.attrs.get("href", "") or ""
        inner = render_children(node).strip()
        if not inner and not href:
            return ""
        if not href:
            return inner
        if not inner:
            return f"[{href}]({href})"
        return f"[{inner}]({href})"

    # Unordered list <ul>
    if tag == "ul":
        items = []
        for child in node.children:
            if isinstance(child, HtmlNode) and child.tag == "li":
                item_md = _render_list_item(child, is_ordered=False, index=0, depth=list_depth)
                if item_md:
                    items.append(item_md)
        if not items:
            return ""
        return "\n\n" + "\n".join(items) + "\n\n"

    # Ordered list <ol>
    if tag == "ol":
        items = []
        try:
            start_idx = int(node.attrs.get("start", "1") or "1")
        except ValueError:
            start_idx = 1

        curr_idx = start_idx
        for child in node.children:
            if isinstance(child, HtmlNode) and child.tag == "li":
                item_md = _render_list_item(child, is_ordered=True, index=curr_idx, depth=list_depth)
                if item_md:
                    items.append(item_md)
                curr_idx += 1
        if not items:
            return ""
        return "\n\n" + "\n".join(items) + "\n\n"

    # Table
    if tag == "table":
        return _render_table(node)

    # General container or inline fallback
    return render_children(node)


def _render_list_item(node: HtmlNode, is_ordered: bool, index: int, depth: int) -> str:
    """Render a single <li> element, handling text and nested sublists."""
    indent = "  " * depth
    prefix = f"{indent}{index}. " if is_ordered else f"{indent}- "

    item_text_parts: List[str] = []
    nested_lists_md: List[str] = []

    for child in node.children:
        if isinstance(child, HtmlNode) and child.tag in ("ul", "ol"):
            nested_md = _render_node_to_markdown(child, list_depth=depth + 1)
            if nested_md.strip():
                nested_lists_md.append(nested_md.strip("\n"))
        else:
            item_text_parts.append(_render_node_to_markdown(child, list_depth=depth))

    content = "".join(item_text_parts).strip()
    result = f"{prefix}{content}"

    if nested_lists_md:
        result += "\n" + "\n".join(nested_lists_md)

    return result


def _render_table(node: HtmlNode) -> str:
    """Convert an HTML <table> node to a Markdown pipe table."""
    rows: List[List[str]] = []

    def collect_rows(curr: HtmlNode) -> None:
        for child in curr.children:
            if isinstance(child, HtmlNode):
                if child.tag == "tr":
                    cells = []
                    for cell in child.children:
                        if isinstance(cell, HtmlNode) and cell.tag in ("td", "th"):
                            cell_text = "".join(_render_node_to_markdown(c) for c in cell.children).strip()
                            cell_text = cell_text.replace("\n", " ")
                            cell_text = cell_text.replace("|", "\\|")
                            cell_text = re.sub(r"\s+", " ", cell_text)
                            cells.append(cell_text)
                    if cells:
                        rows.append(cells)
                elif child.tag in ("thead", "tbody", "tfoot"):
                    collect_rows(child)

    collect_rows(node)

    if not rows:
        return ""

    num_cols = max(len(r) for r in rows)
    if num_cols == 0:
        return ""

    padded_rows = [r + [""] * (num_cols - len(r)) for r in rows]

    header_row = padded_rows[0]
    separator_row = ["---"] * num_cols
    data_rows = padded_rows[1:]

    lines = [
        "| " + " | ".join(header_row) + " |",
        "| " + " | ".join(separator_row) + " |",
    ]

    for row in data_rows:
        lines.append("| " + " | ".join(row) + " |")

    return "\n\n" + "\n".join(lines) + "\n\n"


def _extract_raw_text(node: Union[HtmlNode, str]) -> str:
    """Extract all text contents from node, skipping script/style/head."""
    if isinstance(node, str):
        return node
    if node.tag in IGNORED_TAGS:
        return ""
    if node.tag == "img":
        return node.attrs.get("alt", "") or ""

    parts = []
    for child in node.children:
        parts.append(_extract_raw_text(child))
        if isinstance(child, HtmlNode) and child.tag in (
            "p",
            "div",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
            "tr",
            "li",
        ):
            parts.append("\n")

    return "".join(parts)


def convert_html(html_content: str) -> tuple[str, str]:
    """Convert HTML string to clean Markdown and raw text.

    Returns:
        tuple of (markdown_text, original_raw_text)
    """
    builder = HtmlTreeBuilder()
    builder.feed(html_content)
    builder.close()

    raw_text = _extract_raw_text(builder.root).strip()

    markdown = _render_node_to_markdown(builder.root)
    # Strip trailing whitespace on empty lines, but preserve double-space line breaks
    cleaned_lines = []
    for line in markdown.splitlines():
        if not line.strip():
            cleaned_lines.append("")
        elif line.endswith("  "):
            cleaned_lines.append(line.rstrip(" \t") + "  ")
        else:
            cleaned_lines.append(line.rstrip(" \t"))
    markdown = "\n".join(cleaned_lines)
    # Normalize multiple newlines
    markdown = re.sub(r"\n{3,}", "\n\n", markdown).strip()

    return markdown, raw_text


def convert_html_file(file_path: Union[str, Path]) -> tuple[str, str]:
    """Convert HTML file to clean Markdown and raw text."""
    path = Path(file_path)
    for encoding in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            content = path.read_text(encoding=encoding)
            return convert_html(content)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("html", b"", 0, 0, f"Failed to decode {file_path} with supported encodings.")
