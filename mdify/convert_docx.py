"""DOCX to Markdown converter using standard library zipfile and xml.etree."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _qn(tag: str) -> str:
    """Helper for WordprocessingML qualified names."""
    return f"{{{W_NS}}}{tag}"


@dataclass
class _Run:
    text: str
    bold: bool = False
    italic: bool = False
    link_url: Optional[str] = None


def _is_truthy(elem: Optional[ET.Element]) -> bool:
    """Check if a Word boolean element (like w:b or w:i) is truthy."""
    if elem is None:
        return False
    val = elem.attrib.get(_qn("val"), "true").lower()
    return val not in ("0", "false", "off")


def _wrap_inline(text: str, bold: bool, italic: bool, link_url: Optional[str]) -> str:
    """Wrap run text with markdown formatting, preserving whitespace outside markers."""
    if not text:
        return ""

    match = re.match(r"^(\s*)(.*?)(\s*)$", text, re.DOTALL)
    if not match:
        return text

    leading, core, trailing = match.groups()
    if not core:
        return text

    formatted = core
    if bold and italic:
        formatted = f"***{formatted}***"
    elif bold:
        formatted = f"**{formatted}**"
    elif italic:
        formatted = f"*{formatted}*"

    if link_url:
        formatted = f"[{formatted}]({link_url})"

    return f"{leading}{formatted}{trailing}"


class DocxParser:
    """Parses a WordprocessingML (.docx) document into Markdown."""

    def __init__(self, docx_zip: zipfile.ZipFile) -> None:
        self.zip = docx_zip
        self.relationships: Dict[str, str] = self._load_relationships()
        self.numbering: Dict[Tuple[str, int], str] = self._load_numbering()
        self.list_counters: Dict[Tuple[str, int], int] = {}
        self.raw_text_parts: List[str] = []

    def _load_relationships(self) -> Dict[str, str]:
        """Load relationships to resolve hyperlink URLs."""
        rels = {}
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in self.zip.namelist():
            try:
                tree = ET.fromstring(self.zip.read(rels_path))
                for rel in tree:
                    rel_id = rel.attrib.get("Id")
                    target = rel.attrib.get("Target")
                    if rel_id and target:
                        rels[rel_id] = target
            except Exception:
                pass
        return rels

    def _load_numbering(self) -> Dict[Tuple[str, int], str]:
        """Load numbering definitions from word/numbering.xml."""
        numbering_map: Dict[Tuple[str, int], str] = {}
        numbering_path = "word/numbering.xml"
        if numbering_path not in self.zip.namelist():
            return numbering_map

        try:
            tree = ET.fromstring(self.zip.read(numbering_path))
            abstract_nums: Dict[str, Dict[int, str]] = {}

            # Map abstractNumId -> {ilvl: numFmt}
            for abstract_num in tree.findall(_qn("abstractNum")):
                abs_id = abstract_num.attrib.get(_qn("abstractNumId"))
                if not abs_id:
                    continue
                abstract_nums[abs_id] = {}
                for lvl in abstract_num.findall(_qn("lvl")):
                    ilvl_str = lvl.attrib.get(_qn("ilvl"), "0")
                    num_fmt_elem = lvl.find(_qn("numFmt"))
                    num_fmt = num_fmt_elem.attrib.get(_qn("val"), "decimal") if num_fmt_elem is not None else "decimal"
                    try:
                        ilvl = int(ilvl_str)
                    except ValueError:
                        ilvl = 0
                    abstract_nums[abs_id][ilvl] = num_fmt

            # Map numId -> abstractNumId and populate numbering_map
            for num in tree.findall(_qn("num")):
                num_id = num.attrib.get(_qn("numId"))
                abs_ref = num.find(_qn("abstractNumId"))
                if not num_id or abs_ref is None:
                    continue
                abs_id = abs_ref.attrib.get(_qn("val"))
                if abs_id and abs_id in abstract_nums:
                    for ilvl, fmt in abstract_nums[abs_id].items():
                        numbering_map[(num_id, ilvl)] = fmt
        except Exception:
            pass

        return numbering_map

    def convert(self) -> Tuple[str, str]:
        """Convert the docx document to (markdown, raw_text)."""
        doc_path = "word/document.xml"
        if doc_path not in self.zip.namelist():
            return "", ""

        tree = ET.fromstring(self.zip.read(doc_path))
        body = tree.find(_qn("body"))
        if body is None:
            return "", ""

        blocks: List[Tuple[str, bool]] = []  # (block_text, is_list_item)
        for child in body:
            tag = child.tag
            if tag == _qn("p"):
                block, is_list = self._parse_paragraph(child)
                if block:
                    blocks.append((block, is_list))
            elif tag == _qn("tbl"):
                block = self._parse_table(child)
                if block:
                    blocks.append((block, False))

        if not blocks:
            return "", ""

        # Join list items tightly with single newline; other blocks with double newline
        result_parts = [blocks[0][0]]
        for i in range(1, len(blocks)):
            prev_block, prev_is_list = blocks[i - 1]
            curr_block, curr_is_list = blocks[i]
            if prev_is_list and curr_is_list:
                result_parts.append("\n" + curr_block)
            else:
                result_parts.append("\n\n" + curr_block)

        markdown = "".join(result_parts)
        markdown = re.sub(r"\n{3,}", "\n\n", markdown).strip()
        raw_text = "".join(self.raw_text_parts).strip()
        return markdown, raw_text

    def _parse_paragraph(self, p_elem: ET.Element) -> Tuple[str, bool]:
        """Parse a <w:p> element into (Markdown_string, is_list_item)."""
        p_pr = p_elem.find(_qn("pPr"))

        # 1. Check for Heading style
        heading_level = None
        is_list_item = False
        ilvl = 0
        num_id = None

        if p_pr is not None:
            p_style = p_pr.find(_qn("pStyle"))
            if p_style is not None:
                val = p_style.attrib.get(_qn("val"), "")
                match = re.search(r"(?:heading|Heading)\s*([1-6])", val)
                if match:
                    heading_level = int(match.group(1))

            if heading_level is None:
                outline_lvl = p_pr.find(_qn("outlineLvl"))
                if outline_lvl is not None:
                    try:
                        lvl = int(outline_lvl.attrib.get(_qn("val"), "0"))
                        if 0 <= lvl <= 5:
                            heading_level = lvl + 1
                    except ValueError:
                        pass

            num_pr = p_pr.find(_qn("numPr"))
            if num_pr is not None:
                is_list_item = True
                ilvl_elem = num_pr.find(_qn("ilvl"))
                num_id_elem = num_pr.find(_qn("numId"))
                if ilvl_elem is not None:
                    try:
                        ilvl = int(ilvl_elem.attrib.get(_qn("val"), "0"))
                    except ValueError:
                        ilvl = 0
                if num_id_elem is not None:
                    num_id = num_id_elem.attrib.get(_qn("val"))

        # Extract runs and hyperlinks
        runs = self._extract_runs(p_elem)
        if not runs:
            return "", False

        # Merge adjacent runs with identical formatting
        merged_runs: List[_Run] = []
        for r in runs:
            if not r.text:
                continue
            if (
                merged_runs
                and merged_runs[-1].bold == r.bold
                and merged_runs[-1].italic == r.italic
                and merged_runs[-1].link_url == r.link_url
            ):
                merged_runs[-1].text += r.text
            else:
                merged_runs.append(_Run(text=r.text, bold=r.bold, italic=r.italic, link_url=r.link_url))

        # Build inline markdown text and record raw text
        p_raw = "".join(r.text for r in merged_runs)
        self.raw_text_parts.append(p_raw + "\n")

        p_md = "".join(_wrap_inline(r.text, r.bold, r.italic, r.link_url) for r in merged_runs).strip()
        if not p_md:
            return "", False

        if heading_level:
            return f"{'#' * heading_level} {p_md}", False

        if is_list_item:
            fmt = self.numbering.get((num_id or "", ilvl), "bullet")
            indent = "  " * ilvl
            if fmt == "bullet":
                return f"{indent}- {p_md}", True
            else:
                counter_key = (num_id or "", ilvl)
                counter = self.list_counters.get(counter_key, 0) + 1
                self.list_counters[counter_key] = counter
                return f"{indent}{counter}. {p_md}", True

        return p_md, False

    def _extract_runs(self, parent: ET.Element, link_url: Optional[str] = None) -> List[_Run]:
        """Extract _Run items from paragraph or hyperlink children."""
        runs: List[_Run] = []

        for child in parent:
            tag = child.tag
            if tag == _qn("r"):
                r_pr = child.find(_qn("rPr"))
                bold = False
                italic = False
                if r_pr is not None:
                    bold = _is_truthy(r_pr.find(_qn("b")))
                    italic = _is_truthy(r_pr.find(_qn("i")))

                run_text_parts = []
                for r_child in child:
                    if r_child.tag == _qn("t"):
                        run_text_parts.append(r_child.text or "")
                    elif r_child.tag == _qn("tab"):
                        run_text_parts.append("\t")
                    elif r_child.tag == _qn("br"):
                        run_text_parts.append("\n")

                text = "".join(run_text_parts)
                if text:
                    runs.append(_Run(text=text, bold=bold, italic=italic, link_url=link_url))

            elif tag == _qn("hyperlink"):
                r_id = child.attrib.get(f"{{{R_NS}}}id")
                target_url = self.relationships.get(r_id or "", "") or None
                runs.extend(self._extract_runs(child, link_url=target_url))

        return runs

    def _parse_table(self, tbl_elem: ET.Element) -> str:
        """Parse a <w:tbl> table element into a Markdown pipe table."""
        rows: List[List[str]] = []

        for tr in tbl_elem.findall(_qn("tr")):
            row_cells = []
            for tc in tr.findall(_qn("tc")):
                cell_parts = []
                for p in tc.findall(_qn("p")):
                    p_text, _ = self._parse_paragraph(p)
                    if p_text:
                        cell_parts.append(p_text)
                cell_str = " ".join(cell_parts).strip()
                cell_str = cell_str.replace("\n", " ").replace("|", "\\|")
                cell_str = re.sub(r"\s+", " ", cell_str)
                row_cells.append(cell_str)
            if row_cells:
                rows.append(row_cells)

        if not rows:
            return ""

        num_cols = max(len(r) for r in rows)
        if num_cols == 0:
            return ""

        padded_rows = [r + [""] * (num_cols - len(r)) for r in rows]

        header = padded_rows[0]
        separator = ["---"] * num_cols
        lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(separator) + " |",
        ]

        for row in padded_rows[1:]:
            lines.append("| " + " | ".join(row) + " |")

        return "\n".join(lines)


def convert_docx(docx_bytes_or_file: Union[str, Path, zipfile.ZipFile, bytes]) -> Tuple[str, str]:
    """Convert DOCX file or bytes to Markdown and raw text.

    Returns:
        tuple of (markdown_text, original_raw_text)
    """
    if isinstance(docx_bytes_or_file, zipfile.ZipFile):
        parser = DocxParser(docx_bytes_or_file)
        return parser.convert()

    if isinstance(docx_bytes_or_file, bytes):
        import io

        with zipfile.ZipFile(io.BytesIO(docx_bytes_or_file)) as zf:
            parser = DocxParser(zf)
            return parser.convert()

    path = Path(docx_bytes_or_file)
    with zipfile.ZipFile(path) as zf:
        parser = DocxParser(zf)
        return parser.convert()


def convert_docx_file(file_path: Union[str, Path]) -> Tuple[str, str]:
    """Convert DOCX file path to Markdown and raw text."""
    return convert_docx(file_path)
