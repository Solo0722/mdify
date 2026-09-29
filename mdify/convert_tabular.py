"""CSV and XLSX to Markdown table converters."""

from __future__ import annotations

import csv
import io
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _qn(tag: str) -> str:
    """Helper for SpreadsheetML qualified names."""
    return f"{{{SS_NS}}}{tag}"


def _col_letters_to_index(col_letters: str) -> int:
    """Convert column letters (e.g. 'A', 'Z', 'AA') to 0-based index."""
    idx = 0
    for char in col_letters.upper():
        if "A" <= char <= "Z":
            idx = idx * 26 + (ord(char) - ord("A") + 1)
    return max(0, idx - 1)


def _format_pipe_table(rows: List[List[str]]) -> str:
    """Format a 2D list of strings into a Markdown pipe table."""
    if not rows:
        return ""

    num_cols = max(len(r) for r in rows)
    if num_cols == 0:
        return ""

    # Pad any jagged rows
    padded = [r + [""] * (num_cols - len(r)) for r in rows]

    # Clean up cells
    cleaned_rows = []
    for r in padded:
        cleaned_cells = []
        for cell in r:
            c_text = str(cell).replace("\n", " ").replace("|", "\\|")
            c_text = re.sub(r"\s+", " ", c_text).strip()
            cleaned_cells.append(c_text)
        cleaned_rows.append(cleaned_cells)

    header = cleaned_rows[0]
    separator = ["---"] * num_cols
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]

    for row in cleaned_rows[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


# ==========================================
# CSV Converter
# ==========================================


def convert_csv(csv_content: str) -> Tuple[str, str]:
    """Convert CSV string to Markdown table and raw text."""
    csv_content = csv_content.lstrip("\ufeff")
    if not csv_content.strip():
        return "", ""

    # Try sniffing dialect, fallback to standard comma delimiter
    sample = csv_content[:4096]
    dialect = csv.excel  # default standard comma
    try:
        sniffed = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        if sniffed.delimiter and sniffed.delimiter not in ("\r", "\n"):
            dialect = sniffed
    except Exception:
        try:
            sniffed = csv.Sniffer().sniff(sample)
            if sniffed.delimiter and sniffed.delimiter not in ("\r", "\n"):
                dialect = sniffed
        except Exception:
            dialect = csv.excel

    reader = csv.reader(io.StringIO(csv_content), dialect=dialect)
    rows: List[List[str]] = [row for row in reader if any(cell.strip() for cell in row)]

    if not rows:
        return "", ""

    markdown = _format_pipe_table(rows)
    raw_text = "\n".join(" ".join(cell.strip() for cell in row) for row in rows)

    return markdown, raw_text


def convert_csv_file(file_path: Union[str, Path]) -> Tuple[str, str]:
    """Convert CSV file to Markdown table and raw text."""
    path = Path(file_path)
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            content = path.read_text(encoding=encoding)
            return convert_csv(content)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("csv", b"", 0, 0, f"Failed to decode {file_path} with supported encodings.")


# ==========================================
# XLSX Converter
# ==========================================


class XlsxParser:
    """Parses an OpenXML spreadsheet (.xlsx) into Markdown."""

    def __init__(self, xlsx_zip: zipfile.ZipFile) -> None:
        self.zip = xlsx_zip
        self.shared_strings: List[str] = self._load_shared_strings()
        self.sheets: List[Tuple[str, str]] = self._load_sheet_list()

    def _load_shared_strings(self) -> List[str]:
        """Load xl/sharedStrings.xml if present."""
        strings: List[str] = []
        sst_path = "xl/sharedStrings.xml"
        if sst_path not in self.zip.namelist():
            return strings

        try:
            tree = ET.fromstring(self.zip.read(sst_path))
            for si in tree.findall(_qn("si")):
                text_parts = [t.text or "" for t in si.iter(_qn("t"))]
                strings.append("".join(text_parts))
        except Exception:
            pass

        return strings

    def _load_sheet_list(self) -> List[Tuple[str, str]]:
        """Load sheet names and worksheet paths from xl/workbook.xml & rels."""
        sheets: List[Tuple[str, str]] = []
        wb_path = "xl/workbook.xml"
        rels_path = "xl/_rels/workbook.xml.rels"

        if wb_path not in self.zip.namelist():
            return sheets

        # Map relationship ID to target file
        rel_map: Dict[str, str] = {}
        if rels_path in self.zip.namelist():
            try:
                rels_tree = ET.fromstring(self.zip.read(rels_path))
                for rel in rels_tree:
                    r_id = rel.attrib.get("Id")
                    target = rel.attrib.get("Target")
                    if r_id and target:
                        # Normalize target relative to xl/
                        if not target.startswith("xl/"):
                            if target.startswith("/"):
                                target = target.lstrip("/")
                            else:
                                target = f"xl/{target}"
                        rel_map[r_id] = target
            except Exception:
                pass

        try:
            wb_tree = ET.fromstring(self.zip.read(wb_path))
            sheets_elem = wb_tree.find(_qn("sheets"))
            if sheets_elem is not None:
                for sheet in sheets_elem.findall(_qn("sheet")):
                    name = sheet.attrib.get("name", "Sheet")
                    r_id = sheet.attrib.get(f"{{{R_NS}}}id")
                    target_path = rel_map.get(r_id or "")
                    if not target_path:
                        # Fallback heuristic: guess based on sheetId
                        sheet_id = sheet.attrib.get("sheetId", "1")
                        target_path = f"xl/worksheets/sheet{sheet_id}.xml"
                    sheets.append((name, target_path))
        except Exception:
            pass

        return sheets

    def convert(self) -> Tuple[str, str]:
        """Convert all worksheets into Markdown sections with pipe tables."""
        sheet_sections: List[str] = []
        raw_parts: List[str] = []

        for name, sheet_path in self.sheets:
            if sheet_path not in self.zip.namelist():
                continue

            raw_parts.append(name)
            rows = self._parse_worksheet(sheet_path)
            sheet_md = f"## {name}\n\n"

            if not rows:
                sheet_md += "(Empty sheet)"
            else:
                table_md = _format_pipe_table(rows)
                sheet_md += table_md
                for r in rows:
                    raw_parts.append(" ".join(c for c in r if c))

            sheet_sections.append(sheet_md)

        markdown = "\n\n".join(sheet_sections).strip()
        raw_text = "\n".join(raw_parts).strip()
        return markdown, raw_text

    def _parse_worksheet(self, sheet_path: str) -> List[List[str]]:
        """Parse worksheet XML into a 2D grid of cell text."""
        tree = ET.fromstring(self.zip.read(sheet_path))
        sheet_data = tree.find(_qn("sheetData"))
        if sheet_data is None:
            return []

        # Read cells indexed by (row_idx, col_idx)
        cell_dict: Dict[Tuple[int, int], str] = {}
        max_row = 0
        max_col = 0

        curr_row_idx = 0
        for row_elem in sheet_data.findall(_qn("row")):
            row_attr = row_elem.attrib.get("r")
            if row_attr:
                try:
                    curr_row_idx = int(row_attr) - 1
                except ValueError:
                    pass

            curr_col_idx = 0
            for c_elem in row_elem.findall(_qn("c")):
                r_coord = c_elem.attrib.get("r")
                if r_coord:
                    m = re.match(r"^([A-Za-z]+)(\d+)$", r_coord)
                    if m:
                        curr_col_idx = _col_letters_to_index(m.group(1))
                        try:
                            curr_row_idx = int(m.group(2)) - 1
                        except ValueError:
                            pass

                val = self._extract_cell_value(c_elem)
                if val:
                    cell_dict[(curr_row_idx, curr_col_idx)] = val
                    if curr_row_idx > max_row:
                        max_row = curr_row_idx
                    if curr_col_idx > max_col:
                        max_col = curr_col_idx

                curr_col_idx += 1
            curr_row_idx += 1

        if not cell_dict:
            return []

        # Construct 2D grid
        grid: List[List[str]] = [["" for _ in range(max_col + 1)] for _ in range(max_row + 1)]
        for (r, c), val in cell_dict.items():
            grid[r][c] = val

        # Filter out empty trailing or completely blank rows
        filtered_rows = [r for r in grid if any(cell.strip() for cell in r)]
        return filtered_rows

    def _extract_cell_value(self, c_elem: ET.Element) -> str:
        """Extract formatted string value from a cell element."""
        t_type = c_elem.attrib.get("t")
        v_elem = c_elem.find(_qn("v"))
        raw_v = v_elem.text if v_elem is not None and v_elem.text is not None else ""

        if t_type == "s":
            # Shared string index
            try:
                idx = int(raw_v)
                if 0 <= idx < len(self.shared_strings):
                    return self.shared_strings[idx]
            except ValueError:
                pass
            return ""

        if t_type == "inlineStr":
            is_elem = c_elem.find(_qn("is"))
            if is_elem is not None:
                return "".join(t.text or "" for t in is_elem.iter(_qn("t")))
            return ""

        if t_type == "b":
            return "TRUE" if raw_v == "1" else "FALSE"

        return raw_v


def convert_xlsx(xlsx_bytes_or_file: Union[str, Path, zipfile.ZipFile, bytes]) -> Tuple[str, str]:
    """Convert XLSX file or bytes to Markdown and raw text.

    Returns:
        tuple of (markdown_text, original_raw_text)
    """
    if isinstance(xlsx_bytes_or_file, zipfile.ZipFile):
        parser = XlsxParser(xlsx_bytes_or_file)
        return parser.convert()

    if isinstance(xlsx_bytes_or_file, bytes):
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes_or_file)) as zf:
            parser = XlsxParser(zf)
            return parser.convert()

    path = Path(xlsx_bytes_or_file)
    with zipfile.ZipFile(path) as zf:
        parser = XlsxParser(zf)
        return parser.convert()


def convert_xlsx_file(file_path: Union[str, Path]) -> Tuple[str, str]:
    """Convert XLSX file path to Markdown and raw text."""
    return convert_xlsx(file_path)
