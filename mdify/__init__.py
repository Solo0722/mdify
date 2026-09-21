"""mdify: Convert documents (HTML, DOCX, CSV, XLSX) to Markdown for LLMs."""

__version__ = "0.1.0"

from mdify.convert_html import convert_html, convert_html_file
from mdify.convert_docx import convert_docx, convert_docx_file
from mdify.convert_tabular import (
    convert_csv,
    convert_csv_file,
    convert_xlsx,
    convert_xlsx_file,
)
from mdify.stats import calculate_stats, format_stats

__all__ = [
    "convert_html",
    "convert_html_file",
    "convert_docx",
    "convert_docx_file",
    "convert_csv",
    "convert_csv_file",
    "convert_xlsx",
    "convert_xlsx_file",
    "calculate_stats",
    "format_stats",
]
