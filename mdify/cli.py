"""Command-line interface for mdify."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Sequence

from mdify.convert_docx import convert_docx_file
from mdify.convert_html import convert_html_file
from mdify.convert_tabular import convert_csv_file, convert_xlsx_file
from mdify.stats import calculate_stats, format_stats

SUPPORTED_EXTENSIONS = {
    ".html": convert_html_file,
    ".htm": convert_html_file,
    ".docx": convert_docx_file,
    ".csv": convert_csv_file,
    ".xlsx": convert_xlsx_file,
}


def build_parser() -> argparse.ArgumentParser:
    """Build the command-line argument parser."""
    parser = argparse.ArgumentParser(
        prog="mdify",
        description="Convert documents (HTML, DOCX, CSV, XLSX) to clean Markdown for LLMs.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    convert_parser = subparsers.add_parser("convert", help="Convert a document file to Markdown")
    convert_parser.add_argument(
        "file",
        help="Path to the input document (.html, .htm, .docx, .csv, .xlsx)",
    )
    convert_parser.add_argument(
        "-o",
        "--output",
        dest="output",
        help="Path to output Markdown file (default: write to stdout)",
    )
    convert_parser.add_argument(
        "--stats",
        action="store_true",
        help="Print character count and estimated token count (characters / 4) for original text and Markdown",
    )

    return parser


def run_convert(
    file_path_str: str,
    output_path_str: Optional[str] = None,
    show_stats: bool = False,
    stdout_stream=None,
    stderr_stream=None,
) -> int:
    """Execute document conversion.

    Returns exit code (0 on success, non-zero on error).
    """
    stdout = stdout_stream if stdout_stream is not None else sys.stdout
    stderr = stderr_stream if stderr_stream is not None else sys.stderr

    input_path = Path(file_path_str)
    if not input_path.exists():
        stderr.write(f"Error: File not found: '{file_path_str}'\n")
        return 1

    if not input_path.is_file():
        stderr.write(f"Error: Path is not a file: '{file_path_str}'\n")
        return 1

    ext = input_path.suffix.lower()
    converter = SUPPORTED_EXTENSIONS.get(ext)
    if converter is None:
        supported_list = ", ".join(sorted(SUPPORTED_EXTENSIONS.keys()))
        stderr.write(
            f"Error: Unsupported file extension '{ext}'. Supported formats: {supported_list}\n"
        )
        return 1

    try:
        markdown, raw_text = converter(input_path)
    except Exception as exc:
        stderr.write(f"Error converting '{file_path_str}': {exc}\n")
        return 1

    stats = calculate_stats(raw_text, markdown)
    stats_formatted = format_stats(stats)

    if output_path_str:
        out_path = Path(output_path_str)
        try:
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(markdown, encoding="utf-8")
        except Exception as exc:
            stderr.write(f"Error writing to output file '{output_path_str}': {exc}\n")
            return 1

        if show_stats:
            stdout.write(stats_formatted + "\n")
    else:
        stdout.write(markdown + "\n")
        if show_stats:
            stderr.write(stats_formatted + "\n")

    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Entry point for CLI execution."""
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "convert":
        return run_convert(
            file_path_str=args.file,
            output_path_str=args.output,
            show_stats=args.stats,
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
