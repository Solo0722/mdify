"""Command-line interface with one conversion path for files and stdin."""
from __future__ import annotations

import argparse
import io
import os
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Sequence

from mdify.convert_docx import convert_docx
from mdify.convert_html import convert_html
from mdify.convert_pptx import convert_pptx
from mdify.convert_tabular import convert_csv, convert_xlsx
from mdify._text import decode_text
from mdify.stats import calculate_stats, format_stats


@dataclass(frozen=True)
class Format:
    convert: Callable
    text: bool = False


SUPPORTED_EXTENSIONS = {
    '.html': Format(convert_html, text=True),
    '.htm': Format(convert_html, text=True),
    '.docx': Format(convert_docx),
    '.csv': Format(convert_csv, text=True),
    '.xlsx': Format(convert_xlsx),
    '.pptx': Format(convert_pptx),
}


def build_parser():
    formats = ', '.join(ext.lstrip('.') for ext in SUPPORTED_EXTENSIONS)
    parser = argparse.ArgumentParser(prog='mdify', description='Convert HTML, DOCX, CSV, XLSX, and PPTX to Markdown.')
    sub = parser.add_subparsers(dest='command', required=True)
    convert = sub.add_parser('convert', help='Convert a document to Markdown')
    convert.add_argument('file', help=f'Input file ({formats}), or - for stdin')
    convert.add_argument('-t', '--type', metavar='FORMAT', help=f'Input format ({formats}); required for stdin')
    convert.add_argument('-o', '--output', help='Output Markdown path (default: stdout)')
    convert.add_argument('--stats', action='store_true', help='Compare extracted text and Markdown characters; tokens estimated as characters / 4')
    convert.add_argument('--encoding', help='Explicit input encoding for HTML/CSV (otherwise detect BOM, HTML charset, UTF-8, then Windows-1252/Latin-1)')
    convert.add_argument('--delimiter', help=r'Explicit CSV delimiter; use "\t" for a tab')
    convert.add_argument('--no-notes', action='store_true', help='Exclude PPTX speaker notes')
    return parser


def _write_atomic(path, markdown):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', newline='\n',
                                         dir=path.parent, prefix='.mdify-', suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(markdown)
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def run_convert(file_path_str: str, output_path_str: str | None = None,
                show_stats: bool = False, file_type: str | None = None,
                stdin_stream=None, stdout_stream=None, stderr_stream=None,
                *, encoding: str | None = None, delimiter: str | None = None,
                include_notes: bool = True) -> int:
    stdout = stdout_stream if stdout_stream is not None else sys.stdout
    stderr = stderr_stream if stderr_stream is not None else sys.stderr
    supported = ', '.join(SUPPORTED_EXTENSIONS)
    input_path = None
    if file_path_str == '-':
        if not file_type:
            stderr.write(f"Error: --type FORMAT is required when reading from stdin ('-'). Supported formats: {supported}\n")
            return 1
    else:
        input_path = Path(file_path_str)
        if not input_path.exists():
            stderr.write(f"Error: File not found: '{file_path_str}'\n")
            return 1
        if not input_path.is_file():
            stderr.write(f"Error: Path is not a file: '{file_path_str}'\n")
            return 1
    ext = '.' + file_type.lower().lstrip('.') if file_type else input_path.suffix.lower()
    spec = SUPPORTED_EXTENSIONS.get(ext)
    if spec is None:
        problem = f"Unsupported format '{file_type}'" if input_path is None else f"Unsupported file extension '{ext}'"
        stderr.write(f'Error: {problem}. Supported formats: {supported}\n')
        return 1
    if encoding is not None and not spec.text:
        stderr.write('Error: --encoding is only supported for HTML and CSV\n')
        return 1
    if delimiter is not None and ext != '.csv':
        stderr.write('Error: --delimiter is only supported for CSV\n')
        return 1
    if not include_notes and ext != '.pptx':
        stderr.write('Error: --no-notes is only supported for PPTX\n')
        return 1
    if input_path is not None and output_path_str:
        output = Path(output_path_str)
        try:
            if input_path.resolve() == output.resolve() or (output.exists() and input_path.samefile(output)):
                stderr.write('Error: Input and output must not be the same file\n')
                return 1
        except OSError as exc:
            stderr.write(f'Error checking output path: {exc}\n')
            return 1
    try:
        if input_path is None:
            stream = stdin_stream if stdin_stream is not None else getattr(sys.stdin, 'buffer', sys.stdin)
            source = stream.read()
            if not spec.text and not isinstance(source, bytes):
                raise ValueError('Office formats require binary stdin; use a binary pipe or a file path')
        else:
            source = input_path.read_bytes() if spec.text else input_path
        options = {}
        if spec.text:
            source = decode_text(source, encoding, html=ext in {'.html', '.htm'})
        if ext == '.csv':
            options['delimiter'] = '\t' if delimiter == r'\t' else delimiter
        if ext == '.pptx':
            options['include_notes'] = include_notes
        markdown, raw = spec.convert(source, **options)
    except Exception as exc:
        label = 'stdin' if input_path is None else repr(file_path_str)
        stderr.write(f'Error converting {label}: {exc}\n')
        return 1
    try:
        if output_path_str:
            _write_atomic(Path(output_path_str), markdown)
            if show_stats:
                stdout.write(format_stats(calculate_stats(raw, markdown)) + '\n')
        else:
            stdout.write(markdown + '\n')
            if show_stats:
                stderr.write(format_stats(calculate_stats(raw, markdown)) + '\n')
    except BrokenPipeError:
        return 1
    except (OSError, UnicodeError) as exc:
        stderr.write(f'Error writing output: {exc}\n')
        return 1
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    # Redirected output should be portable UTF-8, including on Windows locales.
    for stream in (sys.stdout, sys.stderr):
        if isinstance(stream, io.TextIOWrapper) and not stream.isatty():
            stream.reconfigure(encoding='utf-8')
    args = build_parser().parse_args(argv)
    return run_convert(args.file, args.output, args.stats, args.type,
                       encoding=args.encoding, delimiter=args.delimiter,
                       include_notes=not args.no_notes)


if __name__ == '__main__':
    sys.exit(main())
