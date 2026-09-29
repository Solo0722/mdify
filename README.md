# Mdify

Mdify is a command-line tool that converts HTML, DOCX, CSV and XLSX files
to Markdown, a format that is compact and easy for LLMs to read.

## Reference project
Modeled on Microsoft's MarkItDown, a Python utility that converts files to
Markdown for LLM pipelines:
https://github.com/microsoft/markitdown

Mdify is an independent implementation of a subset of MarkItDown's
conversions (HTML, DOCX, CSV, XLSX). No code was copied from the reference
project, and MarkItDown is not a dependency.

## AI tools used
- Gemini Flash 3.8

## Features
1. **HTML to Markdown:** headings, lists, links, emphasis, code and tables.
2. **DOCX to Markdown:** headings, lists, bold and italic text, and tables.
3. **CSV/XLSX to Markdown tables:** one section per worksheet.
4. **Stdin & Stdout:** read from stdin with `-` and write to stdout by default.
5. **Stats:** `--stats` prints character counts and an estimated token count
   (characters divided by 4) for the input and output.

## Install and run
```bash
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e . pytest
python sample_data/generate_samples.py
mdify convert sample_data/sample.html
mdify convert sample_data/sample.docx --stats
cat sample_data/sample.html | mdify convert - --type html
pytest
```

## Supported formats and options
1. html
2. docx
3. csv
4. xlsx

### Options
- `FILE`: Input file path, or `-` to read from standard input (`stdin`).
- `-t`, `--type FORMAT`: Input format (`html`, `docx`, `csv`, `xlsx`). Required when reading from stdin (`-`).
- `-o`, `--output OUTPUT`: Path to write the output Markdown file (default: stdout).
- `--stats`: Print character count and estimated token count (`characters / 4`) for original and Markdown output. When writing Markdown to stdout, statistics are printed to stderr so output redirection remains clean.