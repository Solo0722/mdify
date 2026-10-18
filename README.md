# Mdify

Mdify converts HTML, DOCX, CSV, XLSX, and PPTX files to Markdown for LLM
pipelines. It runs on Python 3.10–3.14 with **no runtime dependencies**.

## Install and run

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
mdify convert sample_data/sample.pptx
mdify convert sample_data/sample.docx -o document.md --stats
python -m pytest
```

`python -m mdify` is equivalent to `mdify`. Generate all five samples with
`python sample_data/generate_samples.py`.

## Formats

| Format | Supported content |
| --- | --- |
| HTML / HTM | Headings, paragraphs, inline formatting, links, images, blockquotes, code, nested lists, tables with row/column spans |
| DOCX | Main-body paragraphs and tables, headings and list styles, numbering starts/restarts, bold/italic, hyperlinks, content controls, merged table columns |
| CSV | Quoted and multiline fields, comma/semicolon/tab/pipe detection, explicit delimiter override |
| XLSX | Worksheets in workbook order, shared/inline strings, Boolean/error values, common date/time formats, cached formula results, uncached formula text |
| PPTX | Slides in presentation order, titles, text boxes, groups, inherited bullets, numbered lists, bold/italic, links, tables, speaker notes, image descriptions |

Legacy `.doc`, `.ppt`, and `.xls` files, password-encrypted documents, PDFs,
and OCR are not supported. Save legacy Office files in their modern formats
before conversion.

## Command-line options

```text
mdify convert FILE [-t FORMAT] [-o OUTPUT] [--stats]
                   [--encoding ENCODING] [--delimiter CHARACTER] [--no-notes]
```

- `FILE`: Input path, or `-` for stdin. Extensions are case-insensitive.
- `-t`, `--type`: Override the input format; required for stdin. Accepts `html`,
  `htm`, `docx`, `csv`, `xlsx`, or `pptx`, with or without the leading dot.
- `-o`, `--output`: Write UTF-8 Markdown to a file. Missing parent directories
  are created. Input/output collisions, including aliases to the same file,
  are rejected. Output replacement is atomic so a failed write preserves the
  previous destination.
- `--stats`: Compare extracted-text and Markdown character counts. Estimated
  tokens are `characters / 4`, **not tokenizer measurements**. The input count
  measures extracted text, not file size or the original HTML/XML source.
  Statistics go to stderr when Markdown goes to stdout, otherwise to stdout.
- `--encoding`: Explicit byte decoding for HTML/CSV. Automatic decoding checks
  Unicode BOMs, an HTML charset declaration, UTF-8, Windows-1252, then Latin-1.
  File and byte-stdin inputs use the same decoding. Specify this option when
  the encoding cannot be inferred reliably.
- `--delimiter`: CSV-only override. Use `";"`, `"^"`, or `"\t"`, for example.
  Automatic detection is limited to comma, semicolon, tab, and pipe; ambiguous
  input falls back to comma.
- `--no-notes`: PPTX-only option to exclude speaker notes from both Markdown
  and extracted-text statistics.

```powershell
mdify convert slides.pptx --no-notes -o slides.md
mdify convert data.csv --delimiter ";" --encoding cp1252
mdify convert export.bin --type html -o export.md
```

For text stdin:

```sh
cat sample_data/sample.html | mdify convert - --type html
```

Office formats require **binary stdin**. Shells that decode or re-encode a
pipeline can corrupt a ZIP-based Office file; a direct file path avoids that.
A portable Python example preserves the bytes:

```python
from pathlib import Path
import subprocess
import sys

subprocess.run(
    [sys.executable, "-m", "mdify", "convert", "-", "--type", "pptx", "-o", "slides.md"],
    input=Path("sample_data/sample.pptx").read_bytes(),
    check=True,
)
```

## PPTX output and boundaries

Each slide has a numbered heading and stable anchor. The title comes first;
remaining shapes follow stored shape order, recursively including groups.
This is deterministic, but complex visual layouts may require a different
human reading order. Empty slides are retained and hidden slides are labeled.

Speaker notes appear under `### Speaker notes` by default. Notes page numbers,
headers, footers, and slide-image placeholders are excluded. Body list styles
can inherit from slide layouts and masters. Layout/master placeholder prompt
text and decorative background content are not copied into the output.

Images appear as `[Image: description]` when alternative text is available,
otherwise `[Image]`. Charts, SmartArt, embedded objects, and media use explicit
omission markers. The converter does not export images, interpret diagrams,
extract chart data, render animations, or run OCR.

## Conversion conventions and limitations

- Literal Markdown punctuation is escaped. Links encode spaces and destination
  delimiters. Code spans and fences adapt to embedded backticks; HTML code
  block whitespace is preserved.
- Tables use the first row as a Markdown header. Merged continuation cells are
  blank; content stays in the originating cell. Multiline cells use `<br>`.
  Table rendering is limited to 1,000,000 output cells, including padding.
- DOCX resolves the document, styles, and numbering through package
  relationships. Inserted text is included and deleted text is excluded.
  Headers, footers, footnotes, comments, equations, and drawing/text-box
  contents are outside the current DOCX extraction scope.
- XLSX omits completely empty rows and preserves column positions. Sparse
  worksheets do not allocate all rows from A1 to the last populated cell.
  Common date/time formats are normalized to ISO-style dates and 24-hour times;
  both Excel date systems and elapsed-hour formats are supported. Other custom
  numeric formats retain stored values; this is not an Excel display renderer.
- Formulas are never calculated. Cached results can be stale. Without a cached
  result, the formula is emitted as `=expression`; a shared-formula follower
  without its own expression uses `[Formula result unavailable]`.
- Alphabetic/Roman list numbering is normalized to decimal Markdown numbering.
  Perfect visual reproduction is not a conversion goal.
- Malformed required Office parts, missing parts needed for extraction, and invalid XML produce
  explicit conversion errors instead of silently successful empty output.
  Valid empty documents remain valid conversions.

## Python API

All converters return `(markdown, extracted_text)`. Existing positional APIs
are retained; new format-specific options are keyword-only.

```python
from mdify import convert_pptx_file, convert_csv

markdown, extracted_text = convert_pptx_file("slides.pptx", include_notes=False)
markdown, extracted_text = convert_csv("Name;Count\nExample;3", delimiter=";")
```

Office converters also accept bytes or a caller-owned `zipfile.ZipFile`.
A caller-owned archive remains open after conversion. HTML/CSV file helpers
accept an optional `encoding` keyword.

## Validation

Tests cover regressions, malformed packages, file/stdin equivalence, atomic
output, and synthetic fixtures written by Microsoft Word, Excel, and PowerPoint.
Committed Office fixtures require no Office installation to test. Regeneration
instructions are in `tests/fixtures/README.md`. CI runs on Windows and Linux
with Python 3.10 through 3.14.

## Reference project

Modeled on Microsoft's [MarkItDown](https://github.com/microsoft/markitdown).
Mdify is an independent implementation. No code was copied from the reference
project, and MarkItDown is not a dependency.

## AI tools used

- Gemini Flash 3.8
- OpenAI Codex
