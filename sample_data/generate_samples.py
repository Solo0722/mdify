"""Generate sample files (sample.html, sample.docx, sample.csv, sample.xlsx)

using only the Python standard library.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

SAMPLE_DIR = Path(__file__).resolve().parent


def generate_sample_html(output_path: Path) -> None:
    """Generate sample.html containing heading, lists, bold text, link, and table."""
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Sample Document</title>
    <style>
        body { font-family: sans-serif; }
    </style>
    <script>
        console.log("Ignore me");
    </script>
</head>
<body>
    <h1>Sample HTML Document</h1>
    <p>This is a demonstration document with <strong>bold text</strong> and an <em>italic note</em>.</p>
    <p>Here is a link to the <a href="https://example.com/project">Project Repository</a>.</p>
    <p><img src="https://example.com/logo.png" alt="Project Logo"></p>

    <h2>Lists Section</h2>
    <p>Below is a bulleted list:</p>
    <ul>
        <li>First bullet point</li>
        <li>Second bullet point
            <ul>
                <li>Nested sub-bullet item</li>
            </ul>
        </li>
        <li>Third bullet point</li>
    </ul>

    <p>Below is a numbered list:</p>
    <ol>
        <li>Initial installation step</li>
        <li>Configuration step</li>
        <li>Execution step</li>
    </ol>

    <h2>Data Table</h2>
    <table>
        <thead>
            <tr>
                <th>Component</th>
                <th>Type</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>HTML Converter</td>
                <td>Parser</td>
                <td>Active</td>
            </tr>
            <tr>
                <td>DOCX Converter</td>
                <td>OpenXML</td>
                <td>Active</td>
            </tr>
            <tr>
                <td>Tabular Converter</td>
                <td>Spreadsheet</td>
                <td>Active</td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""
    output_path.write_text(html_content, encoding="utf-8")


def generate_sample_docx(output_path: Path) -> None:
    """Generate sample.docx by constructing a valid OpenXML package by hand."""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/mdify" TargetMode="External"/>
</Relationships>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
  </w:style>
</w:styles>"""

    numbering_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>"""

    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <!-- Heading 1 -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:t>Sample DOCX Document</w:t>
      </w:r>
    </w:p>

    <!-- Paragraph with bold text and link -->
    <w:p>
      <w:r>
        <w:t xml:space="preserve">This document demonstrates conversion with </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>bold text formatting</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> and an external hyperlink to </w:t>
      </w:r>
      <w:hyperlink r:id="rId3">
        <w:r>
          <w:t>mdify project documentation</w:t>
        </w:r>
      </w:hyperlink>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>

    <!-- Heading 2 -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:t>Bulleted List Section</w:t>
      </w:r>
    </w:p>

    <!-- Bulleted list items -->
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>First bullet point item</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>Second bullet point item</w:t>
      </w:r>
    </w:p>

    <!-- Heading 2 -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:t>Numbered List Section</w:t>
      </w:r>
    </w:p>

    <!-- Numbered list items -->
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="2"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>Phase 1: Requirements Analysis</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="2"/>
        </w:numPr>
      </w:pPr>
      <w:r>
        <w:t>Phase 2: Implementation</w:t>
      </w:r>
    </w:p>

    <!-- Heading 2 for Table -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:t>Structured Table</w:t>
      </w:r>
    </w:p>

    <!-- Table -->
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r><w:t>Feature</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>Standard Library</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>Status</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r><w:t>DOCX Support</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>zipfile + xml.etree</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>Ready</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r><w:t>Pipe Tables</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>Standard string formatting</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>Ready</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/numbering.xml", numbering_xml)
        zf.writestr("word/document.xml", document_xml)


def generate_sample_csv(output_path: Path) -> None:
    """Generate sample.csv with tabular data."""
    csv_content = """ID,Product Name,Category,Price,In Stock
101,Markdown Converter,Software,49.99,Yes
102,LLM Context Optimizer,Software,89.00,Yes
103,Prompt Engineering Handbook,Book,24.50,No
104,Document Parser Toolkit,Software,129.95,Yes
"""
    output_path.write_text(csv_content, encoding="utf-8")


def generate_sample_xlsx(output_path: Path) -> None:
    """Generate sample.xlsx by constructing a multi-sheet OpenXML spreadsheet by hand."""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Overview" sheetId="1" r:id="rId1"/>
    <sheet name="Performance" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""

    # Shared strings list
    shared_strings_list = [
        "Module",           # 0
        "Format",           # 1
        "Tokens Supported", # 2
        "convert_html",     # 3
        "HTML/HTM",         # 4
        "Unlimited",        # 5
        "convert_docx",     # 6
        "DOCX",             # 7
        "convert_tabular",  # 8
        "CSV/XLSX",         # 9
        "Benchmark",        # 10
        "Speed (docs/sec)", # 11
        "Memory (MB)",      # 12
        "HTML Ingestion",   # 13
        "DOCX Extraction",  # 14
        "XLSX Parsing",     # 15
    ]

    sst_items = "\n".join(f"  <si><t>{s}</t></si>" for s in shared_strings_list)
    shared_strings_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings_list)}" uniqueCount="{len(shared_strings_list)}">
{sst_items}
</sst>"""

    # Sheet 1: Overview
    sheet1_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
      <c r="B2" t="s"><v>4</v></c>
      <c r="C2" t="s"><v>5</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>6</v></c>
      <c r="B3" t="s"><v>7</v></c>
      <c r="C3" t="s"><v>5</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>8</v></c>
      <c r="B4" t="s"><v>9</v></c>
      <c r="C4" t="s"><v>5</v></c>
    </row>
  </sheetData>
</worksheet>"""

    # Sheet 2: Performance
    sheet2_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>10</v></c>
      <c r="B1" t="s"><v>11</v></c>
      <c r="C1" t="s"><v>12</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>13</v></c>
      <c r="B2"><v>1500</v></c>
      <c r="C2"><v>12</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>14</v></c>
      <c r="B3"><v>850</v></c>
      <c r="C3"><v>18</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>15</v></c>
      <c r="B4"><v>1200</v></c>
      <c r="C4"><v>15</v></c>
    </row>
  </sheetData>
</worksheet>"""

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/sharedStrings.xml", shared_strings_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2_xml)


def generate_all_samples(target_dir: Path | None = None) -> None:
    """Generate all 4 sample files into the target directory."""
    out_dir = target_dir if target_dir is not None else SAMPLE_DIR
    out_dir.mkdir(parents=True, exist_ok=True)

    html_file = out_dir / "sample.html"
    docx_file = out_dir / "sample.docx"
    csv_file = out_dir / "sample.csv"
    xlsx_file = out_dir / "sample.xlsx"

    print(f"Generating samples in: {out_dir}")
    generate_sample_html(html_file)
    print(f"  Created: {html_file.name}")

    generate_sample_docx(docx_file)
    print(f"  Created: {docx_file.name}")

    generate_sample_csv(csv_file)
    print(f"  Created: {csv_file.name}")

    generate_sample_xlsx(xlsx_file)
    print(f"  Created: {xlsx_file.name}")

    print("Sample generation complete.")


if __name__ == "__main__":
    generate_all_samples()
