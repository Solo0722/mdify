"""Tests for DOCX to Markdown conversion."""

import io
import tempfile
import zipfile
from pathlib import Path
from mdify.convert_docx import convert_docx, convert_docx_file


def _create_minimal_docx(
    document_xml: str,
    numbering_xml: str | None = None,
    rels_xml: str | None = None,
) -> bytes:
    """Helper to build an in-memory docx zip archive for testing."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>""",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>""",
        )
        zf.writestr("word/document.xml", document_xml)
        if numbering_xml is not None:
            zf.writestr("word/numbering.xml", numbering_xml)
        if rels_xml is not None:
            zf.writestr("word/_rels/document.xml.rels", rels_xml)

    return buf.getvalue()


def test_docx_headings():
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
          <w:r><w:t>Main Document Title</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
          <w:r><w:t>Section Subtitle</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:pStyle w:val="heading 3"/></w:pPr>
          <w:r><w:t>Subsection</w:t></w:r>
        </w:p>
      </w:body>
    </w:document>"""
    data = _create_minimal_docx(doc_xml)
    md, raw = convert_docx(data)
    assert "# Main Document Title" in md
    assert "## Section Subtitle" in md
    assert "### Subsection" in md
    assert "Main Document Title" in raw


def test_docx_bold_and_italic():
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r><w:t xml:space="preserve">Normal </w:t></w:r>
          <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>bold</w:t>
          </w:r>
          <w:r><w:t xml:space="preserve"> and </w:t></w:r>
          <w:r>
            <w:rPr><w:i/></w:rPr>
            <w:t>italic</w:t>
          </w:r>
          <w:r><w:t xml:space="preserve"> and </w:t></w:r>
          <w:r>
            <w:rPr><w:b/><w:i/></w:rPr>
            <w:t>both</w:t>
          </w:r>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
      </w:body>
    </w:document>"""
    data = _create_minimal_docx(doc_xml)
    md, _ = convert_docx(data)
    assert "**bold**" in md
    assert "*italic*" in md
    assert "***both***" in md


def test_docx_lists():
    num_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
      </w:abstractNum>
      <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="10"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="20"><w:abstractNumId w:val="1"/></w:num>
    </w:numbering>"""

    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
          <w:r><w:t>Bullet point alpha</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>
          <w:r><w:t>Bullet point beta</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="20"/></w:numPr></w:pPr>
          <w:r><w:t>Ordered first</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="20"/></w:numPr></w:pPr>
          <w:r><w:t>Ordered second</w:t></w:r>
        </w:p>
      </w:body>
    </w:document>"""

    data = _create_minimal_docx(doc_xml, numbering_xml=num_xml)
    md, _ = convert_docx(data)
    assert "- Bullet point alpha" in md
    assert "- Bullet point beta" in md
    assert "1. Ordered first" in md
    assert "2. Ordered second" in md


def test_docx_hyperlinks():
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://python.org" TargetMode="External"/>
    </Relationships>"""

    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        <w:p>
          <w:r><w:t>Visit </w:t></w:r>
          <w:hyperlink r:id="rIdLink">
            <w:r><w:t>Official Python Site</w:t></w:r>
          </w:hyperlink>
        </w:p>
      </w:body>
    </w:document>"""

    data = _create_minimal_docx(doc_xml, rels_xml=rels_xml)
    md, _ = convert_docx(data)
    assert "[Official Python Site](https://python.org)" in md


def test_docx_table():
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:tbl>
          <w:tr>
            <w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>Role</w:t></w:r></w:p></w:tc>
          </w:tr>
          <w:tr>
            <w:tc><w:p><w:r><w:t>Alice</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>Engineer</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
      </w:body>
    </w:document>"""
    data = _create_minimal_docx(doc_xml)
    md, _ = convert_docx(data)
    assert "| Name | Role |" in md
    assert "| --- | --- |" in md
    assert "| Alice | Engineer |" in md


def test_docx_file_conversion():
    with tempfile.TemporaryDirectory() as tmpdir:
        fpath = Path(tmpdir) / "test.docx"
        doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>File Header</w:t></w:r></w:p>
          </w:body>
        </w:document>"""
        fpath.write_bytes(_create_minimal_docx(doc_xml))

        md, raw = convert_docx_file(fpath)
        assert "# File Header" in md
        assert "File Header" in raw
