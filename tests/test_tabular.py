"""Tests for CSV and XLSX to Markdown table conversion."""

import io
import tempfile
import zipfile
from pathlib import Path
from mdify.convert_tabular import (
    convert_csv,
    convert_csv_file,
    convert_xlsx,
    convert_xlsx_file,
)


# ==========================================
# CSV Tests
# ==========================================


def test_csv_basic():
    csv_data = """Name,Age,City
Alice,30,New York
Bob,25,San Francisco
"""
    md, raw = convert_csv(csv_data)
    assert "| Name | Age | City |" in md
    assert "| --- | --- | --- |" in md
    assert "| Alice | 30 | New York |" in md
    assert "| Bob | 25 | San Francisco |" in md
    assert "Alice" in raw


def test_csv_quoted_and_pipes():
    csv_data = '''"Product","Description","Price"
"Widget A","Contains | pipes and ""quotes""",19.99
"Widget B","Simple item",9.50
'''
    md, _ = convert_csv(csv_data)
    assert r"Contains \| pipes and" in md
    assert "| Widget A |" in md


def test_csv_empty():
    md, raw = convert_csv("")
    assert md == ""
    assert raw == ""


def test_csv_file_with_bom():
    with tempfile.TemporaryDirectory() as tmpdir:
        fpath = Path(tmpdir) / "bom.csv"
        # Write with UTF-8 BOM
        fpath.write_bytes("\ufeffItem,Qty\nApples,5\n".encode("utf-8-sig"))
        md, _ = convert_csv_file(fpath)
        assert "| Item | Qty |" in md
        assert "| Apples | 5 |" in md


# ==========================================
# XLSX Tests
# ==========================================


def _create_minimal_xlsx(
    sheets: list[tuple[str, str]],
    shared_strings: list[str] | None = None,
) -> bytes:
    """Helper to build an in-memory xlsx zip archive for testing."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        # [Content_Types].xml
        overrides = "\n".join(
            f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(len(sheets))
        )
        content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
          {overrides}
        </Types>"""
        zf.writestr("[Content_Types].xml", content_types)

        # _rels/.rels
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>""",
        )

        # xl/_rels/workbook.xml.rels
        sheet_rels = "\n".join(
            f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
            for i in range(len(sheets))
        )
        sst_rel = f'<Relationship Id="rIdSST" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        wb_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          {sheet_rels}
          {sst_rel}
        </Relationships>"""
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)

        # xl/workbook.xml
        sheet_tags = "\n".join(
            f'<sheet name="{name}" sheetId="{i+1}" r:id="rId{i+1}"/>'
            for i, (name, _) in enumerate(sheets)
        )
        workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            {sheet_tags}
          </sheets>
        </workbook>"""
        zf.writestr("xl/workbook.xml", workbook_xml)

        # xl/sharedStrings.xml
        if shared_strings:
            sst_items = "\n".join(f"<si><t>{s}</t></si>" for s in shared_strings)
            sst_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              {sst_items}
            </sst>"""
            zf.writestr("xl/sharedStrings.xml", sst_xml)

        # Worksheets
        for i, (_, sheet_content) in enumerate(sheets):
            zf.writestr(f"xl/worksheets/sheet{i+1}.xml", sheet_content)

    return buf.getvalue()


def test_xlsx_conversion_multiple_sheets():
    shared = ["Name", "Score", "Bob", "Passed"]
    sheet1_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="s"><v>0</v></c>
          <c r="B1" t="s"><v>1</v></c>
        </row>
        <row r="2">
          <c r="A2" t="s"><v>2</v></c>
          <c r="B2"><v>95</v></c>
        </row>
      </sheetData>
    </worksheet>"""

    sheet2_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="s"><v>3</v></c>
          <c r="B1" t="b"><v>1</v></c>
        </row>
      </sheetData>
    </worksheet>"""

    xlsx_bytes = _create_minimal_xlsx(
        [("Students", sheet1_xml), ("Audit", sheet2_xml)],
        shared_strings=shared,
    )

    md, raw = convert_xlsx(xlsx_bytes)
    assert "## Students" in md
    assert "| Name | Score |" in md
    assert "| Bob | 95 |" in md
    assert "## Audit" in md
    assert "| Passed | TRUE |" in md
    assert "Students" in raw
    assert "Audit" in raw


def test_xlsx_empty_sheet():
    sheet_empty = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData></sheetData>
    </worksheet>"""

    xlsx_bytes = _create_minimal_xlsx([("Blank", sheet_empty)])
    md, _ = convert_xlsx(xlsx_bytes)
    assert "## Blank" in md
    assert "(Empty sheet)" in md


def test_xlsx_file_conversion():
    with tempfile.TemporaryDirectory() as tmpdir:
        fpath = Path(tmpdir) / "test.xlsx"
        sheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1"><v>100</v></c>
              <c r="B1"><v>200</v></c>
            </row>
          </sheetData>
        </worksheet>"""
        fpath.write_bytes(_create_minimal_xlsx([("Data", sheet_xml)]))

        md, _ = convert_xlsx_file(fpath)
        assert "## Data" in md
        assert "| 100 | 200 |" in md
