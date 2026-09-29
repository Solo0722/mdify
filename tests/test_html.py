"""Tests for HTML to Markdown conversion."""

import tempfile
from pathlib import Path
from mdify.convert_html import convert_html, convert_html_file


def test_headings():
    html = """
    <h1>Title H1</h1>
    <h2>Subtitle H2</h2>
    <h3>Section H3</h3>
    <h4>Subsection H4</h4>
    <h5>Minor H5</h5>
    <h6>Detail H6</h6>
    """
    md, raw = convert_html(html)
    assert "# Title H1" in md
    assert "## Subtitle H2" in md
    assert "### Section H3" in md
    assert "#### Subsection H4" in md
    assert "##### Minor H5" in md
    assert "###### Detail H6" in md
    assert "Title H1" in raw
    assert "Detail H6" in raw


def test_paragraphs_and_breaks():
    html = "<p>First paragraph.</p><p>Second paragraph with<br>a line break.</p><hr><p>After hr.</p>"
    md, raw = convert_html(html)
    assert "First paragraph." in md
    assert "Second paragraph with  \na line break." in md
    assert "---" in md
    assert "After hr." in md


def test_bold_and_italic():
    html = """
    <p>This has <b>bold text</b>, <strong>strong text</strong>,
    <i>italic text</i>, and <em>emphasized text</em>.</p>
    <p>And <b><i>bold italic</i></b> together.</p>
    """
    md, _ = convert_html(html)
    assert "**bold text**" in md
    assert "**strong text**" in md
    assert "*italic text*" in md
    assert "*emphasized text*" in md
    assert "***bold italic***" in md or ("**" in md and "*" in md)


def test_inline_code_and_code_blocks():
    html = """
    <p>Run the <code>pip install mdify</code> command.</p>
    <pre><code class="language-python">
def hello():
    print("world")
    </code></pre>
    """
    md, _ = convert_html(html)
    assert "`pip install mdify`" in md
    assert "```python" in md
    assert 'print("world")' in md
    assert "```" in md


def test_links():
    html = '<p>Visit <a href="https://example.com">Example Site</a> or <a href="https://github.com">GitHub</a>.</p>'
    md, _ = convert_html(html)
    assert "[Example Site](https://example.com)" in md
    assert "[GitHub](https://github.com)" in md


def test_image_with_alt():
    html = '<p><img src="https://example.com/photo.jpg" alt="A beautiful sunset"></p>'
    md, raw = convert_html(html)
    assert "![A beautiful sunset](https://example.com/photo.jpg)" in md
    assert "A beautiful sunset" in raw


def test_image_without_alt():
    html = '<p><img src="https://example.com/banner.png"></p>'
    md, _ = convert_html(html)
    assert "![](https://example.com/banner.png)" in md



def test_unordered_lists():
    html = """
    <ul>
        <li>Apple</li>
        <li>Banana</li>
        <li>Cherry</li>
    </ul>
    """
    md, _ = convert_html(html)
    assert "- Apple" in md
    assert "- Banana" in md
    assert "- Cherry" in md


def test_ordered_lists():
    html = """
    <ol>
        <li>First item</li>
        <li>Second item</li>
        <li>Third item</li>
    </ol>
    """
    md, _ = convert_html(html)
    assert "1. First item" in md
    assert "2. Second item" in md
    assert "3. Third item" in md


def test_nested_lists():
    html = """
    <ul>
        <li>Parent Item 1
            <ul>
                <li>Child Item 1.1</li>
                <li>Child Item 1.2</li>
            </ul>
        </li>
        <li>Parent Item 2
            <ol>
                <li>Step 1</li>
                <li>Step 2</li>
            </ol>
        </li>
    </ul>
    """
    md, _ = convert_html(html)
    assert "- Parent Item 1" in md
    assert "  - Child Item 1.1" in md
    assert "  - Child Item 1.2" in md
    assert "- Parent Item 2" in md
    assert "  1. Step 1" in md
    assert "  2. Step 2" in md


def test_table_conversion():
    html = """
    <table>
        <thead>
            <tr>
                <th>Header 1</th>
                <th>Header 2</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Row 1 Col 1</td>
                <td>Row 1 Col 2</td>
            </tr>
            <tr>
                <td>Row 2 Col 1</td>
                <td>Cell with | pipe</td>
            </tr>
        </tbody>
    </table>
    """
    md, _ = convert_html(html)
    assert "| Header 1 | Header 2 |" in md
    assert "| --- | --- |" in md
    assert "| Row 1 Col 1 | Row 1 Col 2 |" in md
    assert r"Cell with \| pipe" in md


def test_ignore_script_and_style():
    html = """
    <head>
        <title>Secret Title</title>
        <style>body { color: red; }</style>
    </head>
    <body>
        <h1>Visible Header</h1>
        <script>
            alert("Malicious or noisy code");
            var x = 10;
        </script>
        <noscript>No script message</noscript>
        <p>Visible Paragraph</p>
    </body>
    """
    md, raw = convert_html(html)
    assert "alert" not in md
    assert "color: red" not in md
    assert "Secret Title" not in md
    assert "# Visible Header" in md
    assert "Visible Paragraph" in md
    assert "alert" not in raw
    assert "color: red" not in raw


def test_html_entities():
    html = "<p>Tom &amp; Jerry &lt;friends&gt; &quot;quote&#39;s&quot;</p>"
    md, _ = convert_html(html)
    assert "Tom & Jerry <friends> \"quote's\"" in md


def test_convert_html_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        fpath = Path(tmpdir) / "test.html"
        fpath.write_text("<h1>File Test</h1><p>Content</p>", encoding="utf-8")
        md, raw = convert_html_file(fpath)
        assert "# File Test" in md
        assert "Content" in md
        assert "File Test" in raw
