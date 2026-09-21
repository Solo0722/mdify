"""Tests for mdify command-line interface."""

import io
import tempfile
from pathlib import Path
import pytest
from mdify.cli import main, run_convert
from mdify.stats import calculate_stats, format_stats

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


def test_stats_calculation():
    orig = "Hello World"  # 11 characters -> 2.75 tokens
    md = "# Hello World"  # 13 characters -> 3.25 tokens
    stats = calculate_stats(orig, md)
    assert stats.original.characters == 11
    assert stats.original.estimated_tokens == 2.75
    assert stats.original.estimated_tokens_int == 3
    assert stats.markdown.characters == 13
    assert stats.markdown.estimated_tokens == 3.25
    assert stats.markdown.estimated_tokens_int == 3

    formatted = format_stats(stats)
    assert "Original text:" in formatted
    assert "Characters: 11" in formatted
    assert "Estimated tokens: 2.75 (estimate: characters / 4)" in formatted
    assert "Markdown output:" in formatted
    assert "Characters: 13" in formatted


def test_cli_convert_stdout():
    html_file = SAMPLE_DIR / "sample.html"
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = run_convert(
        file_path_str=str(html_file),
        stdout_stream=stdout,
        stderr_stream=stderr,
    )
    assert exit_code == 0
    assert "# Sample HTML Document" in stdout.getvalue()
    assert stderr.getvalue() == ""


def test_cli_convert_output_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        out_file = Path(tmpdir) / "output.md"
        html_file = SAMPLE_DIR / "sample.html"
        stdout = io.StringIO()
        stderr = io.StringIO()

        exit_code = run_convert(
            file_path_str=str(html_file),
            output_path_str=str(out_file),
            stdout_stream=stdout,
            stderr_stream=stderr,
        )
        assert exit_code == 0
        assert out_file.exists()
        content = out_file.read_text(encoding="utf-8")
        assert "# Sample HTML Document" in content


def test_cli_convert_with_stats_to_stdout():
    html_file = SAMPLE_DIR / "sample.html"
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = run_convert(
        file_path_str=str(html_file),
        show_stats=True,
        stdout_stream=stdout,
        stderr_stream=stderr,
    )
    assert exit_code == 0
    assert "# Sample HTML Document" in stdout.getvalue()
    # Stats should be in stderr when markdown is in stdout
    assert "Statistics:" in stderr.getvalue()
    assert "Characters:" in stderr.getvalue()
    assert "Estimated tokens:" in stderr.getvalue()


def test_cli_convert_with_stats_to_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        out_file = Path(tmpdir) / "output.md"
        html_file = SAMPLE_DIR / "sample.html"
        stdout = io.StringIO()
        stderr = io.StringIO()

        exit_code = run_convert(
            file_path_str=str(html_file),
            output_path_str=str(out_file),
            show_stats=True,
            stdout_stream=stdout,
            stderr_stream=stderr,
        )
        assert exit_code == 0
        assert out_file.exists()
        # Stats should be in stdout when output is a file
        assert "Statistics:" in stdout.getvalue()


def test_cli_unsupported_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        fake_pdf = Path(tmpdir) / "document.pdf"
        fake_pdf.write_text("PDF content")
        stdout = io.StringIO()
        stderr = io.StringIO()

        exit_code = run_convert(
            file_path_str=str(fake_pdf),
            stdout_stream=stdout,
            stderr_stream=stderr,
        )
        assert exit_code == 1
        assert "Unsupported file extension '.pdf'" in stderr.getvalue()


def test_cli_file_not_found():
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = run_convert(
        file_path_str="non_existent_file.docx",
        stdout_stream=stdout,
        stderr_stream=stderr,
    )
    assert exit_code == 1
    assert "File not found" in stderr.getvalue()


def test_cli_main_entry_point():
    html_file = str(SAMPLE_DIR / "sample.html")
    # Calling main directly with argv list
    exit_code = main(["convert", html_file])
    assert exit_code == 0


def test_cli_all_sample_formats():
    for ext in ("sample.html", "sample.docx", "sample.csv", "sample.xlsx"):
        file_path = SAMPLE_DIR / ext
        stdout = io.StringIO()
        stderr = io.StringIO()

        exit_code = run_convert(
            file_path_str=str(file_path),
            show_stats=True,
            stdout_stream=stdout,
            stderr_stream=stderr,
        )
        assert exit_code == 0, f"Failed converting {ext}: {stderr.getvalue()}"
        assert len(stdout.getvalue().strip()) > 0
        assert "Statistics:" in stderr.getvalue()
