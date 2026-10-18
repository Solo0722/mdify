"""Behavior tests for convert_csv(), generated for Assignment 3.

Expected outputs are explicit, not calculated with the converter or its helpers.
All inputs are fixed strings: these tests need no files, network, clock, or mocks.
The first CSV record becomes the Markdown header; raw text stays unescaped.
"""

import pytest

from mdify.convert_tabular import convert_csv


@pytest.mark.parametrize(
    "separator",
    [",", ";", "\t", "|"],
    ids=["comma", "semicolon", "tab", "pipe"],
)
def test_detects_supported_delimiters(separator):
    content = f"Name{separator}Count\nApple{separator}2\nPear{separator}5\n"

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Name | Count |\n"
        "| --- | --- |\n"
        "| Apple | 2 |\n"
        "| Pear | 5 |"
    )
    assert raw_text == "Name Count\nApple 2\nPear 5"


def test_explicit_delimiter_preserves_commas_inside_values():
    content = "Name^Details\nBolt^steel, small\nNut^brass, large\n"

    markdown, raw_text = convert_csv(content, delimiter="^")

    assert markdown == (
        "| Name | Details |\n"
        "| --- | --- |\n"
        "| Bolt | steel, small |\n"
        "| Nut | brass, large |"
    )
    assert raw_text == "Name Details\nBolt steel, small\nNut brass, large"


@pytest.mark.parametrize(
    "content", ["", " \r\n\t", "\ufeff"], ids=["empty", "whitespace", "bom-only"]
)
def test_empty_input_has_no_table_or_raw_text(content):
    assert convert_csv(content) == ("", "")


@pytest.mark.parametrize(
    "content",
    [",,\n,,\n", '""," "\n,\n'],
    ids=["empty-cells", "quoted-whitespace"],
)
def test_input_with_only_empty_records_has_no_header(content):
    assert convert_csv(content) == ("", "")


def test_bom_is_removed_from_the_first_header():
    markdown, raw_text = convert_csv("\ufeffName,Count\nApple,2\n")

    assert markdown == "| Name | Count |\n| --- | --- |\n| Apple | 2 |"
    assert raw_text == "Name Count\nApple 2"


def test_quoted_comma_and_quotes_stay_in_one_cell_and_pipes_are_escaped():
    content = 'Name,Note\nAlice,"hello, ""world"" | yes"\n'

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Name | Note |\n"
        "| --- | --- |\n"
        '| Alice | hello, "world" \\| yes |'
    )
    assert raw_text == 'Name Note\nAlice hello, "world" | yes'


@pytest.mark.parametrize("line_break", ["\n", "\r\n", "\r"], ids=["lf", "crlf", "cr"])
def test_quoted_line_breaks_stay_inside_their_cell(line_break):
    content = f'Name,Note\r\nAlice,"first{line_break}second"\r\nBob,plain\r\n'

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Name | Note |\n"
        "| --- | --- |\n"
        "| Alice | first<br>second |\n"
        "| Bob | plain |"
    )
    # Only Markdown normalizes embedded line endings to <br>.
    assert raw_text == f"Name Note\nAlice first{line_break}second\nBob plain"


def test_single_column_input_does_not_treat_letters_as_delimiters():
    markdown, raw_text = convert_csv("hello\nworld\n")

    assert markdown == "| hello |\n| --- |\n| world |"
    assert raw_text == "hello\nworld"


def test_blank_records_are_skipped_but_partial_records_and_zero_are_kept():
    content = "Name,Count\n\n, \nApple,2\n,3\nPear,\n0,0\n"

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Name | Count |\n"
        "| --- | --- |\n"
        "| Apple | 2 |\n"
        "|  | 3 |\n"
        "| Pear |  |\n"
        "| 0 | 0 |"
    )
    # Raw text separates cells with spaces, including empty cell positions.
    assert raw_text == "Name Count\nApple 2\n 3\nPear \n0 0"


def test_inconsistent_row_widths_fall_back_to_comma_and_preserve_extra_cells():
    content = "Name,Count\nApple,2,ripe\nPear\n"

    markdown, raw_text = convert_csv(content)

    # Padding must reach the widest row, even when that row is not the header.
    assert markdown == (
        "| Name | Count |  |\n"
        "| --- | --- | --- |\n"
        "| Apple | 2 | ripe |\n"
        "| Pear |  |  |"
    )
    assert raw_text == "Name Count\nApple 2 ripe\nPear"


def test_order_duplicates_and_leading_zeroes_are_preserved():
    content = "Item,Code\nNut,007\nBolt,001\nNut,007\nWasher,000\n"

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Item | Code |\n"
        "| --- | --- |\n"
        "| Nut | 007 |\n"
        "| Bolt | 001 |\n"
        "| Nut | 007 |\n"
        "| Washer | 000 |"
    )
    assert raw_text == "Item Code\nNut 007\nBolt 001\nNut 007\nWasher 000"


def test_unicode_values_are_preserved_in_both_outputs():
    content = "Name,City\nZoë,Zürich\n李雷,北京\n"

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Name | City |\n"
        "| --- | --- |\n"
        "| Zoë | Zürich |\n"
        "| 李雷 | 北京 |"
    )
    assert raw_text == "Name City\nZoë Zürich\n李雷 北京"


def test_markdown_punctuation_is_literal_but_raw_text_is_not_escaped():
    content = "Text,Value\n*literal*,[x](y)\n"

    markdown, raw_text = convert_csv(content)

    assert markdown == (
        "| Text | Value |\n"
        "| --- | --- |\n"
        r"| \*literal\* | \[x\](y) |"
    )
    assert raw_text == "Text Value\n*literal* [x](y)"


@pytest.mark.parametrize(
    "delimiter",
    ["", "::", "\n", "\r", "\x00", '"'],
    ids=["empty", "multiple-characters", "lf", "cr", "nul", "quote"],
)
def test_invalid_delimiters_raise_a_validation_error(delimiter):
    with pytest.raises(ValueError, match="delimiter"):
        convert_csv("Name,Count\nApple,2\n", delimiter=delimiter)
