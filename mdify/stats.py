"""Character and token statistics for mdify conversions."""

from dataclasses import dataclass


@dataclass(frozen=True)
class TextStats:
    characters: int
    estimated_tokens: float

    @property
    def estimated_tokens_int(self) -> int:
        """Returns estimated tokens rounded to nearest integer."""
        return round(self.estimated_tokens)


@dataclass(frozen=True)
class ConversionStats:
    original: TextStats
    markdown: TextStats


def calculate_stats(original_text: str, markdown_text: str) -> ConversionStats:
    """Calculate character and estimated token counts for original and markdown text.

    Estimated token count is calculated as characters divided by 4.
    """
    orig_chars = len(original_text)
    orig_tokens = orig_chars / 4.0

    md_chars = len(markdown_text)
    md_tokens = md_chars / 4.0

    return ConversionStats(
        original=TextStats(characters=orig_chars, estimated_tokens=orig_tokens),
        markdown=TextStats(characters=md_chars, estimated_tokens=md_tokens),
    )


def format_stats(stats: ConversionStats) -> str:
    """Format conversion statistics into a clean, human-readable string."""
    orig_tokens_str = (
        f"{stats.original.estimated_tokens_int}"
        if stats.original.estimated_tokens.is_integer()
        else f"{stats.original.estimated_tokens:.2f}"
    )
    md_tokens_str = (
        f"{stats.markdown.estimated_tokens_int}"
        if stats.markdown.estimated_tokens.is_integer()
        else f"{stats.markdown.estimated_tokens:.2f}"
    )

    lines = [
        "Statistics:",
        "  Original text:",
        f"    Characters: {stats.original.characters}",
        f"    Estimated tokens: {orig_tokens_str} (estimate: characters / 4)",
        "  Markdown output:",
        f"    Characters: {stats.markdown.characters}",
        f"    Estimated tokens: {md_tokens_str} (estimate: characters / 4)",
    ]
    return "\n".join(lines)
