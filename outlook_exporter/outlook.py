"""Outlook COM interaction helpers."""

from typing import List


def search_mail(keywords: str) -> List[str]:
    """Return dummy list of email subjects containing the keyword."""
    sample = ["Meeting transcript", "Other mail", f"Keyword {keywords} found"]
    return [s for s in sample if keywords.lower() in s.lower()]
