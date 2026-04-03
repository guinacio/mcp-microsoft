from __future__ import annotations

import html as html_module
import re


def strip_html(text: str) -> str:
    """Convert a small HTML fragment into readable plain text."""
    plain = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", text, flags=re.DOTALL | re.IGNORECASE)
    plain = re.sub(r"<br\s*/?>", "\n", plain, flags=re.IGNORECASE)
    plain = re.sub(r"</p>", "\n\n", plain, flags=re.IGNORECASE)
    plain = re.sub(r"</?(div|tr|li|h[1-6]|blockquote)[^>]*>", "\n", plain, flags=re.IGNORECASE)
    plain = re.sub(r"<[^>]+>", "", plain)
    plain = html_module.unescape(plain)
    plain = re.sub(r"\n{3,}", "\n\n", plain)
    return plain.strip()
