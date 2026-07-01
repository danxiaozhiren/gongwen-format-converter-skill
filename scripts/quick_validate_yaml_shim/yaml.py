"""Minimal PyYAML-compatible shim for skill quick_validate.py.

The upstream validator only needs `safe_load` for simple SKILL.md frontmatter
and `YAMLError` for parse failures. This shim keeps validation runnable in the
repo's default offline test environment without adding a network-installed
dependency.
"""

from __future__ import annotations


class YAMLError(Exception):
    """Raised when the minimal frontmatter parser cannot parse input."""


def safe_load(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in line:
            raise YAMLError(f"line {line_number}: expected key: value")
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        if not key:
            raise YAMLError(f"line {line_number}: empty key")
        if value.startswith(("'", '"')) or value.endswith(("'", '"')):
            if len(value) < 2 or value[0] != value[-1]:
                raise YAMLError(f"line {line_number}: unmatched quote")
            value = value[1:-1]
        result[key] = value
    return result
