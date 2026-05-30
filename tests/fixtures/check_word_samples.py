#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Smoke-check generated Word fixtures against the formatter diagnostics."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

try:
    from .generate_word_samples import OUTPUT_DIR, generate_all
except ImportError:
    from generate_word_samples import OUTPUT_DIR, generate_all


ROOT = Path(__file__).resolve().parents[2]
FORMATTER = ROOT / "skills" / "gongwen-format-converter" / "scripts" / "format_document.py"
REPORT_DIR = Path("/private/tmp/gongwen-fixture-reports")


EXPECTED_SIGNALS = {
    "01-basic-formal.docx": {
        ("format_diagnostics", "paragraph_controls", "counts", "manual_numbering_paragraph_count"): 3,
        ("format_diagnostics", "paragraph_controls", "manual_numbering_level_counts", "1"): 1,
        ("format_diagnostics", "paragraph_controls", "manual_numbering_level_counts", "2"): 1,
        ("format_diagnostics", "paragraph_controls", "manual_numbering_level_counts", "3"): 1,
    },
    "02-table.docx": {
        ("format_diagnostics", "tables", "table_count"): 1,
        ("format_diagnostics", "tables", "tables", 0, "merged_cell_indicators", "grid_span_count"): 1,
    },
    "03-header-footer-page.docx": {
        ("format_diagnostics", "headers_footers", "sections", 0, "page_field_count"): 1,
        ("format_diagnostics", "headers_footers", "sections", 0, "footer_parts", "default", "num_pages_field_count"): 1,
        ("format_diagnostics", "page", "section_count"): 2,
        ("format_diagnostics", "special_state", "all_field_type_counts", "PAGE"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "all", "type_counts", "NUMPAGES"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "all", "category_counts", "total_pages"): 1,
        ("content_preservation", "before_scope_counts", "header_footer_paragraph"): 1,
    },
    "04-image-seal.docx": {
        ("format_diagnostics", "objects", "inline_shape_count"): 1,
    },
    "05-textbox-shape.docx": {
        ("format_diagnostics", "objects", "text_box_count"): 1,
        ("format_diagnostics", "objects", "legacy_pict_count"): 1,
        ("content_preservation", "before_scope_counts", "text_box_paragraph"): 1,
    },
    "06-auto-numbering.docx": {
        ("format_diagnostics", "paragraph_controls", "counts", "word_numbering_paragraph_count"): 3,
        ("format_diagnostics", "paragraph_controls", "counts", "manual_numbering_paragraph_count"): 1,
        ("format_diagnostics", "paragraph_controls", "word_numbering_num_id_counts", "5"): 3,
        ("format_diagnostics", "paragraph_controls", "word_numbering_level_counts", "0"): 3,
        ("format_diagnostics", "paragraph_controls", "manual_numbering_level_counts", "1"): 1,
    },
    "07-toc-fields.docx": {
        ("format_diagnostics", "special_state", "field_char_count"): 12,
        ("format_diagnostics", "special_state", "body_field_type_counts", "TOC"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "body", "type_counts", "DATE"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "body", "type_counts", "REF"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "body", "type_counts", "PAGEREF"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "body", "category_counts", "table_of_contents"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "body", "category_counts", "date_time"): 1,
        ("format_diagnostics", "special_state", "field_diagnostics", "body", "category_counts", "cross_reference"): 2,
        ("format_diagnostics", "special_state", "bookmark_count"): 1,
    },
    "08-comments-revisions.docx": {
        ("format_diagnostics", "special_state", "comments_relationship_count"): 1,
        ("format_diagnostics", "special_state", "footnotes_relationship_count"): 1,
        ("format_diagnostics", "special_state", "endnotes_relationship_count"): 1,
        ("format_diagnostics", "special_state", "tracked_insertions"): 1,
        ("format_diagnostics", "special_state", "tracked_deletions"): 1,
        ("format_diagnostics", "special_state", "footnote_reference_count"): 1,
        ("format_diagnostics", "special_state", "endnote_reference_count"): 1,
        ("content_preservation", "before_scope_counts", "comments_paragraph"): 1,
        ("content_preservation", "before_scope_counts", "footnotes_paragraph"): 1,
        ("content_preservation", "before_scope_counts", "endnotes_paragraph"): 1,
    },
}


def nested_value(data: dict, path: tuple) -> object:
    value: object = data
    for key in path:
        if isinstance(key, int):
            value = value[key]  # type: ignore[index]
        else:
            value = value[key]  # type: ignore[index]
    return value


def run_diagnostics(sample: Path, report: Path) -> None:
    subprocess.run(
        [
            sys.executable,
            str(FORMATTER),
            str(sample),
            "--diagnose-only",
            "--preset",
            "formal",
            "--report",
            str(report),
        ],
        cwd=ROOT,
        check=True,
    )


def main() -> int:
    generated = generate_all()
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    failures: list[str] = []

    for sample in generated:
        report = REPORT_DIR / f"{sample.stem}.json"
        run_diagnostics(sample, report)
        data = json.loads(report.read_text(encoding="utf-8"))
        for path, minimum in EXPECTED_SIGNALS.get(sample.name, {}).items():
            actual = nested_value(data, path)
            if not isinstance(actual, int) or actual < minimum:
                failures.append(
                    f"{sample.name}: {'.'.join(str(part) for part in path)} "
                    f"expected >= {minimum}, got {actual!r}"
                )

    if failures:
        print("Word fixture smoke check failed:", file=sys.stderr)
        for failure in failures:
            print(f"- {failure}", file=sys.stderr)
        return 1

    print(f"Generated and checked {len(generated)} Word fixtures in {OUTPUT_DIR.relative_to(ROOT)}")
    print(f"Diagnostic reports: {REPORT_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
