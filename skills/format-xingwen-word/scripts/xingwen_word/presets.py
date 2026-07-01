"""Fixed 行文检查版式 presets and report scope constants."""

from __future__ import annotations

from typing import Any


PRESETS: dict[str, dict[str, dict[str, Any]]] = {
    "formal": {
        "_page": {
            "width_mm": 210,
            "height_mm": 297,
            "top_mm": 37,
            "bottom_mm": 35,
            "left_mm": 27,
            "right_mm": 27,
            "text_area_width_mm": 156,
            "text_area_height_mm": 225,
            "grid_line_count": 22,
            "grid_char_count": 28,
            "grid_type": "linesAndChars",
            "line_pitch_twips": 600,
            "char_space_twips": 316,
        },
        "main_title": {
            "font": "方正小标宋简体",
            "latin_font": "Times New Roman",
            "size": 22,
            "bold": False,
            "align": "center",
            "first_indent": 0,
            "line": 36,
            "space_after": 30,
            "keep_together": True,
            "keep_with_next": True,
            "widow_control": True,
            "outline_level": "0",
        },
        "subtitle": {"font": "楷体_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "center", "first_indent": 0, "line": 30, "widow_control": True},
        "recipient": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "left", "first_indent": 0, "line": 30, "keep_with_next": True, "widow_control": True},
        "body": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "justify", "first_indent": 32, "line": 30, "widow_control": True},
        "heading_1": {"font": "黑体", "latin_font": "Times New Roman", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 30, "keep_together": True, "keep_with_next": True, "widow_control": True, "outline_level": "0"},
        "heading_2": {"font": "楷体_GB2312", "latin_font": "Times New Roman", "size": 16, "bold": True, "align": "justify", "first_indent": 32, "line": 30, "keep_together": True, "keep_with_next": True, "widow_control": True, "outline_level": "1"},
        "heading_3": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 30, "keep_together": True, "keep_with_next": True, "widow_control": True, "outline_level": "2"},
        "heading_4": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 30, "keep_together": True, "keep_with_next": True, "widow_control": True, "outline_level": "3"},
        "attachment": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "justify", "first_indent": 32, "line": 30, "widow_control": True},
        "signature": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "right", "first_indent": 0, "line": 30, "widow_control": True},
        "date": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "right", "first_indent": 0, "line": 30, "widow_control": True},
        "needs_review": {"font": "仿宋_GB2312", "latin_font": "Times New Roman", "size": 16, "align": "justify", "first_indent": 32, "line": 30, "widow_control": True},
    },
}
PRESETS["checklist"] = PRESETS["formal"]

COMMON_REPORT_ROLES = [
    "main_title",
    "subtitle",
    "issue_number",
    "metadata",
    "separator",
    "article_title",
    "recipient",
    "body",
    "heading_1",
    "heading_2",
    "heading_3",
    "heading_4",
    "attachment",
    "note",
    "signature",
    "date",
    "cc",
    "printing_area",
]

LIMITED_FORMAT_AREAS = [
    {
        "area": "floating_text_boxes_and_shapes",
        "status": "unsupported",
        "note": "Floating text boxes and shapes are preserved; python-docx does not expose complete safe editing for them.",
    },
    {
        "area": "comments_tracked_changes_and_fields",
        "status": "unsupported",
        "note": "Comments, tracked changes, and field codes are preserved when present but are not fully normalized by this script.",
    },
    {
        "area": "watermarks_footnotes_endnotes_and_bookmarks",
        "status": "unsupported",
        "note": "These structures require deeper OOXML handling and are reported as outside the current automatic formatting scope.",
    },
]
