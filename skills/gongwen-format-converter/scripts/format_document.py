#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Format Chinese official-document and internal-brief drafts as .docx files.

The script is intentionally conservative: it preserves paragraph text, applies
role-based formatting, and keeps reports free of full paragraph content by
default.
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    from docx import Document
    from docx.enum.section import WD_SECTION_START
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor
except ModuleNotFoundError as exc:
    if exc.name != "docx":
        raise
    Document = None
    WD_SECTION_START = None
    WD_STYLE_TYPE = None
    WD_ALIGN_PARAGRAPH = None
    WD_LINE_SPACING = None
    OxmlElement = None
    qn = None
    Mm = None
    Pt = None
    RGBColor = None
    DOCX_IMPORT_ERROR: ModuleNotFoundError | None = exc
else:
    DOCX_IMPORT_ERROR = None


CHINESE_NUM = "一二三四五六七八九十"
DATE_RE = re.compile(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")
ISSUE_RE = re.compile(r"^第\s*[\d%s]+\s*期$" % CHINESE_NUM)
H1_RE = re.compile(r"^[%s]+、" % CHINESE_NUM)
H2_RE = re.compile(r"^（[%s]+）" % CHINESE_NUM)
H3_RE = re.compile(r"^\d+\s*[.．、]")
SEPARATOR_RE = re.compile(r"^[\-_—=─━]{3,}$")


@dataclass
class ParagraphItem:
    text: str
    role_hint: str | None = None
    paragraph: Any | None = None


def pt_value(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value.pt)
    except AttributeError:
        try:
            return float(value)
        except Exception:
            return None


def mm_value(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return round(float(value.mm), 2)
    except AttributeError:
        try:
            return round(float(value), 2)
        except Exception:
            return None


def bool_or_none(value: Any) -> bool | None:
    if value is None:
        return None
    return bool(value)


def enum_name(value: Any) -> str | None:
    if value is None:
        return None
    name = getattr(value, "name", None)
    if name:
        return str(name)
    return str(value)


def safe_xpath(element: Any, expression: str) -> list[Any]:
    if element is None:
        return []
    try:
        return list(element.xpath(expression))
    except Exception:
        return []


def safe_xpath_count(element: Any, expression: str) -> int:
    return len(safe_xpath(element, expression))


def xml_child_val(element: Any, path: str) -> str | None:
    matches = safe_xpath(element, path)
    if not matches:
        return None
    try:
        return matches[0].get(qn("w:val"))
    except Exception:
        return None


def require_docx() -> None:
    if DOCX_IMPORT_ERROR is None:
        return
    requirements_path = Path(__file__).with_name("requirements.txt")
    raise SystemExit(
        "Missing dependency: python-docx.\n"
        f'Install script dependencies with: python -m pip install -r "{requirements_path}"\n'
        "Or install directly with: python -m pip install python-docx"
    )


def text_hash(text: str) -> str:
    return hashlib.sha256(text.strip().encode("utf-8")).hexdigest()[:12]


def exact_text_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:12]


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def is_short_title_like(text: str) -> bool:
    if len(text) > 40:
        return False
    return not text.endswith(("。", "；", ";", "，", ","))


def classify_paragraph(
    item: ParagraphItem,
    index: int,
    total: int,
    preset: str,
    state: dict[str, Any],
) -> tuple[str, list[str]]:
    text = normalize_text(item.text)
    warnings: list[str] = []

    if item.role_hint:
        return item.role_hint, warnings

    if not text:
        return "empty", warnings

    if SEPARATOR_RE.match(text):
        state["seen_separator"] = True
        return "separator", warnings

    if H1_RE.match(text):
        return "heading_1", warnings
    if H2_RE.match(text):
        return "heading_2", warnings
    if H3_RE.match(text):
        return "heading_3", warnings

    if text.startswith("附件"):
        return "attachment", warnings

    if index == 0:
        return "main_title", warnings

    if preset == "brief":
        if index <= 2 and ISSUE_RE.match(text):
            return "issue_number", warnings
        if index <= 3 and DATE_RE.search(text):
            return "metadata", warnings
        if (
            state.get("seen_separator")
            and not state.get("article_title_assigned")
            and is_short_title_like(text)
        ):
            state["article_title_assigned"] = True
            return "article_title", warnings
        if (
            index <= 4
            and not state.get("article_title_assigned")
            and is_short_title_like(text)
            and not ISSUE_RE.match(text)
        ):
            state["article_title_assigned"] = True
            return "article_title", warnings

    if text.endswith("：") and index < max(6, total // 4) and len(text) <= 30:
        return "recipient", warnings

    if index >= total - 3 and DATE_RE.fullmatch(text):
        return "date", warnings

    if index >= total - 4 and len(text) <= 30 and re.search(
        r"(部|处|办|办公室|中心|公司|委员会|局|厅|院|所)$", text
    ):
        return "signature", warnings

    if is_short_title_like(text) and len(text) <= 18 and index < 6:
        warnings.append("short title-like paragraph; treated as body unless confirmed")
        return "needs_review", warnings

    return "body", warnings


ALIGNMENTS = (
    {}
    if WD_ALIGN_PARAGRAPH is None
    else {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
)


PRESETS: dict[str, dict[str, dict[str, Any]]] = {
    "formal": {
        "_page": {"width_mm": 210, "height_mm": 297, "top_mm": 37, "bottom_mm": 35, "left_mm": 28, "right_mm": 26},
        "main_title": {
            "font": "方正小标宋简体",
            "size": 22,
            "bold": False,
            "align": "center",
            "first_indent": 0,
            "line": 32,
            "space_after": 0,
        },
        "subtitle": {"font": "楷体_GB2312", "size": 16, "align": "center", "first_indent": 0},
        "recipient": {"font": "仿宋_GB2312", "size": 16, "align": "left", "first_indent": 0},
        "body": {"font": "仿宋_GB2312", "size": 16, "align": "justify", "first_indent": 32, "line": 28},
        "heading_1": {"font": "黑体", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 28},
        "heading_2": {"font": "楷体_GB2312", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 28},
        "heading_3": {"font": "仿宋_GB2312", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 28},
        "attachment": {"font": "仿宋_GB2312", "size": 16, "align": "justify", "first_indent": 32, "line": 28},
        "signature": {"font": "仿宋_GB2312", "size": 16, "align": "right", "first_indent": 0, "line": 28},
        "date": {"font": "仿宋_GB2312", "size": 16, "align": "right", "first_indent": 0, "line": 28},
        "needs_review": {"font": "仿宋_GB2312", "size": 16, "align": "justify", "first_indent": 32, "line": 28},
    },
    "brief": {
        "_page": {"width_mm": 210, "height_mm": 297, "top_mm": 25, "bottom_mm": 25, "left_mm": 28, "right_mm": 26},
        "main_title": {"font": "宋体", "size": 22, "bold": True, "align": "center", "first_indent": 0, "line": 32},
        "issue_number": {"font": "楷体_GB2312", "size": 16, "bold": False, "align": "center", "first_indent": 0, "line": 28},
        "metadata": {"font": "仿宋_GB2312", "size": 16, "bold": False, "align": "left", "first_indent": 0, "line": 28},
        "separator": {"font": "宋体", "size": 12, "bold": False, "align": "center", "first_indent": 0, "line": 18},
        "article_title": {"font": "宋体", "size": 22, "bold": True, "align": "center", "first_indent": 0, "line": 32},
        "body": {"font": "仿宋_GB2312", "size": 16, "align": "justify", "first_indent": 32, "line": 28},
        "heading_1": {"font": "黑体", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 28},
        "heading_2": {"font": "楷体_GB2312", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 28},
        "heading_3": {"font": "仿宋_GB2312", "size": 16, "bold": False, "align": "justify", "first_indent": 32, "line": 28},
        "attachment": {"font": "仿宋_GB2312", "size": 16, "align": "justify", "first_indent": 32, "line": 28},
        "signature": {"font": "仿宋_GB2312", "size": 16, "align": "right", "first_indent": 0, "line": 28},
        "date": {"font": "仿宋_GB2312", "size": 16, "align": "right", "first_indent": 0, "line": 28},
        "needs_review": {"font": "仿宋_GB2312", "size": 16, "align": "justify", "first_indent": 32, "line": 28},
    },
}

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


def set_east_asia_font(run: Any, font_name: str) -> None:
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)


def apply_style(paragraph: Any, style: dict[str, Any]) -> None:
    fmt = paragraph.paragraph_format
    align = style.get("align")
    if align in ALIGNMENTS:
        paragraph.alignment = ALIGNMENTS[align]
    if "left_indent" in style:
        fmt.left_indent = Pt(style["left_indent"])
    if "right_indent" in style:
        fmt.right_indent = Pt(style["right_indent"])
    if "first_indent" in style:
        fmt.first_line_indent = Pt(style["first_indent"])
    if "space_before" in style:
        fmt.space_before = Pt(style["space_before"])
    else:
        fmt.space_before = Pt(0)
    if "space_after" in style:
        fmt.space_after = Pt(style["space_after"])
    else:
        fmt.space_after = Pt(0)
    if "line" in style and style["line"]:
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        fmt.line_spacing = Pt(style["line"])

    font = style.get("font", "仿宋_GB2312")
    size = style.get("size", 16)
    bold = style.get("bold")
    italic = style.get("italic", False)
    underline = style.get("underline", False)
    color = style.get("color", "000000")
    for run in paragraph.runs:
        set_east_asia_font(run, font)
        run.font.size = Pt(size)
        if bold is not None:
            run.font.bold = bool(bold)
        if italic is not None:
            run.font.italic = bool(italic)
        if underline is not None:
            run.font.underline = bool(underline)
        if color:
            try:
                run.font.color.rgb = RGBColor.from_string(str(color).replace("#", ""))
            except ValueError:
                pass


def apply_page_setup(document: Any, page_style: dict[str, Any]) -> None:
    for section in document.sections:
        section.start_type = WD_SECTION_START.NEW_PAGE
        if "width_mm" in page_style:
            section.page_width = Mm(page_style["width_mm"])
        if "height_mm" in page_style:
            section.page_height = Mm(page_style["height_mm"])
        if "top_mm" in page_style:
            section.top_margin = Mm(page_style["top_mm"])
        if "bottom_mm" in page_style:
            section.bottom_margin = Mm(page_style["bottom_mm"])
        if "left_mm" in page_style:
            section.left_margin = Mm(page_style["left_mm"])
        if "right_mm" in page_style:
            section.right_margin = Mm(page_style["right_mm"])


def read_plain_or_markdown(path: Path | None, stdin: bool, input_name: str) -> tuple[list[ParagraphItem], str]:
    if stdin:
        text = sys.stdin.read()
        suffix = Path(input_name).suffix.lower()
    else:
        assert path is not None
        text = path.read_text(encoding="utf-8")
        suffix = path.suffix.lower()

    items: list[ParagraphItem] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        role_hint = None
        if suffix in {".md", ".markdown"}:
            match = re.match(r"^(#{1,6})\s+(.+)$", line)
            if match:
                level = len(match.group(1))
                line = match.group(2).strip()
                if level == 1:
                    role_hint = "main_title"
                elif level == 2:
                    role_hint = "article_title"
                elif level == 3:
                    role_hint = "heading_1"
                elif level == 4:
                    role_hint = "heading_2"
                else:
                    role_hint = "heading_3"
        items.append(ParagraphItem(text=line, role_hint=role_hint))
    return items, suffix


def iter_table_paragraphs(document: Any) -> list[Any]:
    paragraphs = []
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    return paragraphs


def document_text_units(document: Any) -> list[dict[str, Any]]:
    units: list[dict[str, Any]] = []
    for idx, paragraph in enumerate(document.paragraphs):
        units.append(
            {
                "identity": f"document.paragraphs[{idx}]",
                "scope": "body_paragraph",
                "text": paragraph.text,
            }
        )
    for table_idx, table in enumerate(document.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for paragraph_idx, paragraph in enumerate(cell.paragraphs):
                    units.append(
                        {
                            "identity": (
                                f"tables[{table_idx}].rows[{row_idx}]"
                                f".cells[{cell_idx}].paragraphs[{paragraph_idx}]"
                            ),
                            "scope": "table_cell_paragraph",
                            "text": paragraph.text,
                        }
                    )
    return units


def document_text_fingerprint(document: Any) -> dict[str, Any]:
    units = document_text_units(document)
    unit_fingerprints = [
        {
            "identity": unit["identity"],
            "scope": unit["scope"],
            "hash": exact_text_hash(unit["text"]),
            "length": len(unit["text"]),
        }
        for unit in units
    ]
    combined = hashlib.sha256(
        json.dumps(unit_fingerprints, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()[:16]
    return {
        "scope": "all body paragraph and table-cell paragraph text after input parsing",
        "unit_count": len(unit_fingerprints),
        "non_empty_unit_count": len([unit for unit in units if unit["text"].strip()]),
        "combined_hash": combined,
        "units": unit_fingerprints,
    }


def compare_text_fingerprints(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    before_units = before.get("units", [])
    after_units = after.get("units", [])
    changed_indexes = []
    for idx, (before_unit, after_unit) in enumerate(zip(before_units, after_units), start=1):
        if (
            before_unit.get("identity") != after_unit.get("identity")
            or before_unit.get("hash") != after_unit.get("hash")
            or before_unit.get("length") != after_unit.get("length")
        ):
            changed_indexes.append(idx)
    if len(before_units) != len(after_units):
        changed_indexes.extend(range(min(len(before_units), len(after_units)) + 1, max(len(before_units), len(after_units)) + 1))

    return {
        "text_changed": before.get("combined_hash") != after.get("combined_hash"),
        "paragraph_or_table_text_unit_count_changed": before.get("unit_count") != after.get("unit_count"),
        "paragraph_order_changed": [
            unit.get("identity") for unit in before_units
        ] != [
            unit.get("identity") for unit in after_units
        ],
        "generated_missing_elements": False,
        "before_unit_count": before.get("unit_count"),
        "after_unit_count": after.get("unit_count"),
        "before_non_empty_unit_count": before.get("non_empty_unit_count"),
        "after_non_empty_unit_count": after.get("non_empty_unit_count"),
        "before_text_hash": before.get("combined_hash"),
        "after_text_hash": after.get("combined_hash"),
        "changed_unit_indexes": changed_indexes[:20],
        "scope": before.get("scope"),
    }


def document_items(document: Any) -> list[ParagraphItem]:
    return [
        ParagraphItem(text=p.text, paragraph=p)
        for p in document.paragraphs
        if p.text and p.text.strip()
    ]


def create_document_from_items(items: list[ParagraphItem]) -> Any:
    document = Document()
    if document.paragraphs and not document.paragraphs[0].text:
        first = document.paragraphs[0]
        if items:
            first.add_run(items[0].text)
            items[0].paragraph = first
            rest = items[1:]
        else:
            rest = []
    else:
        rest = items
    for item in rest:
        p = document.add_paragraph()
        p.add_run(item.text)
        item.paragraph = p
    return document


def extract_run_font(run: Any) -> str | None:
    rpr = run._element.rPr
    if rpr is not None and rpr.rFonts is not None:
        east_asia = rpr.rFonts.get(qn("w:eastAsia"))
        if east_asia:
            return east_asia
    return run.font.name


def extract_run_style(run: Any) -> dict[str, Any]:
    style: dict[str, Any] = {}
    font = extract_run_font(run)
    if font:
        style["font"] = font
    if run.font.size is not None:
        style["size"] = pt_value(run.font.size)
    if run.font.bold is not None:
        style["bold"] = bool(run.font.bold)
    if run.font.italic is not None:
        style["italic"] = bool(run.font.italic)
    if run.font.underline is not None:
        style["underline"] = bool(run.font.underline)
    if run.font.strike is not None:
        style["strikethrough"] = bool(run.font.strike)
    if run.font.superscript:
        style["vertical_align"] = "superscript"
    elif run.font.subscript:
        style["vertical_align"] = "subscript"
    if run.font.all_caps is not None:
        style["all_caps"] = bool(run.font.all_caps)
    if run.font.small_caps is not None:
        style["small_caps"] = bool(run.font.small_caps)
    if run.font.highlight_color is not None:
        style["highlight"] = enum_name(run.font.highlight_color)
    if run.font.color is not None and run.font.color.rgb is not None:
        style["color"] = str(run.font.color.rgb)
    char_spacing = xml_child_val(run._element, "./w:rPr/w:spacing")
    if char_spacing is not None:
        style["character_spacing_twips"] = char_spacing
    emphasis = xml_child_val(run._element, "./w:rPr/w:em")
    if emphasis is not None:
        style["emphasis_mark"] = emphasis
    vanish = safe_xpath_count(run._element, "./w:rPr/w:vanish")
    if vanish:
        style["hidden"] = True
    return style


def extract_paragraph_style(paragraph: Any) -> dict[str, Any]:
    fmt = paragraph.paragraph_format
    style: dict[str, Any] = {}
    if paragraph.style is not None:
        style["style_name"] = paragraph.style.name
        style["style_id"] = paragraph.style.style_id
    if paragraph.alignment is not None:
        reverse = {value: key for key, value in ALIGNMENTS.items()}
        style["align"] = reverse.get(paragraph.alignment)
    if fmt.left_indent is not None:
        style["left_indent"] = pt_value(fmt.left_indent)
    if fmt.right_indent is not None:
        style["right_indent"] = pt_value(fmt.right_indent)
    if fmt.first_line_indent is not None:
        style["first_indent"] = pt_value(fmt.first_line_indent)
    if fmt.line_spacing is not None:
        style["line"] = pt_value(fmt.line_spacing)
    if fmt.space_before is not None:
        style["space_before"] = pt_value(fmt.space_before)
    if fmt.space_after is not None:
        style["space_after"] = pt_value(fmt.space_after)
    if fmt.keep_together is not None:
        style["keep_together"] = bool(fmt.keep_together)
    if fmt.keep_with_next is not None:
        style["keep_with_next"] = bool(fmt.keep_with_next)
    if fmt.page_break_before is not None:
        style["page_break_before"] = bool(fmt.page_break_before)
    if fmt.widow_control is not None:
        style["widow_control"] = bool(fmt.widow_control)
    outline_level = xml_child_val(paragraph._element, "./w:pPr/w:outlineLvl")
    if outline_level is not None:
        style["outline_level"] = outline_level
    num_ids = safe_xpath(paragraph._element, "./w:pPr/w:numPr/w:numId")
    if num_ids:
        style["numbering"] = {
            "num_id": num_ids[0].get(qn("w:val")),
            "level": xml_child_val(paragraph._element, "./w:pPr/w:numPr/w:ilvl"),
        }
    tab_stops = []
    try:
        for tab in fmt.tab_stops:
            tab_stops.append(
                {
                    "position_pt": pt_value(tab.position),
                    "alignment": enum_name(tab.alignment),
                    "leader": enum_name(tab.leader),
                }
            )
    except Exception:
        tab_stops = []
    if tab_stops:
        style["tab_stops"] = tab_stops

    for run in paragraph.runs:
        if run.text.strip():
            style.update(extract_run_style(run))
            break
    return {k: v for k, v in style.items() if v is not None}


def compact_style(style: dict[str, Any]) -> dict[str, Any]:
    keys = [
        "style_name",
        "style_id",
        "font",
        "size",
        "bold",
        "italic",
        "underline",
        "strikethrough",
        "color",
        "highlight",
        "vertical_align",
        "character_spacing_twips",
        "emphasis_mark",
        "hidden",
        "align",
        "first_indent",
        "left_indent",
        "right_indent",
        "line",
        "space_before",
        "space_after",
        "keep_together",
        "keep_with_next",
        "page_break_before",
        "widow_control",
        "outline_level",
        "numbering",
        "tab_stops",
    ]
    return {key: style.get(key) for key in keys if style.get(key) is not None}


def style_signature(style: dict[str, Any]) -> str:
    return json.dumps(compact_style(style), ensure_ascii=False, sort_keys=True)


def page_diagnostics(document: Any, preset: str) -> dict[str, Any]:
    sections = []
    expected = PRESETS[preset]["_page"]
    for idx, section in enumerate(document.sections):
        sect_pr = section._sectPr
        columns = safe_xpath(sect_pr, "./w:cols")
        doc_grid = safe_xpath(sect_pr, "./w:docGrid")
        gutter = getattr(section, "gutter", None)
        page = {
            "index": idx + 1,
            "width_mm": mm_value(section.page_width),
            "height_mm": mm_value(section.page_height),
            "top_margin_mm": mm_value(section.top_margin),
            "bottom_margin_mm": mm_value(section.bottom_margin),
            "left_margin_mm": mm_value(section.left_margin),
            "right_margin_mm": mm_value(section.right_margin),
            "gutter_mm": mm_value(gutter),
            "header_distance_mm": mm_value(section.header_distance),
            "footer_distance_mm": mm_value(section.footer_distance),
            "different_first_page_header_footer": bool(section.different_first_page_header_footer),
            "section_start_type": enum_name(section.start_type),
            "column_count": columns[0].get(qn("w:num")) if columns else None,
            "column_space_twips": columns[0].get(qn("w:space")) if columns else None,
            "document_grid": {
                "type": doc_grid[0].get(qn("w:type")),
                "line_pitch": doc_grid[0].get(qn("w:linePitch")),
                "char_space": doc_grid[0].get(qn("w:charSpace")),
            } if doc_grid else None,
            "orientation": "landscape"
            if section.page_width and section.page_height and section.page_width > section.page_height
            else "portrait",
        }
        differences = []
        for key, expected_key in [
            ("top_margin_mm", "top_mm"),
            ("bottom_margin_mm", "bottom_mm"),
            ("left_margin_mm", "left_mm"),
            ("right_margin_mm", "right_mm"),
        ]:
            value = page.get(key)
            exp = expected.get(expected_key)
            if value is not None and exp is not None and abs(value - exp) > 1:
                differences.append({"field": key, "actual": value, "expected": exp})
        if page["width_mm"] and page["height_mm"]:
            if abs(page["width_mm"] - 210) > 2 or abs(page["height_mm"] - 297) > 2:
                differences.append({"field": "paper", "actual": f"{page['width_mm']}x{page['height_mm']}mm", "expected": "A4 210x297mm"})
        page["differences_from_preset"] = differences
        sections.append(page)
    return {"section_count": len(sections), "sections": sections}


def header_footer_diagnostics(document: Any) -> dict[str, Any]:
    sections = []
    for idx, section in enumerate(document.sections):
        header_parts = {
            "default": section.header,
            "first_page": section.first_page_header,
            "even_page": section.even_page_header,
        }
        footer_parts = {
            "default": section.footer,
            "first_page": section.first_page_footer,
            "even_page": section.even_page_footer,
        }
        header_details = {}
        footer_details = {}
        for name, part in header_parts.items():
            paragraphs = [p for p in part.paragraphs if p.text.strip()]
            header_details[name] = {
                "has_text": bool(paragraphs),
                "paragraph_count": len(paragraphs),
            }
        for name, part in footer_parts.items():
            paragraphs = [p for p in part.paragraphs if p.text.strip()]
            footer_details[name] = {
                "has_text": bool(paragraphs),
                "paragraph_count": len(paragraphs),
            }
        header_text = " ".join(
            p.text.strip()
            for part in header_parts.values()
            for p in part.paragraphs
            if p.text.strip()
        )
        footer_text = " ".join(
            p.text.strip()
            for part in footer_parts.values()
            for p in part.paragraphs
            if p.text.strip()
        )
        sections.append(
            {
                "index": idx + 1,
                "has_header_text": bool(header_text),
                "has_footer_text": bool(footer_text),
                "different_first_page_header_footer": bool(section.different_first_page_header_footer),
                "header_parts": header_details,
                "footer_parts": footer_details,
            }
        )
    return {"sections": sections}


def table_diagnostics(document: Any) -> dict[str, Any]:
    tables = []
    for idx, table in enumerate(document.tables):
        sample_styles = []
        row_heights = []
        column_widths = []
        cell_margins = []
        header_row_count = 0
        for row in table.rows[:2]:
            if row.height is not None:
                row_heights.append(mm_value(row.height))
            if safe_xpath_count(row._tr, "./w:trPr/w:tblHeader"):
                header_row_count += 1
            for cell in row.cells[:3]:
                if cell.width is not None:
                    column_widths.append(mm_value(cell.width))
                tc_mar = safe_xpath(cell._tc, "./w:tcPr/w:tcMar")
                if tc_mar:
                    margins = {}
                    for side in ["top", "bottom", "left", "right"]:
                        side_nodes = safe_xpath(tc_mar[0], f"./w:{side}")
                        if side_nodes:
                            margins[side] = side_nodes[0].get(qn("w:w"))
                    if margins:
                        cell_margins.append(margins)
                for p in cell.paragraphs:
                    if p.text.strip():
                        sample_styles.append(compact_style(extract_paragraph_style(p)))
                        break
        tbl_pr = table._tbl.tblPr
        tbl_width = safe_xpath(tbl_pr, "./w:tblW")
        tbl_borders = safe_xpath_count(tbl_pr, "./w:tblBorders/*")
        tables.append(
            {
                "index": idx + 1,
                "rows": len(table.rows),
                "columns": len(table.columns),
                "alignment": enum_name(table.alignment),
                "autofit": bool(table.autofit),
                "width_twips": tbl_width[0].get(qn("w:w")) if tbl_width else None,
                "width_type": tbl_width[0].get(qn("w:type")) if tbl_width else None,
                "border_element_count": tbl_borders,
                "header_row_count_in_sample": header_row_count,
                "sample_row_heights_mm": row_heights[:5],
                "sample_cell_widths_mm": column_widths[:8],
                "sample_cell_margins_twips": cell_margins[:3],
                "sample_style_count": len(sample_styles),
                "sample_styles": sample_styles[:3],
            }
        )
    return {"table_count": len(tables), "tables": tables}


def object_diagnostics(document: Any) -> dict[str, Any]:
    inline_shapes = getattr(document, "inline_shapes", [])
    inline_shape_details = []
    for idx in range(min(len(inline_shapes), 10)):
        shape = inline_shapes[idx]
        inline_shape_details.append(
            {
                "index": idx + 1,
                "type": enum_name(getattr(shape, "type", None)),
                "width_mm": mm_value(getattr(shape, "width", None)),
                "height_mm": mm_value(getattr(shape, "height", None)),
            }
        )
    body = document._element.body
    return {
        "inline_shape_count": len(inline_shapes),
        "inline_shape_samples": inline_shape_details,
        "drawing_element_count": safe_xpath_count(body, ".//w:drawing"),
        "legacy_pict_count": safe_xpath_count(body, ".//w:pict"),
        "text_box_count": safe_xpath_count(body, ".//w:txbxContent"),
        "watermark_like_shape_count": safe_xpath_count(body, ".//w:pict//v:shape"),
        "image_or_object_note": "Inline shapes may include images, seals, charts, or embedded objects; inspect manually before moving or resizing.",
    }


def style_system_diagnostics(document: Any) -> dict[str, Any]:
    type_names = {}
    builtin_count = 0
    custom_count = 0
    samples = []
    for style in document.styles:
        style_type = enum_name(style.type)
        type_names[style_type] = type_names.get(style_type, 0) + 1
        if style.builtin:
            builtin_count += 1
        else:
            custom_count += 1
        if len(samples) < 20:
            samples.append(
                {
                    "name": style.name,
                    "style_id": style.style_id,
                    "type": style_type,
                    "builtin": bool(style.builtin),
                }
            )
    return {
        "style_count": len(document.styles),
        "builtin_style_count": builtin_count,
        "custom_style_count": custom_count,
        "style_type_counts": type_names,
        "style_samples": samples,
    }


def special_state_diagnostics(document: Any) -> dict[str, Any]:
    body = document._element.body
    reltypes = [rel.reltype for rel in document.part.rels.values()]
    return {
        "comments_relationship_count": len([rel for rel in reltypes if "comments" in rel]),
        "footnotes_relationship_count": len([rel for rel in reltypes if "footnotes" in rel]),
        "endnotes_relationship_count": len([rel for rel in reltypes if "endnotes" in rel]),
        "tracked_insertions": safe_xpath_count(body, ".//w:ins"),
        "tracked_deletions": safe_xpath_count(body, ".//w:del"),
        "field_simple_count": safe_xpath_count(body, ".//w:fldSimple"),
        "field_char_count": safe_xpath_count(body, ".//w:fldChar"),
        "hyperlink_count": safe_xpath_count(body, ".//w:hyperlink"),
        "bookmark_count": safe_xpath_count(body, ".//w:bookmarkStart"),
        "footnote_reference_count": safe_xpath_count(body, ".//w:footnoteReference"),
        "endnote_reference_count": safe_xpath_count(body, ".//w:endnoteReference"),
        "hidden_text_run_count": safe_xpath_count(body, ".//w:rPr/w:vanish"),
    }


def role_style_summary(
    items: list[ParagraphItem],
    roles: list[dict[str, Any]],
    preset: str,
) -> dict[str, Any]:
    preset_styles = PRESETS[preset]
    summary: dict[str, Any] = {}
    for item, role_info in zip(items, roles):
        paragraph = item.paragraph
        if paragraph is None:
            continue
        role = role_info["role"]
        style = compact_style(extract_paragraph_style(paragraph))
        bucket = summary.setdefault(role, {"count": 0, "style_variants": Counter(), "samples": []})
        bucket["count"] += 1
        bucket["style_variants"][style_signature(style)] += 1
        if len(bucket["samples"]) < 3:
            expected = compact_style(preset_styles.get(role, preset_styles["body"]))
            diffs = []
            for key, exp in expected.items():
                actual = style.get(key)
                if actual is not None and exp is not None and actual != exp:
                    diffs.append({"field": key, "actual": actual, "expected": exp})
            bucket["samples"].append(
                {
                    "paragraph_index": role_info["index"],
                    "length": role_info["length"],
                    "hash": role_info["hash"],
                    "style": style,
                    "differences_from_preset": diffs,
                }
            )

    normalized: dict[str, Any] = {}
    for role, data in summary.items():
        variants = [
            {"count": count, "style": json.loads(signature)}
            for signature, count in data["style_variants"].most_common()
        ]
        normalized[role] = {
            "count": data["count"],
            "style_variant_count": len(variants),
            "style_variants": variants[:5],
            "samples": data["samples"],
        }
    return normalized


def build_format_diagnostics(
    document: Any,
    items: list[ParagraphItem],
    roles: list[dict[str, Any]],
    preset: str,
) -> dict[str, Any]:
    role_summary = role_style_summary(items, roles, preset)
    warnings = []
    inconsistent_roles = [
        role for role, data in role_summary.items() if data.get("style_variant_count", 0) > 1
    ]
    if inconsistent_roles:
        warnings.append("Style variants detected within roles: " + ", ".join(sorted(inconsistent_roles)))
    if document.tables:
        warnings.append("Tables detected; review table borders, widths, and header rows manually if strict formatting is required.")
    if getattr(document, "inline_shapes", []):
        warnings.append("Images or inline objects detected; preserve seals/images unless the user explicitly requests repositioning.")
    special_state = special_state_diagnostics(document)
    if special_state.get("tracked_insertions") or special_state.get("tracked_deletions"):
        warnings.append("Tracked changes detected; preserve revisions unless the user explicitly requests accepting or rejecting them.")
    if special_state.get("field_simple_count") or special_state.get("field_char_count"):
        warnings.append("Field codes detected; verify generated tables of contents, dates, or references after formatting.")

    return {
        "page": page_diagnostics(document, preset),
        "headers_footers": header_footer_diagnostics(document),
        "role_style_summary": role_summary,
        "tables": table_diagnostics(document),
        "objects": object_diagnostics(document),
        "style_system": style_system_diagnostics(document),
        "special_state": special_state,
        "diagnostic_warnings": warnings,
        "diagnostic_scope": [
            "page setup",
            "headers and footers",
            "paragraph roles",
            "font family and size",
            "bold italic underline strikethrough color highlight",
            "indentation spacing alignment pagination controls tab stops numbering",
            "heading hierarchy",
            "tables",
            "inline images or objects",
            "style system",
            "comments tracked changes fields hyperlinks footnotes endnotes",
        ],
    }


def build_coverage(
    document: Any,
    roles: list[dict[str, Any]],
    counts: dict[str, int],
    mode: str,
    diagnostics: dict[str, Any] | None,
) -> dict[str, Any]:
    detected_roles = sorted(role for role in counts if role != "empty")
    formatted: list[dict[str, Any]] = []
    preserved: list[dict[str, Any]] = []
    diagnosed_only: list[dict[str, Any]] = []
    needs_review: list[dict[str, Any]] = []

    if mode == "diagnose-only":
        diagnosed_only.extend(
            [
                {"area": "page_setup", "status": "diagnosed_only"},
                {"area": "headers_footers", "status": "diagnosed_only"},
                {"area": "paragraph_roles", "status": "diagnosed_only", "roles": detected_roles},
                {"area": "typography_and_paragraph_layout", "status": "diagnosed_only"},
                {"area": "tables", "status": "diagnosed_only", "count": len(document.tables)},
                {
                    "area": "inline_images_or_objects",
                    "status": "diagnosed_only",
                    "count": len(getattr(document, "inline_shapes", [])),
                },
            ]
        )
    else:
        formatted.extend(
            [
                {"area": "page_setup", "status": "formatted"},
                {"area": "paragraph_roles", "status": "formatted", "roles": detected_roles},
                {"area": "typography_and_paragraph_layout", "status": "formatted"},
            ]
        )
        if document.tables:
            formatted.append(
                {
                    "area": "table_text",
                    "status": "formatted",
                    "count": len(document.tables),
                    "note": "Table paragraph text received body-style formatting where directly accessible.",
                }
            )
            preserved.append(
                {
                    "area": "table_structure",
                    "status": "preserved",
                    "count": len(document.tables),
                    "note": "Borders, widths, row heights, merged cells, and table layout are preserved unless explicitly handled.",
                }
            )

    inline_shape_count = len(getattr(document, "inline_shapes", []))
    if inline_shape_count:
        preserved.append(
            {
                "area": "inline_images_or_objects",
                "status": "preserved",
                "count": inline_shape_count,
                "note": "Images, seals, charts, or embedded inline objects are preserved unless the user explicitly asks for repositioning.",
            }
        )

    header_footer = diagnostics.get("headers_footers") if diagnostics else header_footer_diagnostics(document)
    header_footer_with_text = [
        section
        for section in header_footer.get("sections", [])
        if section.get("has_header_text") or section.get("has_footer_text")
    ]
    if header_footer_with_text and mode != "diagnose-only":
        preserved.append(
            {
                "area": "headers_footers",
                "status": "preserved",
                "section_count": len(header_footer_with_text),
                "note": "Existing header/footer text is preserved by default.",
            }
        )

    not_detected = [
        {"area": "paragraph_role", "status": "not_detected", "role": role}
        for role in COMMON_REPORT_ROLES
        if role not in counts
    ]

    for role_info in roles:
        if role_info.get("role") == "needs_review" or role_info.get("warnings"):
            needs_review.append(
                {
                    "area": "paragraph_role",
                    "status": "needs_review",
                    "paragraph_index": role_info.get("index"),
                    "role": role_info.get("role"),
                    "warnings": role_info.get("warnings", []),
                }
            )

    unsupported = list(LIMITED_FORMAT_AREAS)
    coverage = {
        "formatted": formatted,
        "preserved": preserved,
        "diagnosed_only": diagnosed_only,
        "not_detected": not_detected,
        "unsupported": unsupported,
        "needs_review": needs_review,
    }
    coverage["summary"] = {
        key: len(value)
        for key, value in coverage.items()
        if isinstance(value, list)
    }
    return coverage


def extract_template_profile(template_path: Path, preset: str) -> dict[str, Any]:
    document = Document(str(template_path))
    items = document_items(document)
    roles, counts = classify_items(items, preset)
    diagnostics = build_format_diagnostics(document, items, roles, preset)
    profile: dict[str, Any] = {"roles": {}, "page": {}, "role_counts": counts, "diagnostics": diagnostics}
    if document.sections:
        section = document.sections[0]
        profile["page"] = {
            "width_mm": mm_value(section.page_width),
            "height_mm": mm_value(section.page_height),
            "top_mm": mm_value(section.top_margin),
            "bottom_mm": mm_value(section.bottom_margin),
            "left_mm": mm_value(section.left_margin),
            "right_mm": mm_value(section.right_margin),
        }
        profile["page"] = {k: v for k, v in profile["page"].items() if v is not None}

    for item, role_info in zip(items, roles):
        role = role_info["role"]
        if role not in profile["roles"] and item.paragraph is not None:
            style = extract_paragraph_style(item.paragraph)
            if style:
                profile["roles"][role] = style
    return profile


def build_template_profile_report(
    *,
    template_path: Path,
    template_profile: dict[str, Any],
    preset: str,
    target_items: list[ParagraphItem] | None,
    target_roles: list[dict[str, Any]] | None,
    include_text: bool,
) -> dict[str, Any]:
    template_roles = set(template_profile.get("roles", {}).keys())
    target_role_counts: dict[str, int] = {}
    uncovered_target_roles: list[str] = []
    target_paragraphs: list[dict[str, Any]] = []

    if target_items is not None and target_roles is not None:
        for role_info in target_roles:
            role = role_info["role"]
            target_role_counts[role] = target_role_counts.get(role, 0) + 1
        uncovered_target_roles = sorted(set(target_role_counts) - template_roles)
        target_paragraphs = target_roles
        if include_text:
            enriched = []
            for role, item in zip(target_roles, target_items):
                row = dict(role)
                row["text"] = item.text
                enriched.append(row)
            target_paragraphs = enriched

    unresolved_items = []
    for role in uncovered_target_roles:
        unresolved_items.append(
            {
                "role": role,
                "reason": "Target document contains this role, but the template did not provide a matching style.",
                "recommended_options": [
                    "Use the selected preset fallback",
                    "Preserve the target document's existing formatting",
                    "Specify a custom style for this role",
                ],
            }
        )

    if target_items is None:
        unresolved_items.extend(
            [
                {
                    "role": "target roles",
                    "reason": "No target document was provided during template extraction.",
                    "recommended_options": [
                        "Provide a target document for coverage analysis",
                        "Proceed later and use preset fallback for uncovered roles",
                    ],
                },
                {
                    "role": "tables/images/page numbers",
                    "reason": "Template object styles can be listed, but target-specific treatment requires target inspection.",
                    "recommended_options": [
                        "Use official/internal-brief recommendation",
                        "Preserve target formatting",
                        "Specify custom handling",
                    ],
                },
            ]
        )

    return {
        "mode": "template-profile",
        "preset": preset,
        "template": str(template_path),
        "content_preservation": {
            "template_text_reported": bool(include_text),
            "target_text_reported": bool(include_text),
            "generated_missing_elements": False,
            "note": "Template extraction reports style fingerprints and omits full text unless include_text_in_report is enabled.",
        },
        "coverage": {
            "formatted": [],
            "preserved": [],
            "diagnosed_only": [
                {"area": "template_style_profile", "status": "diagnosed_only"},
                {"area": "target_role_coverage", "status": "diagnosed_only"} if target_items is not None else {"area": "target_role_coverage", "status": "not_detected"},
            ],
            "not_detected": [
                {"area": "target_document", "status": "not_detected"}
            ] if target_items is None else [],
            "unsupported": list(LIMITED_FORMAT_AREAS),
            "needs_review": unresolved_items,
        },
        "template_style_profile": {
            "page": template_profile.get("page", {}),
            "role_counts": template_profile.get("role_counts", {}),
            "role_styles": template_profile.get("roles", {}),
            "diagnostics": template_profile.get("diagnostics"),
        },
        "target_coverage": {
            "target_role_counts": target_role_counts,
            "uncovered_target_roles": uncovered_target_roles,
            "target_paragraphs": target_paragraphs,
        },
        "confirmation_required": bool(unresolved_items),
        "unresolved_items": unresolved_items,
        "recommended_next_step": (
            "Ask the user how to handle unresolved items before applying the template. "
            "If the user wants to proceed, use the selected preset as fallback and report it."
        ),
        "content_policy": "Template and target full text are omitted unless include_text_in_report is enabled.",
    }


def classify_items(items: list[ParagraphItem], preset: str) -> tuple[list[dict[str, Any]], dict[str, int]]:
    state = {"seen_separator": False, "article_title_assigned": False}
    roles: list[dict[str, Any]] = []
    counts: dict[str, int] = {}
    total = len(items)
    for idx, item in enumerate(items):
        role, warnings = classify_paragraph(item, idx, total, preset, state)
        counts[role] = counts.get(role, 0) + 1
        roles.append(
            {
                "index": idx + 1,
                "role": role,
                "length": len(item.text.strip()),
                "hash": text_hash(item.text),
                "warnings": warnings,
            }
        )
    return roles, counts


def merge_style(base: dict[str, Any], override: dict[str, Any] | None) -> dict[str, Any]:
    merged = dict(base)
    if override:
        merged.update({k: v for k, v in override.items() if v is not None})
    return merged


def format_document(
    document: Any,
    items: list[ParagraphItem],
    preset: str,
    template_profile: dict[str, Any] | None,
) -> tuple[list[dict[str, Any]], dict[str, int], list[str]]:
    preset_styles = PRESETS[preset]
    page_style = dict(preset_styles["_page"])
    if template_profile and template_profile.get("page"):
        page_style.update(template_profile["page"])
    apply_page_setup(document, page_style)

    roles, counts = classify_items(items, preset)
    template_fallback_roles: set[str] = set()
    for item, info in zip(items, roles):
        paragraph = item.paragraph
        if paragraph is None:
            continue
        role = info["role"]
        base_style = preset_styles.get(role, preset_styles["body"])
        template_style = None
        if template_profile:
            template_style = template_profile.get("roles", {}).get(role)
            if not template_style:
                template_fallback_roles.add(role)
        apply_style(paragraph, merge_style(base_style, template_style))

    for paragraph in iter_table_paragraphs(document):
        if paragraph.text.strip():
            apply_style(paragraph, preset_styles["body"])

    notes = []
    if template_fallback_roles:
        notes.append(
            "Template did not include matching styles for roles: "
            + ", ".join(sorted(template_fallback_roles))
            + f"; used the {preset} preset as the fallback."
        )
    return roles, counts, notes


def default_output_path(input_path: Path | None, input_name: str) -> Path:
    if input_path is not None:
        return input_path.with_name(input_path.stem + "_formatted.docx")
    return Path(input_name).with_suffix("").with_name(Path(input_name).stem + "_formatted.docx")


def build_report(
    *,
    source_name: str,
    preset: str,
    mode: str,
    roles: list[dict[str, Any]],
    counts: dict[str, int],
    output: Path | None,
    template_used: str | None,
    style_notes: list[str],
    diagnostics: dict[str, Any] | None,
    include_text: bool,
    items: list[ParagraphItem],
    content_preservation: dict[str, Any],
    coverage: dict[str, Any],
) -> dict[str, Any]:
    paragraph_reports = roles
    if include_text:
        paragraph_reports = []
        for role, item in zip(roles, items):
            enriched = dict(role)
            enriched["text"] = item.text
            paragraph_reports.append(enriched)

    warnings = []
    if counts.get("needs_review"):
        warnings.append("Some title-like paragraphs need human confirmation.")
    if template_used:
        warnings.append("Template replication used style fingerprints only; substantive template text was not reported.")
    warnings.extend(style_notes)
    if diagnostics:
        warnings.extend(diagnostics.get("diagnostic_warnings", []))
    return {
        "source": source_name,
        "mode": mode,
        "preset": preset,
        "template_used": template_used,
        "output": str(output) if output else None,
        "paragraph_count": len(items),
        "role_counts": counts,
        "content_preservation": content_preservation,
        "coverage": coverage,
        "paragraphs": paragraph_reports,
        "warnings": warnings,
        "style_notes": style_notes,
        "format_diagnostics": diagnostics,
        "content_policy": "Full paragraph text is omitted unless include_text_in_report is enabled.",
    }


def write_report(report: dict[str, Any], report_path: Path | None) -> None:
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if report_path:
        report_path.write_text(payload, encoding="utf-8")
    else:
        print(payload)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", help="Input .docx, .md, or .txt file")
    parser.add_argument("--stdin", action="store_true", help="Read Markdown/plain text from stdin")
    parser.add_argument("--input-name", default="stdin.txt", help="Virtual input name when using --stdin")
    parser.add_argument("--output", help="Output .docx path")
    parser.add_argument("--preset", choices=sorted(PRESETS.keys()), default="brief")
    parser.add_argument("--template", help="Optional .docx template to replicate")
    parser.add_argument("--extract-template", help="Only extract a .docx template style profile; do not format output")
    parser.add_argument("--target", help="Optional target file used with --extract-template for coverage analysis")
    parser.add_argument("--report", help="Optional JSON report path")
    parser.add_argument("--identify-only", action="store_true", help="Diagnose structure and formatting; do not save .docx")
    parser.add_argument("--diagnose-only", action="store_true", help="Alias for --identify-only")
    parser.add_argument("--include-text-in-report", action="store_true", help="Include full paragraph text in JSON report")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if not args.extract_template and not args.stdin and not args.input:
        raise SystemExit("Provide an input file or use --stdin.")

    require_docx()
    if args.extract_template:
        template_path = Path(args.extract_template).resolve()
        template_profile = extract_template_profile(template_path, args.preset)
        target_items = None
        target_roles = None
        if args.target:
            target_path = Path(args.target).resolve()
            target_suffix = target_path.suffix.lower()
            if target_suffix == ".docx":
                target_document = Document(str(target_path))
                target_items = document_items(target_document)
            elif target_suffix in {".md", ".markdown", ".txt", ""}:
                target_items, _ = read_plain_or_markdown(target_path, False, target_path.name)
            else:
                raise SystemExit(f"Unsupported target type: {target_suffix}. Use .docx, .md, .markdown, or .txt.")
            target_roles, _ = classify_items(target_items, args.preset)
        report = build_template_profile_report(
            template_path=template_path,
            template_profile=template_profile,
            preset=args.preset,
            target_items=target_items,
            target_roles=target_roles,
            include_text=args.include_text_in_report,
        )
        write_report(report, Path(args.report).resolve() if args.report else None)
        return 0

    input_path = Path(args.input).resolve() if args.input else None
    source_name = args.input_name if args.stdin else str(input_path)
    suffix = Path(args.input_name).suffix.lower() if args.stdin else input_path.suffix.lower()

    if suffix == ".docx" and input_path is not None:
        document = Document(str(input_path))
        items = document_items(document)
    elif suffix in {".md", ".markdown", ".txt", ""}:
        items, _ = read_plain_or_markdown(input_path, args.stdin, args.input_name)
        document = create_document_from_items(items)
    else:
        raise SystemExit(f"Unsupported input type: {suffix}. Use .docx, .md, .markdown, or .txt.")
    before_text_fingerprint = document_text_fingerprint(document)

    template_profile = None
    template_used = None
    if args.template:
        template_path = Path(args.template).resolve()
        template_profile = extract_template_profile(template_path, args.preset)
        template_used = str(template_path)

    diagnose_only = args.identify_only or args.diagnose_only
    output_path = None if diagnose_only else Path(args.output).resolve() if args.output else default_output_path(input_path, args.input_name).resolve()

    if diagnose_only:
        roles, counts = classify_items(items, args.preset)
        style_notes = []
        diagnostics = build_format_diagnostics(document, items, roles, args.preset)
        mode = "diagnose-only"
        after_text_fingerprint = document_text_fingerprint(document)
    else:
        roles, counts, style_notes = format_document(document, items, args.preset, template_profile)
        after_text_fingerprint = document_text_fingerprint(document)
        document.save(str(output_path))
        diagnostics = None
        mode = "template-replication" if template_profile else "format-only"

    content_preservation = compare_text_fingerprints(before_text_fingerprint, after_text_fingerprint)
    coverage = build_coverage(document, roles, counts, mode, diagnostics)

    report = build_report(
        source_name=source_name,
        preset=args.preset,
        mode=mode,
        roles=roles,
        counts=counts,
        output=output_path,
        template_used=template_used,
        style_notes=style_notes,
        diagnostics=diagnostics,
        include_text=args.include_text_in_report,
        items=items,
        content_preservation=content_preservation,
        coverage=coverage,
    )
    if args.report or diagnose_only:
        write_report(report, Path(args.report).resolve() if args.report else None)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
