#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Format Chinese official-document and internal-brief drafts as .docx files.

The script is intentionally conservative: it preserves paragraph text, applies
role-based formatting, and keeps reports free of full paragraph content by
default.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor


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
        return None


def text_hash(text: str) -> str:
    return hashlib.sha256(text.strip().encode("utf-8")).hexdigest()[:12]


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


ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


PRESETS: dict[str, dict[str, dict[str, Any]]] = {
    "formal": {
        "_page": {"top_mm": 37, "bottom_mm": 35, "left_mm": 28, "right_mm": 26},
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
        "_page": {"top_mm": 25, "bottom_mm": 25, "left_mm": 28, "right_mm": 26},
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


def extract_paragraph_style(paragraph: Any) -> dict[str, Any]:
    fmt = paragraph.paragraph_format
    style: dict[str, Any] = {}
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

    for run in paragraph.runs:
        if run.text.strip():
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
            if run.font.color is not None and run.font.color.rgb is not None:
                style["color"] = str(run.font.color.rgb)
            break
    return {k: v for k, v in style.items() if v is not None}


def extract_template_profile(template_path: Path, preset: str) -> dict[str, Any]:
    document = Document(str(template_path))
    items = document_items(document)
    roles, _ = classify_items(items, preset)
    profile: dict[str, Any] = {"roles": {}, "page": {}}
    if document.sections:
        section = document.sections[0]
        profile["page"] = {
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
    include_text: bool,
    items: list[ParagraphItem],
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
    return {
        "source": source_name,
        "mode": mode,
        "preset": preset,
        "template_used": template_used,
        "output": str(output) if output else None,
        "paragraph_count": len(items),
        "role_counts": counts,
        "paragraphs": paragraph_reports,
        "warnings": warnings,
        "style_notes": style_notes,
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
    parser.add_argument("--report", help="Optional JSON report path")
    parser.add_argument("--identify-only", action="store_true", help="Only classify paragraph roles; do not save .docx")
    parser.add_argument("--include-text-in-report", action="store_true", help="Include full paragraph text in JSON report")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.stdin and not args.input:
        raise SystemExit("Provide an input file or use --stdin.")

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

    template_profile = None
    template_used = None
    if args.template:
        template_path = Path(args.template).resolve()
        template_profile = extract_template_profile(template_path, args.preset)
        template_used = str(template_path)

    output_path = None if args.identify_only else Path(args.output).resolve() if args.output else default_output_path(input_path, args.input_name).resolve()

    if args.identify_only:
        roles, counts = classify_items(items, args.preset)
        style_notes = []
        mode = "identify-only"
    else:
        roles, counts, style_notes = format_document(document, items, args.preset, template_profile)
        document.save(str(output_path))
        mode = "template-replication" if template_profile else "format-only"

    report = build_report(
        source_name=source_name,
        preset=args.preset,
        mode=mode,
        roles=roles,
        counts=counts,
        output=output_path,
        template_used=template_used,
        style_notes=style_notes,
        include_text=args.include_text_in_report,
        items=items,
    )
    if args.report or args.identify_only:
        write_report(report, Path(args.report).resolve() if args.report else None)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
