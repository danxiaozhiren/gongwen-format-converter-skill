"""Report helpers for the 行文检查版式 formatter."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .presets import PRESETS
from .roles import ParagraphItem


def document_grid_reference(preset: str) -> dict[str, Any] | None:
    page = PRESETS[preset].get("_page", {})
    grid_keys = [
        "grid_line_count",
        "grid_char_count",
        "line_pitch_twips",
        "char_space_twips",
    ]
    if not any(page.get(key) is not None for key in grid_keys):
        return None
    return {
        "line_count": page.get("grid_line_count"),
        "char_count": page.get("grid_char_count"),
        "grid_type": page.get("grid_type"),
        "line_pitch_twips": page.get("line_pitch_twips"),
        "char_space_twips": page.get("char_space_twips"),
        "status": "reference_grid",
        "note": (
            "The formatter writes Word docGrid linePitch/charSpace when available. "
            "Line and character counts are layout references, not a guaranteed rendered page count; "
            "verify strict 22-line/28-character appearance in Word/WPS when required."
        ),
    }


def coverage_area_names(entries: list[dict[str, Any]]) -> list[str]:
    names = []
    for entry in entries:
        area = entry.get("area") or entry.get("role")
        if area and area not in names:
            names.append(str(area))
    return names


def content_status(content_preservation: dict[str, Any]) -> str:
    changed = bool(
        content_preservation.get("text_changed")
        or content_preservation.get("text_unit_count_changed")
        or content_preservation.get("paragraph_order_changed")
    )
    if not changed:
        return "preserved"
    if content_preservation.get("generated_missing_elements"):
        return "changed_by_explicit_layout_addition"
    return "changed"


def report_format_source(preset: str) -> str:
    if preset == "checklist":
        return "xingwen-checklist 行文出手前对照检查事项"
    return preset


def build_user_message(
    *,
    status: str,
    mode: str,
    output: Path | None,
    requested_output: Path | None,
    output_withheld: bool,
    preservation_status: str,
    needs_review_count: int,
    unsupported_areas: list[str],
) -> dict[str, Any]:
    can_do = "我能做已有 .docx 的行文检查版式整理或格式诊断，并保持正文不改。"
    next_steps = ["严格 22 行/28 字视觉效果需要在 Word/WPS 中最终确认。"]
    if unsupported_areas:
        next_steps.append("报告中列出的 unsupported areas 需要人工查看，脚本不会自动修改这些对象。")
    if needs_review_count:
        next_steps.append("请先核对 needs_review 段落角色，再把输出作为最终稿使用。")

    if output_withheld or status == "failed_content_preservation":
        return {
            "title": "已停止交付：内容保持校验失败",
            "body": "脚本检测到候选文件文本与原文不一致，已删除候选文件，未生成或覆盖正式输出。",
            "can_do": can_do,
            "current_state": "候选文件已拦截，正式输出未交付。",
            "next_action": "先查看报告里的 content_preservation.changed_unit_indexes，确认变化来源后再重新运行。",
            "output": None,
            "requested_output": str(requested_output) if requested_output else None,
            "next_steps": [
                "查看报告里的 content_preservation.changed_unit_indexes 定位变化来源。",
                "确认脚本不会改正文后再重新运行。",
            ],
        }

    if mode == "diagnose-only":
        return {
            "title": "已完成格式诊断",
            "body": "本次只生成诊断报告，没有生成新的 Word 文件。",
            "can_do": can_do,
            "current_state": "诊断报告已生成，Word 文件未被修改。",
            "next_action": "根据报告确认是否继续执行现文格式化。",
            "output": None,
            "requested_output": None,
            "next_steps": next_steps,
        }

    if status == "needs_review":
        return {
            "title": "已完成版式整理，但有段落需要确认",
            "body": "输出文件已生成，正文文本保持校验通过；部分段落角色置信度较低，建议人工核对。",
            "can_do": can_do,
            "current_state": "候选文件已通过内容保持校验并交付，但有段落角色需要人工确认。",
            "next_action": "先核对 needs_review 段落，再在 Word/WPS 中做最终视觉检查。",
            "output": str(output) if output else None,
            "requested_output": str(requested_output) if requested_output else None,
            "next_steps": next_steps,
        }

    body = "输出文件已生成，正文文本保持校验通过。"
    if preservation_status == "changed_by_explicit_layout_addition":
        body = "输出文件已生成；文本变化仅来自用户明确要求生成的版式元素。"
    return {
        "title": "已完成行文检查版式整理",
        "body": body,
        "can_do": can_do,
        "current_state": "输出文件已生成，并已通过内容保持校验。",
        "next_action": "打开输出文件做最终视觉确认，尤其检查 22 行/28 字和复杂对象区域。",
        "output": str(output) if output else None,
        "requested_output": str(requested_output) if requested_output else None,
        "next_steps": next_steps,
    }


def build_agent_summary(
    *,
    preset: str,
    mode: str,
    output: Path | None,
    requested_output: Path | None,
    output_withheld: bool,
    content_preservation: dict[str, Any],
    coverage: dict[str, Any],
    warnings: list[str],
) -> dict[str, Any]:
    preservation_status = content_status(content_preservation)
    needs_review = coverage.get("needs_review", [])
    unsupported = coverage.get("unsupported", [])
    recommended_actions: list[str] = []

    if preservation_status == "changed":
        status = "failed_content_preservation"
        recommended_actions.append("Stop before delivering the formatted file; compare source and output text units.")
    elif needs_review:
        status = "needs_review"
        recommended_actions.append("Confirm paragraphs marked needs_review before relying on their role formatting.")
    else:
        status = "ok"

    if preservation_status == "changed_by_explicit_layout_addition":
        recommended_actions.append("Content hash changed only with explicit generated layout elements; review generated_layout_elements.")
    if output_withheld:
        recommended_actions.append("Output was withheld because content preservation failed; rerun only after inspecting changed_unit_indexes.")
    if unsupported:
        recommended_actions.append(
            "Note the unsupported automatic scope; manually review these areas only when the source contains such complex Word objects."
        )
    grid_reference = document_grid_reference(preset)
    if grid_reference:
        recommended_actions.append("Verify strict line/character grid appearance in Word/WPS when 22 lines and 28 characters per line matter.")
    if mode == "diagnose-only":
        recommended_actions.append("Use the diagnostic report to confirm whether to run format-only before generating a .docx.")
    if not recommended_actions:
        recommended_actions.append("Review the coverage summary and open the output in Word/WPS for final visual confirmation.")
    unsupported_areas = coverage_area_names(unsupported)
    user_message = build_user_message(
        status=status,
        mode=mode,
        output=output,
        requested_output=requested_output,
        output_withheld=output_withheld,
        preservation_status=preservation_status,
        needs_review_count=len(needs_review),
        unsupported_areas=unsupported_areas,
    )

    return {
        "status": status,
        "mode": mode,
        "preset": preset,
        "format_source": report_format_source(preset),
        "output": str(output) if output else None,
        "requested_output": str(requested_output) if requested_output else None,
        "output_withheld": output_withheld,
        "content_status": preservation_status,
        "formatted_areas": coverage_area_names(coverage.get("formatted", [])),
        "preserved_areas": coverage_area_names(coverage.get("preserved", [])),
        "diagnosed_only_areas": coverage_area_names(coverage.get("diagnosed_only", [])),
        "needs_review_count": len(needs_review),
        "unsupported_areas": unsupported_areas,
        "not_detected_count": len(coverage.get("not_detected", [])),
        "document_grid_reference": grid_reference,
        "warning_count": len(warnings),
        "recommended_actions": recommended_actions,
        "user_message": user_message,
    }


def build_report(
    *,
    source_name: str,
    preset: str,
    mode: str,
    roles: list[dict[str, Any]],
    counts: dict[str, int],
    output: Path | None,
    requested_output: Path | None,
    output_withheld: bool,
    style_notes: list[str],
    diagnostics: dict[str, Any] | None,
    include_text: bool,
    items: list[ParagraphItem],
    content_preservation: dict[str, Any],
    coverage: dict[str, Any],
    format_changes: dict[str, Any],
    format_actions: dict[str, Any],
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
    warnings.extend(style_notes)
    if diagnostics:
        warnings.extend(diagnostics.get("diagnostic_warnings", []))
    summary = build_agent_summary(
        preset=preset,
        mode=mode,
        output=output,
        requested_output=requested_output,
        output_withheld=output_withheld,
        content_preservation=content_preservation,
        coverage=coverage,
        warnings=warnings,
    )
    return {
        "source": source_name,
        "mode": mode,
        "preset": preset,
        "output": str(output) if output else None,
        "requested_output": str(requested_output) if requested_output else None,
        "output_withheld": output_withheld,
        "summary": summary,
        "paragraph_count": len(items),
        "role_counts": counts,
        "content_preservation": content_preservation,
        "format_changes": format_changes,
        "format_actions": format_actions,
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
