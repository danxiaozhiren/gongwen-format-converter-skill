---
name: gongwen-format-converter
description: Format Chinese official documents, internal briefs, meeting minutes, reports, notices, requests, and work-information drafts into polished .docx deliverables. Use this skill whenever the user wants to convert or clean up 公文、公文格式、内部简报、信息简报、会议简报、会议纪要、通知、请示、报告、函, especially when they say the content should not be rewritten and they only need fonts, spacing, heading levels, margins, indentation, page setup, or a Word/Markdown/plain-text draft turned into a formatted official-looking document.
---

# Gongwen Format Converter

## Purpose

Turn existing Chinese administrative writing into a formatted deliverable without changing the user's content unless they explicitly request rewriting. Favor local processing, minimal disclosure, and reproducible .docx output.

Use this skill for:

- Formatting an existing `.docx` while preserving its text.
- Converting `.md`, `.txt`, or pasted text into a formatted `.docx`.
- Identifying paragraph roles before formatting.
- Replicating the layout of an internal template without exposing the template content.
- Producing a format-check report for 公文 or internal brief materials.

## Privacy Posture

Treat all user-provided 公文, meeting, production, operating, personnel, finance, and internal brief materials as confidential by default.

- Do not browse the web for user document contents.
- Do not quote internal source text in the response unless the user explicitly asks.
- When discussing template extraction, describe styles and structure only, not the template's substantive content.
- Prefer reports that show role counts, warnings, and missing fields instead of full paragraph text.
- If a document appears sensitive, tell the user you will preserve content locally and only surface formatting findings.

## Mode Selection

Ask only when the mode is unclear. Otherwise infer the mode from the user's request.

| Mode | Use When | Main Output |
| --- | --- | --- |
| Format-only | User says content is done, do not modify text, just adjust format | Formatted `.docx` and format report |
| Role-identification | Structure is messy or user wants you to check what each paragraph is | Paragraph role report, then formatted `.docx` if requested |
| Template-replication | User provides a sample/template document and a target document | Target `.docx` using the template's style fingerprint |

For all modes, keep the source order of paragraphs unless the user asks to reorganize.

## Input Handling

Accept these inputs:

- `.docx`: modify a copy of the document, preserving paragraph text and common run content.
- `.md`: parse Markdown headings and paragraphs, then generate `.docx`.
- `.txt`: split by non-empty lines and generate `.docx`.
- pasted text: save only if needed, then process as plain text.
- template `.docx` + target file/text: extract page and paragraph style patterns from the template, then apply them to the target.

The final deliverable should normally be `.docx`, because Markdown cannot express reliable fonts, page margins, fixed line spacing, or Word paragraph indentation.

## Formatting Presets

Read `references/format-presets.md` when choosing or explaining formatting details.

Default presets:

- `formal`: formal 公文-style page setup and hierarchy, suitable for 通知、请示、报告、函、纪要 drafts.
- `brief`: internal information brief style, suitable for 信息简报、会议简报、工作动态、生产经营分析会材料.

If the user provides a template, prefer template replication over generic presets.

## Paragraph Role Detection

Read `references/role-detection.md` before changing detection logic or when a document is ambiguous.

Common roles:

- `main_title`
- `subtitle`
- `issue_number`
- `metadata`
- `separator`
- `article_title`
- `recipient`
- `body`
- `heading_1`
- `heading_2`
- `heading_3`
- `attachment`
- `signature`
- `date`

When uncertain, mark the paragraph as `needs_review` in the report instead of inventing structure.

## Script Workflow

Use `scripts/format_document.py` for deterministic formatting.

Examples:

```bash
python scripts/format_document.py input.docx --output output.docx --preset formal --report report.json
python scripts/format_document.py draft.md --output formatted.docx --preset brief --report report.json
python scripts/format_document.py target.docx --template sample.docx --output styled.docx --report report.json
python scripts/format_document.py --stdin --input-name draft.md --output formatted.docx --preset brief
python scripts/format_document.py input.docx --identify-only --report roles.json
```

Behavior:

- Default to preserving text.
- Generate a `.docx` unless `--identify-only` is used.
- Generate a JSON report when `--report` is provided.
- Do not include full paragraph text in the report by default.
- Use `--include-text-in-report` only for non-sensitive samples or when the user explicitly asks.

## Deliverable Response

After processing, give the user:

- The output file path.
- The selected mode and preset/template basis.
- A short summary of role counts and important warnings.
- Any limitations, such as tables/images not fully restyled or font availability depending on the user's Word environment.

Keep the response concise and avoid exposing internal text.

## Important Boundaries

- Do not rewrite, polish, summarize, or add missing content in format-only mode.
- Do not fabricate issuing authority, document number, dates, policy basis, leader opinions, or attendance lists.
- Do not claim strict legal compliance with a standard; describe the output as following common 公文/内部简报 formatting practice unless the user provides an official unit template.
- If the document contains tables, images, or complex text boxes, preserve them where possible and disclose any formatting limitations.
