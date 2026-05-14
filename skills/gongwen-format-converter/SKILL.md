---
name: gongwen-format-converter
description: Format Chinese official documents, internal briefs, meeting minutes, reports, notices, requests, and work-information drafts into polished .docx deliverables. Use this skill whenever the user wants to convert or clean up 公文、公文格式、内部简报、信息简报、会议简报、会议纪要、通知、请示、报告、函, especially when they say the content should not be rewritten and they only need fonts, spacing, heading levels, margins, indentation, page setup, or a Word/Markdown/plain-text draft turned into a formatted official-looking document.
---

# Gongwen Format Converter

## Purpose

Turn existing Chinese administrative writing into a formatted deliverable without changing the user's content. Apply official-document, internal-brief, or template rules only to content and objects that already exist in the source. Favor local processing, minimal disclosure, and reproducible .docx output.

Use this skill for:

- Formatting an existing `.docx` while preserving its text.
- Converting `.md`, `.txt`, or pasted text into a formatted `.docx`.
- Diagnosing the whole document's structure and formatting before making changes.
- Replicating the layout of an internal template without exposing the template content.
- Producing a format-check report for 公文 or internal brief materials.

## Content Zero-Change Rule

This rule has priority over all formatting presets and official-document references.

- Preserve existing text, paragraph order, tables, images, headers, footers, and document objects.
- Do not rewrite, polish, summarize, add, delete, reorder, or infer substantive content.
- Never create missing official-document elements such as 主送机关, 发文字号, 发文机关, 签发人, 附件名称, 署名, 日期, 印章, 抄送机关, 印发机关, 版记, or 页码 unless the user explicitly asks for that separate content/layout addition.
- If an element is absent, report it as `not_detected` / `not_processed`; do not describe it as something the formatter filled or should fill during format-only work.
- Use official and template rules to decide how existing elements should look, not to invent elements that are not present.

## Privacy Posture

Treat all user-provided 公文, meeting, production, operating, personnel, finance, and internal brief materials as confidential by default.

- Do not browse the web for user document contents.
- Do not quote internal source text in the response unless the user explicitly asks.
- When discussing template extraction, describe styles and structure only, not the template's substantive content.
- Prefer reports that show role counts, coverage status, warnings, and hashes instead of full paragraph text.
- If a document appears sensitive, tell the user you will preserve content locally and only surface formatting findings.

## Mode Selection

If the user provides only an article/file and the requested operation is unclear, stop and ask them to choose a mode before processing. This matters because formatting, structure diagnosis, and template replication make different assumptions about what may be changed.

Use this short clarification:

```text
这份材料我只处理已有内容的格式，不补写任何缺失公文要素。你想选哪一种？
1. 现文格式化：内容不改，只调整已有文字、段落、页面、表格、图片、页眉页脚等格式。
2. 现文格式诊断：不改文档，完整识别已有内容的页面、段落、文字、样式、表格、图片和对象格式状态。
3. 现文模板套用：只学习模板样式并套到目标文档已有内容，不复制模板正文，也不补目标缺失内容。

如果你不确定，我建议选 1「现文格式化」。
```

Infer the mode only when the user's wording is explicit:

- "内容不要改/只调格式/套公文格式" means existing-content `Format-only`.
- "先看看格式/诊断格式/识别段落和样式/哪些不规范" means existing-content `Format-diagnostics`.
- "参考这份模板/复刻格式/套成同款" plus a template file means existing-content `Template-replication`.

| Mode | Use When | Main Output |
| --- | --- | --- |
| Format-only | User says content is done, do not modify text, just adjust existing formatting | Formatted `.docx` and coverage report |
| Format-diagnostics | User wants to inspect existing structure and formatting before editing, or asks what is nonstandard | Full format diagnostic report; no document changes |
| Template-replication | User provides a sample/template document and a target document | Template style profile first, then confirmed target `.docx` using existing target content only |

For all modes, keep the source order of paragraphs and do not add missing content unless the user explicitly asks for that separate task.

## Input Handling

Accept these inputs:

- `.docx`: modify a copy of the document, preserving paragraph text and common run content.
- `.md`: parse Markdown headings and paragraphs, treating Markdown markers as formatting signals, then generate `.docx`.
- `.txt`: split by non-empty lines and generate `.docx`.
- pasted text: save only if needed, then process as plain text.
- template `.docx` + target file/text: extract page and paragraph style patterns from the template, then apply them to the target.

The final deliverable should normally be `.docx`, because Markdown cannot express reliable fonts, page margins, fixed line spacing, or Word paragraph indentation.

## Formatting Presets

Read `references/official-format-scope.md` first for formal 公文 work or whenever the user asks for "strict", "official", "规范", "GB/T 9704", or "党政机关公文格式". This reference defines the official format scope and the priority order for user instructions, official templates, presets, and fallback decisions.

Read `references/format-presets.md` when choosing or explaining formatting details.

Default presets:

- `formal`: formal 公文-style page setup and hierarchy, suitable for 通知、请示、报告、函、纪要 drafts.
- `brief`: internal information brief style, suitable for 信息简报、会议简报、工作动态、生产经营分析会材料.

If the user provides a template, prefer template replication over generic presets.

## Format Plan Confirmation

Before changing a document in format-only mode, provide a concise format plan unless the user has already provided exact styles and asked you to proceed immediately. The plan should list the specific formatting range, not just say "format the document." Make clear that the formatter only applies rules to existing content.

Include:

- format source: explicit user rules, user template, `formal` official-document preset, or `brief` internal-brief preset;
- page setup: paper, margins, orientation, page-number policy;
- official/document elements already present: title, recipient, body, attachments, signature/date, notes, copy/printing area when detected;
- typography: fonts, sizes, bold/italic/underline/color for title, body, and heading levels;
- paragraph layout: first-line indent, left/right indent, line spacing, paragraph spacing, alignment;
- hierarchy and numbering: `一、`, `（一）`, `1.`, `（1）`;
- objects and whole-document formatting: tables, images, seals, headers/footers, page numbers, separators, text boxes, shapes, styles, comments/revisions when detectable;
- coverage status: `formatted`, `preserved`, `diagnosed_only`, `not_detected`, `unsupported`, and `needs_review`.

Use this priority:

1. User-specified formatting overrides everything.
2. A supplied official/unit template overrides built-in presets.
3. Formal 公文 uses the official-scope/formal preset.
4. Internal briefs use the brief preset but must be described as internal-office convention, not official GB/T compliance.
5. Ambiguous existing items require a question when formatting would be risky; absent content elements are reported as `not_detected`, not filled. If the user says to proceed, use the closest preset for existing content and report the fallback.

## Full Format Scope

Treat "format" as the full Word presentation and layout surface, not only title/body fonts.

- Page: paper size, orientation, margins, gutter, sections, columns, document grid, page-number areas, header/footer distances.
- Paragraph: alignment, first-line/left/right indentation, line spacing, before/after spacing, outline level, pagination controls, tab stops, bullets, and numbering.
- Run/text: East Asian and Latin fonts, size, color, bold, italic, underline, strikethrough, emphasis marks, superscript/subscript, character spacing, highlight, and language attributes when available.
- Existing structural roles: copy number, secrecy/urgency labels, issuer mark, document number, signer, title, subtitle, issue number, metadata, separator, article title, recipient, body, headings, attachment note, note, signature, date, cc, printing area, and version-record text.
- Tables: width, column width, row height, cell margins, borders, shading, cell alignment, header rows, table text, and cross-page behavior when safely detectable.
- Objects: images, seal images, text boxes, shapes, lines, charts, formulas, embedded objects, watermarks, and object anchors when safely detectable.
- Headers/footers and version-record areas: header text/lines, footer text, page numbers, separator lines, cc/printing area formatting, and first-page/odd-even differences.
- Style system: paragraph styles, character styles, heading styles, direct formatting conflicts, and style variants.
- Special document state: comments, tracked changes, fields, table of contents, cross-references, footnotes/endnotes, hyperlinks, bookmarks, and hidden text.

For each area, automatically format only safe existing content. Preserve and report anything risky, unsupported, or not reliably detectable.

## Template Replication Rules

Template replication is a two-stage workflow. Do not apply a template immediately when important style coverage is unclear.

Stage 1: extract and confirm the template profile.

- Extract the template style fingerprint.
- List the template's page setup, role styles, object/table diagnostics, and covered roles.
- If a target document is provided, compare target roles with template roles and list uncovered existing target styles.
- Ask the user how to handle unresolved style coverage before applying the template.

Stage 2: apply the confirmed template profile.

- Apply the extracted template styles to matching target roles.
- Use the user's explicit choices for unresolved items.
- If the user tells you to proceed without answering, use the selected `formal` or `brief` preset as fallback and report every fallback.

Extract and apply the template's page and paragraph style fingerprint, including:

- page margins;
- alignment;
- first-line, left, and right indentation;
- line spacing and paragraph spacing;
- font name and font size;
- bold, italic, underline;
- font color when explicitly set;
- representative role styles such as title, issue number, metadata, article title, body, and numbered headings.

If the target document contains an existing role that has no matching style in the template, do not silently invent a template style. Ask the user whether to:

1. use the recommended 公文/内部简报 preset for that uncovered role;
2. preserve the target document's existing formatting for that role;
3. specify a custom style.

If the user wants you to continue without confirmation, choose the relevant preset (`formal` or `brief`) as the fallback and disclose this in the report.

For decorative or ambiguous elements such as underlines, colored text, separators, or unusual spacing:

- preserve them when the template clearly uses them for the same role;
- remove/normalize them when using the standard 公文/内部简报 fallback, because official-looking documents should avoid decorative formatting unless the unit template requires it;
- ask before making a one-off judgment when the user's intent is unclear.

Use this confirmation prompt after Stage 1:

```text
我已提取模板格式清单。模板覆盖了：[列出角色/页面/对象范围]。
目标文档中以下已有格式项模板未覆盖：[列出未覆盖角色或对象]。

请选择未覆盖项处理方式：
1. 使用公文/内部简报推荐格式；
2. 保留目标文档原格式；
3. 我指定自定义格式。

确认后我再生成套用模板后的 Word。
```

## Format Diagnostics

Read `references/official-format-scope.md` and `references/role-detection.md` before diagnosing formal documents. Diagnostics are broader than role detection: they should inspect the whole document's current formatting without modifying it.

Report these areas when available:

- page setup: paper size, orientation, margins, gutter, sections, columns, document grid where detectable, and differences from the selected preset;
- headers and footers: whether header/footer text exists, first-page/odd-even differences, and whether existing page-number handling needs confirmation;
- official/internal structure already present: likely 版头/主体/版记 or internal brief header components, reported as detected or not detected without filling absent content;
- paragraph roles: title, issue number, metadata, article title, recipient, body, headings, attachment, note, signature, date, cc, printing/version-record text when detectable;
- typography: East Asian/Latin font name, size, bold, italic, underline, strikethrough, color, highlight, and other run-level variants when available;
- paragraph layout: first-line indent, left/right indent, line spacing, paragraph spacing, alignment, outline level, tab stops, numbering/bullets, and pagination controls when available;
- hierarchy consistency: `一、`, `（一）`, `1.`, `（1）`;
- objects: tables, images, seals, text boxes, shapes, watermarks, formulas, charts, separators, fields, comments, tracked changes, hyperlinks, footnotes/endnotes where detectable;
- style system: style names, direct formatting, style variants, and conflicts;
- coverage: `formatted`, `preserved`, `diagnosed_only`, `not_detected`, `unsupported`, and `needs_review`;
- consistency: roles that have multiple style variants;
- differences from `formal` or `brief` preset;
- unclear existing items that need user confirmation.

Do not quote full paragraph text by default. Use paragraph indexes, lengths, hashes, role counts, and style summaries.

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

If the local Python environment does not have `python-docx`, install the script dependency first:

```bash
python -m pip install -r scripts/requirements.txt
```

Examples:

```bash
python scripts/format_document.py input.docx --output output.docx --preset formal --report report.json
python scripts/format_document.py draft.md --output formatted.docx --preset brief --report report.json
python scripts/format_document.py --extract-template sample.docx --target target.docx --preset brief --report template_profile.json
python scripts/format_document.py target.docx --template sample.docx --output styled.docx --report report.json
python scripts/format_document.py --stdin --input-name draft.md --output formatted.docx --preset brief
python scripts/format_document.py input.docx --diagnose-only --report diagnostics.json
python scripts/format_document.py input.docx --identify-only --report diagnostics.json
```

Behavior:

- Default to preserving text.
- Report `content_preservation` and `coverage` when possible.
- Generate a `.docx` unless `--diagnose-only` / `--identify-only` is used.
- Generate a JSON report when `--report` is provided.
- Do not include full paragraph text in the report by default.
- Use `--include-text-in-report` only for non-sensitive samples or when the user explicitly asks.
- In template mode, report any roles that fell back to a preset because the template did not contain a matching style.
- In template replication work, prefer `--extract-template` first so the user can review covered styles and unresolved items before applying.

## Deliverable Response

After processing, give the user:

- The output file path.
- The selected mode and preset/template basis.
- A short coverage summary: what was formatted, preserved, diagnosed only, not detected, unsupported, and needs review.
- A short summary of role counts and important warnings.
- Any limitations, such as tables/images not fully restyled or font availability depending on the user's Word environment.

Keep the response concise and avoid exposing internal text.

## Important Boundaries

- Do not rewrite, polish, summarize, delete, reorder, or add missing content in any mode.
- Do not fabricate issuing authority, document number, dates, policy basis, leader opinions, attendance lists, attachments, recipients, signatures, seals, cc lines, printing areas, version records, or page numbers.
- Do not claim strict legal compliance with a standard; describe the output as following common 公文/内部简报 formatting practice unless the user provides an official unit template.
- If the document contains tables, images, complex text boxes, comments, tracked changes, fields, or embedded objects, preserve them where possible and disclose any formatting limitations.
