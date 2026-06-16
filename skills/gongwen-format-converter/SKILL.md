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

## Agent Quick Path

Follow this path before reading deeper details:

1. Determine the mode: `Format-only`, `Format-diagnostics`, or `Template-replication`. If the user only says "process this" or gives a file without intent, ask the clarification in Mode Selection.
2. State the zero-change boundary: preserve existing content and objects; do not add missing official-document elements unless explicitly requested.
3. Choose the format source: explicit user rules, supplied template, `formal`, or `brief`.
4. Read `references/official-format-scope.md` for formal/strict/GB/T work, `references/format-presets.md` for preset details, and `references/role-detection.md` when role classification is ambiguous.
5. Before modifying a document in `Format-only`, give a short format plan unless the user already gave exact styles and told you to proceed.
6. Run `scripts/format_document.py` with `--report` whenever possible. Prefer reports over quoting document text.
7. For diagnostics, use `--diagnose-only` and do not create a formatted copy.
8. For template replication, run `--extract-template` first, review uncovered roles with the user, then apply `--template` only after confirmation or explicit instruction to proceed.
9. Inspect `content_preservation`, `coverage.summary`, `role_counts`, `warnings`, `format_changes.page`, and `format_changes.paragraph_controls` before responding.
10. Deliver the output path, selected mode/preset/template basis, coverage summary, important warnings, and limitations.

## Interaction Protocol

Keep the user oriented before every meaningful step. Use short status blocks instead of long explanations.

Before any document-changing action, say:

```text
当前步骤：格式计划确认
我将做：按 [默认行文检查表 / 用户模板 / 用户指定规则] 处理已有内容的页面、字体、行距、标题层级、表格文字、页眉页脚和页码等安全格式。
不会做：不改正文，不补写缺失公文要素，不移动或重排复杂对象，除非你明确要求。
产出：格式化后的 .docx 和覆盖报告。
下一步：你确认后我开始生成。
```

Before diagnostics, say:

```text
当前步骤：格式诊断
我将做：只读取文档结构和格式，生成页面、段落、字体、表格、对象、页眉页脚、字段和特殊状态的诊断报告。
不会做：不保存格式化副本，不改正文，不接受或拒绝修订。
产出：诊断 JSON 或摘要报告。
下一步：诊断完成后我会说明哪些可自动处理、哪些需要人工确认。
```

Before template replication, say:

```text
当前步骤：模板覆盖分析
我将做：只提取模板的页面、字体、字号、行距、缩进、标题层级和对象覆盖情况，不复制模板正文。
不会做：不会直接套用模板到目标文档，除非覆盖情况确认或你明确要求继续。
产出：模板样式指纹、目标文档未覆盖项和处理建议。
下一步：你确认未覆盖项处理方式后，我再生成套用模板的 Word。
```

While working, provide concise progress updates at these checkpoints:

- mode selected and format source chosen;
- command about to run, with output/report paths;
- report read completed, including `content_preservation` result;
- final file ready, including coverage summary and remaining limitations.

If a step fails, say the failed step, the reason, what was not changed, and the next safe option. Do not continue into a riskier action silently.

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
2. 现文格式诊断：不改文档，尽可能识别已有内容的页面、段落、文字、样式、表格、图片和对象格式状态，并说明覆盖边界。
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

- `formal`: default 行文出手前检查表 template for formal 公文-style drafts, suitable for 通知、请示、报告、函、纪要 drafts.
- `brief`: internal information brief style, suitable for 信息简报、会议简报、工作动态、生产经营分析会材料.

The built-in `formal` default uses:

- page: A4; margins top 37 mm, bottom 35 mm, left 27 mm, right 27 mm; 22 lines/page and 28 characters/line reference;
- main title: 方正小标宋简体, 2号, not bold, centered; multi-line title spacing 36 pt and balanced trapezoid/diamond-like line breaks when safely controllable;
- title/body gap: one 3号-line equivalent, expressed as spacing rather than inserted content when possible;
- body: 仿宋_GB2312, 3号, not bold, fixed line spacing 30 pt, first-line indent 2 Chinese characters;
- headings: `一、` uses 黑体 3号 not bold; `（一）` uses 楷体_GB2312 3号 bold; `1.` and `（1）` use 仿宋_GB2312 3号 not bold;
- digits and Latin letters: Times New Roman; page numbers: 宋体 4号.

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
- objects and whole-document coverage: tables, images, seals, headers/footers, page numbers, separators, text boxes, shapes, styles, comments/revisions when detectable;
- coverage status: `formatted`, `preserved`, `diagnosed_only`, `not_detected`, `unsupported`, and `needs_review`.

Use this priority:

1. User-specified formatting overrides everything.
2. A supplied official/unit template overrides built-in presets.
3. Formal 公文 uses the official-scope/formal preset.
4. Internal briefs use the brief preset but must be described as internal-office convention, not official GB/T compliance.
5. Ambiguous existing items require a question when formatting would be risky; absent content elements are reported as `not_detected`, not filled. If the user says to proceed, use the closest preset for existing content and report the fallback.

## Word Format Surface and Safety Boundary

Treat "format" as the full Word presentation and layout surface for diagnosis and coverage reporting, not only title/body fonts. Do not claim complete Word-format handling. Automatically modify only safe existing content; preserve or report objects and states that cannot be reliably edited.

- Page: paper size, orientation, margins, gutter, sections, columns, document grid, text-area reference, existing page-number areas, header/footer distances, and page-level before/after changes.
- Paragraph: alignment, first-line/left/right indentation, line spacing, before/after spacing, outline level, pagination controls, tab stops, bullets, and numbering. Treat literal markers like `一、`, `（一）`, and `1.` as existing text unless the user explicitly asks for Word automatic numbering.
- Run/text: East Asian and Latin fonts, size, color, bold, italic, underline, strikethrough, emphasis marks, superscript/subscript, character spacing, highlight, and language attributes when available.
- Existing structural roles: copy number, secrecy/urgency labels, issuer mark, document number, signer, title, subtitle, issue number, metadata, separator, article title, recipient, body, headings, attachment note, note, signature, date, cc, printing area, and version-record text.
- Tables: width, column width, row height, cell margins, borders, shading, cell alignment, header rows, table text, and cross-page behavior when safely detectable.
- Objects: images, seal images, text boxes, shapes, lines, charts, formulas, embedded objects, watermarks, and object anchors when safely detectable.
- Headers/footers and version-record areas: header text/lines, footer text, page numbers, separator lines, cc/printing area formatting, and first-page/odd-even differences.
- Style system: paragraph styles, character styles, heading styles, direct formatting conflicts, and style variants.
- Special document state: comments, tracked changes, fields, table of contents, cross-references, footnotes/endnotes, hyperlinks, bookmarks, and hidden text.

For each area, automatically format only safe existing content. Preserve and report anything risky, unsupported, or not reliably detectable. If asked whether the skill can fully process Word formatting, answer that it handles common 公文/内部简报 formatting and coverage diagnostics, but not complete Word object/style/rendering equivalence.

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
- paragraph layout: first-line indent, left/right indent, line spacing, paragraph spacing, alignment, outline level, tab stops, numbering/bullets, pagination controls, and `format_changes.paragraph_controls` when available;
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
- `heading_4`
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

Run commands from the skill folder. If working from a repository root, prefix paths with `skills/gongwen-format-converter/`.

Command selection:

| User intent | Command pattern |
| --- | --- |
| Format an existing `.docx` with formal rules | `python scripts/format_document.py input.docx --output output.docx --preset formal --report report.json` |
| Format Markdown/text as an internal brief | `python scripts/format_document.py draft.md --output formatted.docx --preset brief --report report.json` |
| Diagnose only, no document changes | `python scripts/format_document.py input.docx --diagnose-only --preset formal --report diagnostics.json` |
| Analyze a template before applying it | `python scripts/format_document.py --extract-template sample.docx --target target.docx --preset brief --report template_profile.json` |
| Apply a confirmed template profile | `python scripts/format_document.py target.docx --template sample.docx --output styled.docx --preset brief --report report.json` |
| Add missing page numbers by explicit request | `python scripts/format_document.py input.docx --output numbered.docx --preset formal --add-page-numbers --report report.json` |
| Normalize table structure by explicit request | `python scripts/format_document.py input.docx --output table-normalized.docx --preset formal --format-tables --report report.json` |
| Include full text in a report by explicit request | Add `--include-text-in-report`; otherwise omit it for privacy. |

Examples:

```bash
python scripts/format_document.py input.docx --output output.docx --preset formal --report report.json
python scripts/format_document.py draft.md --output formatted.docx --preset brief --report report.json
python scripts/format_document.py --extract-template sample.docx --target target.docx --preset brief --report template_profile.json
python scripts/format_document.py target.docx --template sample.docx --output styled.docx --report report.json
python scripts/format_document.py --stdin --input-name draft.md --output formatted.docx --preset brief
python scripts/format_document.py input.docx --diagnose-only --report diagnostics.json
python scripts/format_document.py input.docx --identify-only --report diagnostics.json
python scripts/format_document.py input.docx --output numbered.docx --preset formal --add-page-numbers --report report.json
python scripts/format_document.py input.docx --output table-normalized.docx --preset formal --format-tables --report report.json
```

Behavior:

- Default to preserving text.
- Report `content_preservation` and `coverage` when possible.
- Report `format_changes.page` for page setup, text-area, and document-grid changes after formatting.
- Report `format_changes.paragraph_controls` for keep-with-next, keep-together, widow control, outline levels, tab stops, and numbering/bullet diagnostics.
- Format existing page-number fields when detected. Add page numbers only when the user explicitly asks and `--add-page-numbers` is used.
- Format table text by default. Normalize table structure only when the user explicitly asks and `--format-tables` is used.
- Preserve Word automatic numbering definitions and manual bullet markers by default; style literal heading numbers as part of the existing paragraph text.
- Generate a `.docx` unless `--diagnose-only` / `--identify-only` is used.
- Generate a JSON report when `--report` is provided.
- Do not include full paragraph text in the report by default.
- Use `--include-text-in-report` only for non-sensitive samples or when the user explicitly asks.
- In template mode, report any roles that fell back to a preset because the template did not contain a matching style.
- In template replication work, prefer `--extract-template` first so the user can review covered styles and unresolved items before applying.

Failure handling:

- Missing `python-docx`: install `scripts/requirements.txt`, then rerun the same command.
- Unavailable fonts such as `方正小标宋简体`, `仿宋_GB2312`, `楷体_GB2312`, or `Times New Roman`: still set the requested font names in `.docx`; warn that final rendering depends on the user's Word/WPS font environment.
- Missing LibreOffice/OpenOffice renderer: do not block normal formatting or diagnostics; say render-level smoke checks were skipped unless the user requires them.
- Protected, corrupt, or unsupported `.docx`: stop before rewriting content manually; report the file-level failure and ask for an unlocked/valid copy.
- Documents with tracked changes, comments, fields, TOC, cross-references, footnotes, endnotes, floating text boxes, shapes, seals, or embedded objects: preserve and report them unless the user explicitly asks for a separate risky operation.
- Content-preservation mismatch in the report: treat it as a failed formatting run. Do not deliver the formatted file as final until the cause is understood or the user accepts the risk.

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
