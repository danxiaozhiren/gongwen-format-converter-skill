---
name: format-xingwen-word
description: Format existing Chinese `.docx` documents with the fixed 行文出手前对照检查事项 layout. Use when the user wants Word fonts, spacing, margins, title hierarchy, page numbers, or table formatting adjusted without changing content.
---

# 行文检查版式 Word 助手

## Scope

Process only existing `.docx` files. Apply the fixed “行文出手前对照检查事项” Word layout to existing content without rewriting, adding, deleting, reordering, summarizing, or polishing any substantive text.

Use this skill for two modes only:

- `format-only`: adjust formatting and produce a `.docx`.
- `diagnose-only`: inspect the document and write a JSON report without producing a `.docx`.

Do not use this skill for Markdown, plain text, pasted text conversion, template replication, batch processing, OCR/PDF conversion, red-head/版头 generation, or writing missing official elements.

## Required References

- Read `references/checklist-format.md` when applying or explaining the fixed layout.
- Read `references/role-detection.md` when paragraph role classification is ambiguous or when changing detection behavior.

## Zero-Change Rule

This rule has priority over all formatting choices.

- Preserve existing text, paragraph order, table text, images, headers, footers, comments, fields, and document objects where possible.
- Do not create missing 主送机关、发文字号、签发人、署名、日期、附件、印章、抄送、版记, or page numbers unless the user explicitly asks for page numbers and `--add-page-numbers` is used.
- If content preservation fails, do not deliver the candidate file. Report `failed_content_preservation`, `output_withheld: true`, and the requested output path.
- If the user asks to补写、润色、改写、总结 or infer missing official elements, refuse that part and ask for an authoritative `.docx` whose content is already final.

## Interaction Contract

Every user-facing step must make three things clear:

- 能做什么: only existing `.docx` layout formatting or diagnostics, with zero content changes.
- 当前在干什么: waiting for file/mode, running diagnostics, formatting candidate, checking content preservation, or ready to deliver.
- 下一步做什么: ask for the missing choice/file, run the script, inspect the report, deliver output, or stop delivery.

When the user gives a vague request such as “处理一下”, “整理一下”, “按要求调格式”, or only sends a Word file, ask a short mode question before running. Do not make the user infer what the skill is about.

Mode prompt:

```text
我可以做两件事，都会保持正文不改：
1. 现文格式化：按“行文出手前对照检查事项”整理版式，生成新的 .docx。
2. 现文格式诊断：只检查页面、标题、正文、表格、页码等格式，生成报告，不改 Word。

当前我已准备处理已有 .docx。下一步请选择 1 或 2；如果要加页码或规范表格，请一并说明。
```

Use `format-only` without asking only when the user clearly requests a formatted Word output. Use `diagnose-only` when the user says “先诊断/只检查/不要生成新文件”.

- Treat “加页码” as explicit permission to use `--add-page-numbers`; otherwise only format existing page-number fields.
- Treat “表格也按清单整理/规范一下” as explicit permission to use `--format-tables`; otherwise preserve table structure and format table text only.
- If the user asks to补写、润色、改写、总结 or infer missing official elements, refuse that part and ask whether to proceed with zero-change formatting or diagnostics of the existing `.docx` only.

Missing-file prompt:

```text
我能处理的是已有 .docx 的版式整理或诊断，不改正文、不补写。
当前还没有可处理的 Word 文件。
下一步请发我已有的 .docx 文件或本地路径；Markdown、纯文本、PDF、截图或粘贴文本不能用这个 skill 转 Word。
```

Unsupported-content prompt:

```text
这个 skill 不补写、不润色、不改正文。
当前请求里包含内容改写或补要素，这部分我不能做。
下一步可以继续做 1. 现文格式化，或 2. 现文格式诊断；需要我按哪个范围继续？
```

## Workflow

1. Confirm the input is an existing `.docx`; if not, ask for one and do not run the script.
2. Confirm mode when the user's intent is vague. Do not run until the user has chosen formatting or diagnostics.
3. Before running, state one sentence with status and next step: “当前我将只调版式，不改正文；下一步会生成候选文件并先检查内容保持报告，通过后再交付。”
4. Run `scripts/format_document.py` from the skill directory or prefix paths from the repository root. Always pass `--report`.
5. Read the JSON report before responding. Check `summary.status`, `summary.user_message`, `content_preservation`, `output_withheld`, `coverage`, `role_counts`, and `warnings`.
6. Deliver the `.docx` path only when `summary.output_withheld` is false and `summary.output` is not null.

## Script Commands

Install dependencies if needed:

```bash
python -m pip install -r scripts/requirements.txt
```

Format an existing Word document:

```bash
python scripts/format_document.py input.docx --output output.docx --report report.json
```

Diagnose without writing output:

```bash
python scripts/format_document.py input.docx --diagnose-only --report diagnostics.json
```

Format existing page numbers and add missing page numbers only when explicitly requested:

```bash
python scripts/format_document.py input.docx --output numbered.docx --add-page-numbers --report report.json
```

Normalize existing table structure only when explicitly requested:

```bash
python scripts/format_document.py input.docx --output table-normalized.docx --format-tables --report report.json
```

## Response

After processing, keep the response concise:

- output path, or say output was withheld;
- report path;
- mode: `format-only` or `diagnose-only`;
- format source: `xingwen-checklist`;
- content-preservation status;
- important role counts, warnings, and unsupported areas;
- note that strict visual checks still require opening in Word/WPS when 22 lines per page and 28 characters per line matter.

Use `summary.user_message` as the first sentence when it exists. Do not dump raw JSON to the user unless they ask for it.
Never quote confidential document text unless the user explicitly asks.
