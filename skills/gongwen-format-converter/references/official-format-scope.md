# Official Format Scope

Use this file as the first reference for formal 公文 formatting decisions.

## Normative Sources

Primary sources:

- 《党政机关公文处理工作条例》: defines document types, official document elements, A4 paper use, and requires layouts to follow the national standard.
- `GB/T 9704-2012 党政机关公文格式`: current recommended national standard for party/government official-document layout. The national standards platform lists it as current and continuing effective after the 2025-05-30 review.

Scope of `GB/T 9704-2012`:

- paper requirements;
- page layout;
- printing and binding requirements;
- arrangement rules for official-document elements;
- sample layouts.

Do not present internal brief conventions as official law or mandatory GB/T rules. Internal briefs may refer to these rules for clean office-document presentation.

## Existing-Content Application Rule

Use official sources as formatting guidance for existing content and objects only.

- Do not add, rewrite, delete, reorder, or infer substantive content.
- Do not create missing official-document elements such as 主送机关, 发文字号, 签发人, 附件说明, 署名, 日期, 印章, 抄送机关, 印发机关, 版记, or 页码 unless the user explicitly asks for that separate addition.
- If an official element is absent, report it as `not_detected` / `not_processed`.
- If an existing element is ambiguous, mark it `needs_review` before applying risky formatting.
- Use `formatted`, `preserved`, `diagnosed_only`, `not_detected`, `unsupported`, and `needs_review` to make coverage transparent.

## Official Element Scope

When formatting a formal official document, detect these elements from the official-document processing regulation and apply formatting only when they already exist:

| Element | Formatting Role |
| --- | --- |
| 份号 | Copy sequence number; used for classified documents |
| 密级和保密期限 | Classification level and duration |
| 紧急程度 | Urgency label |
| 发文机关标志 | Issuing authority mark |
| 发文字号 | Issuing number |
| 签发人 | Signatory; required for upward documents |
| 标题 | Issuing authority + subject + document type |
| 主送机关 | Main recipient |
| 正文 | Main body |
| 附件说明 | Attachment order and name |
| 发文机关署名 | Issuing authority signature |
| 成文日期 | Signing/adoption date |
| 印章 | Seal |
| 附注 | Notes such as circulation scope |
| 附件 | Attachment content |
| 抄送机关 | CC recipients |
| 印发机关和印发日期 | Printing/distribution authority and date |
| 页码 | Page number |

The formatter should not invent missing official elements. If a formal document lacks an element that may be required for its intended use, report it as not detected and ask only if the user explicitly wants completeness guidance.

## Official Formatting Range

Use this checklist when creating a format plan or report. Apply automatic changes only to safe existing format properties; preserve and report risky or unsupported items.

The built-in `formal` preset uses the user's default 行文出手前检查表 for routine formatting. If the user asks for a strict national-standard check, or provides a unit template, follow the Format Source Priority below and disclose any difference from the built-in default.

### 1. Paper and Page

- Paper: A4, 210 mm x 297 mm.
- Orientation: portrait unless the user or template explicitly requires otherwise.
- Built-in formal page margins / layout frame:
  - top margin: 37 mm;
  - bottom margin: 35 mm;
  - left margin: 27 mm;
  - right margin: 27 mm;
  - text area reference: 156 mm x 225 mm, excluding page number.
- Line grid reference: normally 22 lines per page and 28 characters per line.

### 2. Text Direction and Color

- Text direction: horizontal left-to-right.
- Main text color: black by default.
- Official issuing authority mark may use red when the document has a formal red-head layout.
- Do not introduce arbitrary colors, shading, italic, or underline in formal mode unless the user or official template requires them.

### 3. Fonts and Sizes

Use these defaults when no official unit template or user override is provided:

| Role | Default Style |
| --- | --- |
| 发文机关标志 | 小标宋体, red, large display size according to template/standard layout |
| 标题 | 2号方正小标宋简体; not bold; centered; multi-line titles should be balanced as trapezoid/diamond-like lines |
| 正文 | 3号仿宋_GB2312; justified; first-line indent 2 characters; fixed 30 pt line spacing |
| 第一层标题 `一、` | 3号黑体; not bold |
| 第二层标题 `（一）` | 3号楷体_GB2312; bold |
| 第三层标题 `1.` | 3号仿宋_GB2312; not bold |
| 第四层标题 `（1）` | 3号仿宋_GB2312; not bold |
| 发文字号、主送机关、附件说明、署名、日期、附注、抄送、印发机关等 | usually 3号仿宋体 unless the template requires otherwise |
| 签发人姓名 | usually 3号楷体 |
| 数字及字母 | Times New Roman |
| 页码 | 4号宋体 |

If exact fonts are unavailable on the machine, still set the requested font names in the `.docx` and warn that rendering depends on local font installation.

### 4. Paragraph and Heading Hierarchy

- Body paragraphs: first-line indent 2 Chinese characters.
- Body line spacing: fixed 30 pt, compatible with the 22 lines/page reference.
- Multi-line main-title spacing: fixed 36 pt.
- Main title to first body paragraph spacing: one 3号-line equivalent when safe to express as spacing rather than inserted content.
- Heading sequence for body structure:
  - first level: `一、`
  - second level: `（一）`
  - third level: `1.`
  - fourth level: `（1）`
- Preserve existing heading text and order in format-only mode.
- If heading levels are inconsistent, report them and ask before renumbering.

### 5. Official-Document Zones

For formal documents, distinguish:

- 版头: 份号、密级和保密期限、紧急程度、发文机关标志、发文字号、签发人.
- 主体: 标题、主送机关、正文、附件说明、发文机关署名、成文日期、印章、附注、附件.
- 版记: 抄送机关、印发机关和印发日期.
- 页码: page number.

Do not force a casual internal brief into a full 版头/版记 layout unless the user asks for formal 公文 output.

### 6. Page Numbers, Headers, and Footers

- Formal documents commonly use page numbers when more than one page.
- If existing page numbers are present, diagnose or format them according to the selected rule/template.
- If the source/template has no page number and the user did not specify one, report `not_detected`; ask before adding page numbers because adding them changes document content/layout.
- Avoid arbitrary headers/footers in formal mode unless they are part of the official template.

### 7. Attachments

- Attachment description belongs after the body when attachments are referenced.
- Attachment content should be formatted separately and clearly identified.
- If the document mentions attachments but no attachment files/content are provided, report the absence rather than inventing attachment names or content.

### 8. Seal and Signature Area

- Preserve existing seal images where possible.
- Do not generate or fabricate seals.
- Do not invent issuing authority signature or date.

### 9. Tables, Images, and Special Objects

`GB/T 9704-2012` centers on official-document layout; tables/images often need document-specific treatment.

Default policy:

- Preserve tables/images unless user asks for restyling.
- Normalize table text to the body font when safe.
- Ask before changing table borders, image placement, seal placement, or complex text boxes.

### 10. Word Format Surface Coverage

For strict existing-content formatting, inspect and report the broad Word formatting surface where detectable, but automatically change only the safe subset that the script can handle reliably:

- page and section setup;
- paragraph and run formatting;
- structural role formatting for detected official/internal elements;
- tables, images, seals, text boxes, shapes, charts, formulas, and embedded objects;
- headers, footers, page numbers, and version-record separator lines when already present;
- styles, direct-formatting conflicts, and style variants;
- comments, tracked changes, fields, table of contents, hyperlinks, bookmarks, footnotes/endnotes, and hidden text.

Automatically change only properties that are safe and within the user's selected format source. Preserve or diagnose the rest. Do not describe this as complete Word-format processing, because complex objects, style inheritance, field updates, and rendered visual equivalence remain outside the current automatic scope.

## Format Source Priority

Use this priority order:

1. Explicit user instructions.
2. Official unit template supplied by the user.
3. `GB/T 9704-2012` formal preset for formal 公文.
4. Internal brief preset that refers to formal-document typography but is marked as non-official.
5. Ask the user.

If the user says to proceed without answering, use the relevant preset for existing content and report every fallback decision.

## Format Plan Before Processing

When the user asks for format-only work and the format source is unclear, present a concise plan before editing:

```text
我将按 [formal/brief/template] 处理，范围包括：
- 页面：A4、页边距、版心/行数参考、已有页码策略
- 元素：只处理已检测到的标题、主送机关、正文、附件、落款、日期、版记等
- 字体：标题/正文/各级标题/表格等已有内容的字体字号
- 段落：缩进、行距、段前段后、对齐、编号和制表位
- 对象：表格、图片、印章、文本框、页眉页脚、批注/修订等按可安全处理范围处理
- 覆盖报告：formatted / preserved / diagnosed_only / not_detected / unsupported / needs_review

未检测到项：[列出，仅报告不补写]
未明确项：[列出需要确认的现有格式项]
默认处理：[列出]
是否继续？
```

For sensitive internal materials, keep the plan at the style level and do not quote the document text.
