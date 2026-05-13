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

## Official Element Scope

When formatting a formal official document, check for these elements from the official-document processing regulation:

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

The formatter should not invent missing official elements. If a formal document lacks an element that may be required for its intended use, flag it rather than filling it.

## Official Formatting Range

Use this checklist when creating a format plan or report.

### 1. Paper and Page

- Paper: A4, 210 mm x 297 mm.
- Orientation: portrait unless the user or template explicitly requires otherwise.
- Page margins / layout frame:
  - top margin: 37 mm;
  - bottom margin: 35 mm;
  - left binding margin: 28 mm;
  - right margin: 26 mm;
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
| 标题 | 2号小标宋体; centered; may wrap into trapezoid/diamond-like balanced lines |
| 正文 | 3号仿宋体; justified; first-line indent 2 characters |
| 第一层标题 `一、` | 3号黑体 |
| 第二层标题 `（一）` | 3号楷体 |
| 第三层标题 `1.` | 3号仿宋体 |
| 第四层标题 `（1）` | 3号仿宋体 |
| 发文字号、主送机关、附件说明、署名、日期、附注、抄送、印发机关等 | usually 3号仿宋体 unless the template requires otherwise |
| 签发人姓名 | usually 3号楷体 |

If exact fonts are unavailable on the machine, still set the requested font names in the `.docx` and warn that rendering depends on local font installation.

### 4. Paragraph and Heading Hierarchy

- Body paragraphs: first-line indent 2 Chinese characters.
- Body line spacing: fixed line rhythm compatible with the 22 lines/page reference.
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

- Formal documents should have page numbers when more than one page.
- If the source/template has no page number and the user did not specify one, ask whether to add page numbers for formal output.
- Avoid arbitrary headers/footers in formal mode unless they are part of the official template.

### 7. Attachments

- Attachment description belongs after the body when attachments are referenced.
- Attachment content should be formatted separately and clearly identified.
- If the document mentions attachments but no attachment files/content are provided, flag the missing attachment rather than inventing it.

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

## Format Source Priority

Use this priority order:

1. Explicit user instructions.
2. Official unit template supplied by the user.
3. `GB/T 9704-2012` formal preset for formal 公文.
4. Internal brief preset that refers to formal-document typography but is marked as non-official.
5. Ask the user.

If the user says to proceed without answering, use the relevant preset and report every fallback decision.

## Format Plan Before Processing

When the user asks for format-only work and the format source is unclear, present a concise plan before editing:

```text
我将按 [formal/brief/template] 处理，范围包括：
- 页面：A4、页边距、版心/行数参考、页码策略
- 元素：标题、主送机关、正文、附件、落款、日期等
- 字体：标题/正文/各级标题字体字号
- 段落：缩进、行距、段前段后、对齐
- 装饰：颜色、下划线、分隔线按规范/模板处理

未明确项：[列出]
默认处理：[列出]
是否继续？
```

For sensitive internal materials, keep the plan at the style level and do not quote the document text.
