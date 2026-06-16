# Format Presets

Use these presets as practical office-document formatting defaults. They are not a substitute for a user's official unit template. If the user supplies a template, prefer the template.

## Formal Preset

Use for 通知、请示、报告、函、纪要 and other formal 公文-style drafts. The built-in default follows the user's 行文出手前检查表 unless a supplied unit template or explicit user instruction overrides it.

Page:

- Paper: A4.
- Margins: top 37 mm, bottom 35 mm, left 27 mm, right 27 mm.
- Text area reference: 156 mm x 225 mm.
- Document grid reference: 22 lines per page and 28 characters per line where Word grid settings can be applied safely.
- Body line rhythm: fixed line spacing 30 pt.
- Main-title line rhythm: fixed line spacing 36 pt for multi-line titles.
- Main-title-to-body gap: use one 3号-line equivalent when formatting can express it safely.

Text:

- Main title: 2号 方正小标宋简体, not bold, centered.
- Multi-line main title: prefer balanced trapezoid/diamond-like line breaks; do not force an equal-width rectangle or a short-middle hourglass layout.
- Body: 3号 仿宋_GB2312, black, justified, first-line indent 2 Chinese characters.
- First-level heading `一、`: 3号 黑体, not bold.
- Second-level heading `（一）`: 3号 楷体_GB2312, bold.
- Third-level heading `1.` or `1、`: 3号 仿宋_GB2312, not bold.
- Fourth-level heading `（1）`: 3号 仿宋_GB2312, not bold.
- Recipient: 3号 仿宋_GB2312, left aligned, no first-line indent, ends with full-width colon when present.
- Attachment line: 3号 仿宋_GB2312, first-line indent 2 Chinese characters.
- Signature/date: 3号 仿宋_GB2312, right aligned.
- Digits and Latin letters: Times New Roman.
- Page numbers: 4号 宋体.

Paragraph:

- Body paragraph: first-line indent 2 characters.
- Headings inside the body usually keep a 2-character first-line indent.
- Paragraph spacing: 0 before/after unless the template shows otherwise.
- Safe paragraph controls: use widow control for body text, keep headings with the following paragraph, and assign outline levels to detected heading roles.
- Avoid decorative styles, excessive bolding, colored text, and casual Markdown-like bullets.
- Use black text by default. Remove decorative underline, italic, and arbitrary color unless the user or an official template requires them.
- Existing literal numbering markers such as `一、`, `（一）`, `1.` and `（1）` remain text; do not convert them to Word automatic numbering unless explicitly requested.

## Brief Preset

Use for internal 信息简报、会议简报、工作动态、生产经营分析会材料.

Page:

- Paper: A4.
- Margins: top 25 mm, bottom 25 mm, left 28 mm, right 26 mm unless the user's template differs.
- Text area reference: 156 mm x 247 mm with a clean line grid when safe.
- Keep a clean internal-office layout rather than a full red-head official-document layout.

Text:

- Brief name: 小二 or 2号, 宋体/小标宋, bold or strong, centered.
- Issue number: 3号 楷体 or 仿宋, centered.
- Publisher/date line: 3号 仿宋, left aligned or tab-aligned according to source/template.
- Separator: thin horizontal rule if present in source/template.
- Article title: 2号 or 小二, 宋体/小标宋, bold, centered.
- Body: 3号 仿宋_GB2312, justified, first-line indent 2 Chinese characters.
- First-level heading: 3号 黑体.
- Second-level heading: 3号 楷体_GB2312.
- Third-level heading: 3号 仿宋_GB2312.
- Use black text by default. Keep separator lines only when they are part of the brief header or source/template structure.
- Keep heading/article-title paragraphs with the following paragraph when safe; preserve existing Word numbering and manual bullets unless explicitly instructed otherwise.

Structure:

Typical internal brief order:

1. Brief name
2. Issue number
3. Publisher and date
4. Separator line
5. Article title
6. Lead paragraph
7. Body paragraphs and numbered requirements

Do not force all internal briefs into this order. Preserve the user's paragraph order unless asked.

## Font Notes

Chinese office documents often refer to fonts that may not be installed on every machine, such as 方正小标宋简体 or 仿宋_GB2312. It is still useful to set the requested font names in the `.docx`; Word will render the closest available font if needed. Tell the user if exact rendering depends on local font availability.

## Template Fallback Policy

When a template does not contain a comparable role, prefer asking the user what to do. If the user asks to proceed, apply the closest preset role:

- Missing title-like style: use `main_title` or `article_title` from the chosen preset.
- Missing numbered heading style: use the corresponding `heading_1`, `heading_2`, or `heading_3` preset.
- Missing body style: use preset body formatting.
- Missing decorative detail such as underline or color: normalize to standard black, no underline, no italic unless the template clearly used that detail for the same role.

Always mention fallback decisions in the report so the user can decide whether to rerun with custom instructions.
