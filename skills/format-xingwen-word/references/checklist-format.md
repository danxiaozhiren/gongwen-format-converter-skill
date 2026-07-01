# 行文出手前对照检查事项

Use these fixed Word formatting rules for the `xingwen-checklist` layout. Treat them as an office-format checklist, not as a promise of legal compliance or exact rendering equivalence across Word/WPS/LibreOffice.

## Page

- Paper: A4 portrait.
- Margins: top 3.7 cm, bottom 3.5 cm, left 2.7 cm, right 2.7 cm.
- Text area reference: 156 mm wide x 225 mm high.
- Document grid reference: 22 lines per page and 28 characters per line.
- Whole document line spacing: fixed 30 pt for body and heading levels.

## Title

- Main title: 方正小标宋简体, 2号, not bold, centered.
- Multi-line main title: fixed 36 pt line spacing.
- Main title to body: leave the visual equivalent of one 3号 line after the title; the script applies 30 pt after-spacing.
- Do not claim automatic 梯形/正菱形 visual balancing. Preserve existing line breaks and ask for manual review when title shape matters.

## Body And Headings

- Body: 仿宋_GB2312, 3号, not bold, justified, first-line indent 2 Chinese characters.
- Level 1 heading: 黑体, 3号, not bold, marker `一、`.
- Level 2 heading: 楷体_GB2312, 3号, bold, marker `（一）`.
- Level 3 heading: 仿宋_GB2312, 3号, not bold, marker `1.` / `1．` / `1、`.
- Level 4 heading: 仿宋_GB2312, 3号, not bold, marker `（1）`.
- Latin letters and numbers: Times New Roman.

## Page Numbers

- Existing page numbers: format as 宋体, 4号, centered when safely detected.
- Missing page numbers: do not add unless the user explicitly asks and the script is run with `--add-page-numbers`.

## Boundaries

- Do not generate red-head/版头 red lines, issuer marks, document-number positions, signatures, seals, copy/printing areas, or other absent official elements.
- Do not batch-process directories.
- Do not parse Markdown tables or map Markdown rich-text markers such as `**bold**` to Word styles.
- Keep complex Word objects conservative: preserve or diagnose images, seals, text boxes, shapes, comments, tracked changes, fields, TOC, footnotes, and endnotes unless a supported explicit option handles them.
