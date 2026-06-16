# Paragraph Role Detection

Classify paragraphs conservatively. The goal is to apply formatting without changing content or inventing missing structure.

## Role Rules

`main_title`

- First short centered-looking paragraph in a `.docx`, or first Markdown `#` heading.
- Often contains 文种 words such as 通知、请示、报告、函、纪要、简报.
- For internal brief documents, the first line may be the brief name rather than the article title.

`issue_number`

- Short line matching `第X期`, `第 X 期`, or similar.
- Common in internal briefs.

`metadata`

- Short line near the top containing a publisher, department, date, or both.
- Often appears after an issue number in an internal brief.

`separator`

- A line made mostly of dashes, underscores, equals signs, or box-drawing characters.
- In `.docx`, this can also be represented by a bottom border; if detection is uncertain, preserve and style as a separator only when text clearly indicates it.

`article_title`

- In internal briefs, a short title after the brief header and separator.
- Usually centered in the source or visually distinct.

`recipient`

- Appears before body text in formal documents.
- Often ends in `：` and names an organization or group, such as `各部门：`.

`heading_1`

- Starts with Chinese numbering such as `一、`, `二、`, `三、`.

`heading_2`

- Starts with parenthesized Chinese numbering such as `（一）`, `（二）`, `（三）`.

`heading_3`

- Starts with Arabic numbering such as `1.`, `1．`, `1、`.

`heading_4`

- Starts with parenthesized Arabic numbering such as `（1）`, `（2）`, `（3）`.

`attachment`

- Starts with `附件：`, `附件1：`, `附件：1.`, or similar.

`signature`

- Short line near the end that looks like an issuing department, office, company, center, committee, or unit name.
- Do not invent this role when the line could be ordinary body text.

`date`

- Short line near the end matching `YYYY年M月D日` or similar.

`body`

- Default role for normal prose.

`needs_review`

- Use when a paragraph is short and title-like but its role is unclear.
- Use for unusual numbering, mixed hierarchy, or lines that appear to be metadata but lack dates/organizational cues.

## Markdown Mapping

- `#` maps to `main_title`.
- `##` maps to `article_title` for brief preset, otherwise `heading_1` unless it is the first heading.
- `###` maps to `heading_1` or `heading_2` depending on surrounding structure.
- Plain paragraphs are classified by numbering and position.
- Preserve list item text; remove Markdown list markers only if they are purely formatting and not part of the user's content.

## Ambiguity Handling

When role detection is uncertain:

- Preserve the paragraph text and order.
- Apply body formatting rather than risky heading formatting.
- Add a warning to the report.
- Ask the user for confirmation only when the ambiguity materially affects the output.

## Template Replication

Template replication should extract:

- Page margins and paper size.
- Representative font name, size, bold, italic, underline, font color, alignment, first-line indent, left/right indent, paragraph spacing, and line spacing by detected role.
- Header-like sequence for internal briefs.

It should not extract or report full template text by default.

If the target document has roles not found in the template, mark those roles as fallback candidates. Prefer asking the user whether to use preset formatting, preserve target formatting, or define a custom style. If no confirmation is possible, use the relevant 公文/内部简报 preset and report the fallback.
