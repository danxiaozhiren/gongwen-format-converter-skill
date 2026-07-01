# Paragraph Role Detection

Classify existing Word paragraphs conservatively. The goal is to apply the fixed checklist layout without changing content or inventing missing structure.

## Roles

`main_title`

- First high-confidence title-like paragraph near the beginning.
- Often contains 文种 words such as 通知、请示、报告、函、纪要.
- Do not classify dates, document numbers such as `中办发〔2024〕1号`, issue numbers, recipients, or metadata as the main title.

`recipient`

- Early paragraph ending in `：`, such as `各部门：`.

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

- Short organization-like line near the end and usually before a trailing date.
- Do not invent this role when the line could be ordinary body text.

`date`

- Short line near the end matching `YYYY年M月D日`.

`metadata`

- Document number, issue number, signer marker, or other short administrative metadata.

`body`

- Default role for normal prose.

`needs_review`

- Use when a short title-like or metadata-like paragraph is ambiguous.
- Reports should include `confidence`, `reason_codes`, and warnings so the agent knows what needs human confirmation.

## Ambiguity Handling

- Preserve paragraph text and order.
- Prefer body formatting over risky heading/title formatting when confidence is low.
- Ask the user only when ambiguity materially affects the deliverable.
- Never fill missing official-document elements during role detection.
