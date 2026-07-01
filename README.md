# format-xingwen-word

`format-xingwen-word` 是一个面向 Agent 的 Word-only Skill。它专注一个场景：用户已有 `.docx` 内容定稿后，按“行文出手前对照检查事项”整理版式，**不改写、不补写、不润色、不重排正文内容**。

早期实验名是 `gongwen-format-converter`。当前版本已作为新的 Word-only skill 收敛为 `format-xingwen-word`，不再维护旧入口，也不再承载 Markdown、纯文本、模板套用等泛化能力。

迁移说明：旧目录下的 `evals/`、`references/` 和 `scripts/` 已迁入 `skills/format-xingwen-word/` 并按 Word-only 场景重整。旧 `format-presets.md`、`official-format-scope.md` 不再保留为主能力资料；固定清单规则集中在 `references/checklist-format.md`，角色识别规则集中在 `references/role-detection.md`。

## 安装

请从这个路径安装主 skill：

```text
https://github.com/danxiaozhiren/gongwen-format-converter-skill/tree/main/skills/format-xingwen-word
```

本地安装目录示例：

```text
Windows: %USERPROFILE%\.codex\skills\format-xingwen-word
macOS/Linux: ~/.codex/skills/format-xingwen-word
```

使用时触发：

```text
$format-xingwen-word
```

最省事的说法：

```text
这份 Word 帮我处理一下，内容不要改。
```

如果说法比较模糊，skill 会先用很短的问题把范围说清楚：

```text
我可以做两件事，都会保持正文不改：
1. 现文格式化：生成整理后的 .docx。
2. 现文格式诊断：只生成报告，不改 Word。

当前我已准备处理已有 .docx。下一步请选择 1 或 2；如果要加页码或规范表格，请一并说明。
```

交互原则是让用户始终知道三件事：这个 skill 能做什么、当前处在哪一步、下一步会做什么。用户明确要求“整理成 Word/生成新文件”时会直接格式化；明确说“先诊断/只检查”时只生成诊断报告。

## 能做什么

- 处理已有 `.docx`。
- 按固定行文检查版式设置 A4、页边距、文档网格参考、标题、正文、四级标题、行距、缩进、中文字体、西文字体。
- 诊断 `.docx` 当前页面、段落、角色、表格、页眉页脚、页码字段、图片/对象、批注/修订/域等状态。
- 保留正文和表格文本顺序，失败时拦截输出并返回非零退出码。
- 格式化已有页码字段。
- 在用户明确要求时用 `--add-page-numbers` 添加页码。
- 在用户明确要求时用 `--format-tables` 规范表格结构。

## 固定版式

默认规则来自“行文出手前对照检查事项”：

- 页面：A4；上 3.7 cm，下 3.5 cm，左 2.7 cm，右 2.7 cm。
- 网格参考：每页 22 行，每行 28 字；版心参考 156 mm x 225 mm。
- 大标题：方正小标宋简体，二号，不加粗；多行标题固定 36 磅。
- 正文：仿宋_GB2312，三号，不加粗，固定 30 磅。
- 一级标题：黑体，三号，不加粗，序号 `一、`。
- 二级标题：楷体_GB2312，三号，加粗，序号 `（一）`。
- 三级标题：仿宋_GB2312，三号，不加粗，序号 `1.`。
- 四级标题：仿宋_GB2312，三号，不加粗，序号 `（1）`。
- 数字和字母：Times New Roman。
- 页码：宋体，四号。

严格的 22 行/28 字视觉效果仍需在 Word/WPS 中打开确认；脚本写入的是 Word 网格和段落格式参考，不承诺跨渲染器完全一致。

## 不会做什么

- 不处理 Markdown、`.txt`、粘贴纯文本或 stdin 转 Word。
- 不做模板复制、模板样式提取或范文内容回显。
- 不批量处理多个文件。
- 不生成红头、版头红线、发文字号位置、签发人、署名、日期、印章、抄送、版记等缺失公文要素。
- 不解析 Markdown 表格或把 `**bold**`、`_italic_` 等标记映射为 Word 富文本。
- 不做 OCR、PDF 转 Word 或扫描图片识别。

## 命令行

安装依赖：

```bash
python -m pip install -r skills/format-xingwen-word/scripts/requirements.txt
```

格式化已有 Word：

```bash
python skills/format-xingwen-word/scripts/format_document.py input.docx \
  --output output.docx \
  --report report.json
```

只诊断，不生成 Word：

```bash
python skills/format-xingwen-word/scripts/format_document.py input.docx \
  --diagnose-only \
  --report diagnostics.json
```

明确要求添加页码：

```bash
python skills/format-xingwen-word/scripts/format_document.py input.docx \
  --output numbered.docx \
  --add-page-numbers \
  --report report.json
```

明确要求规范表格结构：

```bash
python skills/format-xingwen-word/scripts/format_document.py input.docx \
  --output table-normalized.docx \
  --format-tables \
  --report report.json
```

交互提示：

- 非 `.docx` 输入会直接拒绝，并提示需要已有 Word 文件。
- 损坏、加密或无法打开的 `.docx` 会给出明确错误，不会生成输出。
- 成功报告里会包含 `summary.user_message`，包括 `can_do`、`current_state`、`next_action`，Agent 可直接用它向用户说明结果。
- 内容保持校验失败时，候选文件会被删除，正式输出不会生成或覆盖。

## 验证

```bash
make check
make skill-validate
make evals
make coverage-matrix
make golden-word
make smoke-word
make render-word
```

`make render-word` 依赖 LibreOffice/OpenOffice；当前环境缺少渲染器或相关动态库时允许跳过。覆盖边界见 [`docs/word-format-coverage.md`](docs/word-format-coverage.md)。

## 项目结构

```text
skills/format-xingwen-word/
  SKILL.md
  agents/openai.yaml
  scripts/format_document.py
  references/checklist-format.md
  references/role-detection.md
  evals/evals.json
```
