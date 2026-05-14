# 公文/内部简报格式处理 Skill

`gongwen-format-converter` 是一个面向 Codex / TRAE Skills 的 Word 格式处理 Skill。

它的核心用途是：**在不改写、不补写、不重排正文内容的前提下，把已有 `.docx`、Markdown、纯文本或粘贴材料整理成公文/内部简报风格的 Word 版式**。

它不是公文内容生成器，也不是材料润色器。它只处理已有内容和已有对象的格式。

## 快速安装

### 方式一：让 Agent 安装

把下面这段话发给支持 Skills 的 agent：

```text
请从这个 GitHub 地址安装 gongwen-format-converter skill：
https://github.com/danxiaozhiren/gongwen-format-converter-skill/tree/main/skills/gongwen-format-converter
```

安装完成后，重启 Codex / TRAE / 你的 agent 环境，让新 Skill 被重新发现。

### 方式二：手动下载安装

下载或克隆本仓库，然后把这个目录：

```text
skills/gongwen-format-converter
```

复制到本地 skills 目录。常见位置：

```text
Windows: %USERPROFILE%\.codex\skills\gongwen-format-converter
macOS/Linux: ~/.codex/skills/gongwen-format-converter
```

复制后，目标目录下应能直接看到：

```text
SKILL.md
agents/
evals/
references/
scripts/
```

然后重启 agent。

## 快速开始

安装后，建议在请求里显式点名 Skill：

```text
$gongwen-format-converter

模式：现文格式化
输入文件：D:\path\材料.docx
预设：formal
要求：内容不要改、不补写、不润色，只调整已有内容和对象的格式，并给出覆盖报告。
```

常用请求示例：

```text
$gongwen-format-converter

这个 docx 内容已经定稿了，只帮我按正式公文习惯调格式，不要改正文，不要补缺失要素。
```

```text
$gongwen-format-converter

先不要改文档，帮我诊断这份材料当前的页面、标题、正文、行距、缩进、表格、图片、页眉页脚和样式问题。
```

```text
$gongwen-format-converter

参考这份模板的格式，套到目标文档上。只学习模板样式，不复制模板正文，也不补目标文档缺失内容。请先给出模板覆盖情况。
```

## 它会做什么

- 调整已有内容的 Word 版式：页面、页边距、版心、文档网格、字体、字号、颜色、行距、缩进、段前段后、标题层级等。
- 识别并处理已有段落角色：标题、主送机关、正文、附件说明、落款、日期、附注、抄送、版记等。
- 处理已有表格中的文字格式；在用户明确要求时，可进一步规范表格宽度、边框、单元格边距、垂直居中和跨页表头。
- 诊断图片、印章、文本框、形状、页眉页脚、页码、批注、修订、域代码等对象和特殊状态。
- 输出 `.docx`，并可生成 JSON 报告，说明文本是否保持不变、哪些格式已处理、哪些被保留、哪些需要人工确认。

## 它不会默认做什么

- 不改写、不润色、不总结、不删减正文。
- 不调整原有段落顺序。
- 不自动补写主送机关、发文字号、发文机关、签发人、附件名称、落款、日期、印章、抄送机关、印发机关、版记等缺失要素。
- 不把手写编号 `一、`、`（一）`、`1.` 强行改成 Word 自动编号。
- 不移动、缩放或重排图片、印章、文本框、复杂形状，除非用户明确要求。
- 不联网处理用户文档正文。

## 三种模式

| 模式 | 适合场景 | 输出 |
| --- | --- | --- |
| 现文格式化 | 内容已定稿，只需要调 Word 格式 | 格式化后的 `.docx` 和覆盖报告 |
| 现文格式诊断 | 先检查格式问题，不改文件 | 完整诊断报告 |
| 现文模板套用 | 用一份模板/范文的格式套到目标文档 | 先输出模板覆盖分析，确认后再生成目标 `.docx` |

如果用户只给了一份材料但没有说清楚要做什么，Skill 应该先询问选择哪种模式。

## 格式覆盖范围

| 类型 | 当前处理方式 |
| --- | --- |
| 页面级格式 | 自动处理 A4、页边距、版心、文档网格、页眉页脚距离等安全项 |
| 段落级格式 | 自动处理对齐、缩进、固定行距、段前段后、孤行控制、大纲级别、分页控制等 |
| 文字级格式 | 自动处理中文/西文字体、字号、颜色、加粗、斜体、下划线、上下标、字符间距等可识别项 |
| 已有公文元素 | 只对已存在的标题、正文、附件、落款、日期、版记等做格式处理 |
| 页码 | 默认只诊断/格式化已有 PAGE 字段；明确要求时才添加页码 |
| 表格 | 默认处理表内文字并诊断结构；明确要求时才规范表格结构 |
| 图片和对象 | 默认诊断并保留，避免误动印章、图片、文本框、形状 |
| 特殊状态 | 批注、修订、目录、复杂域、脚注尾注、书签等默认诊断并保留 |

## 报告说明

报告重点看这几项：

- `content_preservation`：正文和表格文本是否保持不变。
- `coverage`：哪些范围已格式化、已保留、仅诊断、未检测、暂不支持或需要人工确认。
- `format_changes.page`：页面、版心、文档网格等页面级格式处理前后的变化。
- `format_changes.paragraph_controls`：分页控制、大纲级别、制表位、编号/项目符号等段落控制项的变化。

报告默认不输出完整正文，避免泄露内部材料内容。

## 高级用法

脚本依赖 `python-docx`。如需直接运行脚本，可先安装依赖：

```powershell
python -m pip install -r skills\gongwen-format-converter\scripts\requirements.txt
```

格式化文档：

```powershell
python skills\gongwen-format-converter\scripts\format_document.py D:\path\材料.docx `
  --output D:\path\材料_格式化.docx `
  --preset formal `
  --report report.json
```

只诊断，不生成新文档：

```powershell
python skills\gongwen-format-converter\scripts\format_document.py D:\path\材料.docx `
  --diagnose-only `
  --preset formal `
  --report diagnostics.json
```

明确要求添加页码时：

```powershell
python skills\gongwen-format-converter\scripts\format_document.py D:\path\材料.docx `
  --output D:\path\材料_带页码.docx `
  --preset formal `
  --add-page-numbers `
  --report report.json
```

明确要求规范表格结构时：

```powershell
python skills\gongwen-format-converter\scripts\format_document.py D:\path\材料.docx `
  --output D:\path\材料_表格规范.docx `
  --preset formal `
  --format-tables `
  --report report.json
```

提取模板格式并分析目标文档覆盖情况：

```powershell
python skills\gongwen-format-converter\scripts\format_document.py `
  --extract-template D:\path\模板.docx `
  --target D:\path\新材料.docx `
  --preset brief `
  --report template_profile.json
```

## 当前限制

- 图片、印章、浮动文本框、复杂形状、水印等对象目前以诊断和保留为主。
- 批注、修订、复杂域、目录、脚注尾注、书签等特殊结构不会被自动重写。
- Word 自动编号定义会被检测和保留，但不会默认重建编号体系。
- 最终显示效果仍可能受本机字体影响，例如 `方正小标宋简体`、`仿宋_GB2312` 是否安装。

## Skill 目录

```text
skills/gongwen-format-converter/
```

核心行为规则在：

```text
skills/gongwen-format-converter/SKILL.md
```
