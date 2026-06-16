# 公文/内部简报格式处理 Skill

`gongwen-format-converter` 是一个面向 Agent 的 Word 格式处理 Skill，适用于支持 Skills 或类似可扩展技能机制的智能体环境。

它的核心用途是：**在不改写、不补写、不重排正文内容的前提下，把已有 `.docx`、Markdown、纯文本或粘贴材料整理成公文/内部简报风格的 Word 版式**。

它不是公文内容生成器，也不是材料润色器，也不是完整 Word 排版引擎。它只处理已有内容和已有对象中可安全处理的格式；复杂对象默认诊断和保留。

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
- 诊断图片、印章、文本框、形状、页眉页脚、页码、批注、修订、域代码等对象和特殊状态；对高风险对象默认不主动重写。
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
| 现文格式诊断 | 先检查格式问题，不改文件 | 覆盖诊断报告 |
| 现文模板套用 | 用一份模板/范文的格式套到目标文档 | 先输出模板覆盖分析，确认后再生成目标 `.docx` |

如果用户只给了一份材料但没有说清楚要做什么，Skill 应该先询问选择哪种模式。

## 格式参考依据

正式公文格式优先参考以下官方规范：

- 《党政机关公文处理工作条例》（中共中央办公厅、国务院办公厅，中办发〔2012〕14号）
  - 中国政府网：<https://www.gov.cn/zwgk/2013-02/22/content_2337704.htm>
- `GB/T 9704-2012`《党政机关公文格式》
  - 全国标准信息公共服务平台：<https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=F3CC9BEF482524C895FDA7A08BB4A70E>
  - 该平台显示标准状态为“现行”，发布日期为 `2012-06-29`，实施日期为 `2012-07-01`。

这个 Skill 的 `formal` 预设会把这些规范转化为可执行的 Word 格式处理规则，例如：

- A4 纸张、版心、页边距、文档网格等页面设置。
- 标题、正文、各级标题、附件说明、落款、日期、版记等已有元素的字体、字号、行距、缩进和对齐。
- 正文常用的 3 号仿宋、标题层级字体、固定行距、每页行数/每行字数参考等格式。
- 页眉页脚、页码、表格、图片对象和特殊状态的诊断与保留策略。

注意：官方规范用于判断“已有元素应该如何排版”，不是用来让 Skill 自动补写缺失元素。实际单位模板、地方细则或用户给定模板与通用规范不一致时，优先按用户提供的模板和明确要求处理，并在报告中说明。

## 格式覆盖范围

详细状态矩阵见 [`docs/word-format-coverage.md`](docs/word-format-coverage.md)。当前项目的定位是“常见公文/简报 Word 安全格式整理”，不是 Word 全对象、全样式、全视觉效果的完整处理。

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
| 渲染级校验 | 当前不做 PDF/截图级视觉比对，最终显示仍需用 Word/WPS 人工确认 |

## 报告说明

报告重点看这几项：

- `content_preservation`：正文、表格、页眉页脚、文本框、批注、脚注和尾注文本是否保持不变。
- `coverage`：哪些范围已格式化、已保留、仅诊断、未检测、暂不支持或需要人工确认。
- `format_changes.page`：页面、版心、文档网格等页面级格式处理前后的变化。
- `format_changes.paragraph_controls`：分页控制、大纲级别、制表位、编号/项目符号等段落控制项的变化；诊断报告还会列出 Word 自动编号的 `num_id`、层级分布，以及手写编号的层级分布和跳变提示。
- `format_diagnostics.special_state.field_diagnostics`：字段类型和类别统计，例如 `TOC`、`PAGE`、`NUMPAGES`、`DATE`、`REF`、`PAGEREF`；默认只报告类型、类别和指令哈希，不输出完整字段指令。
- `format_diagnostics.special_state.field_update_risks`：字段更新风险提示，例如目录、页码、日期和交叉引用在格式化后可能需要到 Word/WPS 中更新域。

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

## 开发验证

仓库包含一组可重新生成的最小 Word 样例，用于检查覆盖矩阵中的关键格式面：

```bash
make smoke-word
make render-word
make test
make compile
```

`make render-word` 是可选渲染 smoke check：需要本机安装 LibreOffice/OpenOffice，或通过 `WORD_RENDERER` 指定 `soffice` 路径；未检测到渲染器时会跳过。需要强制失败时使用 `make render-word-required`。

样例和说明见 [`tests/fixtures/README.md`](tests/fixtures/README.md)。

## 当前限制

- 当前项目不能宣称“完全处理 Word 格式”；复杂对象、完整样式继承和渲染一致性仍在覆盖矩阵中作为后续能力规划。
- 图片、印章、浮动文本框、复杂形状、水印等对象目前以诊断和保留为主。
- 批注、修订、复杂域、目录、脚注尾注、书签等特殊结构不会被自动重写。
- Word 自动编号定义会被检测、保留并报告编号 ID 和层级分布，但不会默认重建编号体系。
- 最终显示效果仍可能受本机字体影响，例如 `方正小标宋简体`、`仿宋_GB2312` 是否安装。

## Skill 目录

```text
skills/gongwen-format-converter/
```

核心行为规则在：

```text
skills/gongwen-format-converter/SKILL.md
```
