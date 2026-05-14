# 公文/内部简报格式整理 Skill

这个仓库沉淀了一个面向 TRAE SOLO / Codex Skills 的公文格式处理 Skill：`gongwen-format-converter`。

它的核心目标是：在不改写、不补写、不重排正文内容的前提下，把 `.docx`、Markdown、纯文本或粘贴材料整理成符合公文/内部简报习惯的 Word 版式，并给出透明的格式覆盖报告。

## 当前定位

这个 Skill 是一个“现有内容 Word 全格式处理器”，而不是“公文内容补全器”。

最高原则：

- 只处理已有内容和已有对象的格式。
- 不新增主送机关、发文字号、日期、落款、附件、版记、印章、页码等缺失要素。
- 不改写、不润色、不总结、不删除、不调整段落顺序。
- 未检测到的公文元素只在报告中标记为 `not_detected` / `not_processed`。
- 能安全格式化的自动处理；不能安全处理的保留、诊断并报告。

报告会尽量给出：

- `content_preservation`：用文本指纹说明正文是否保持不变。
- `coverage`：说明哪些格式项已处理、已保留、仅诊断、未检测、暂不支持或需要人工确认。

## 一句话安装

把下面这句话发给支持 Skills 的 agent：

```text
请从这个 GitHub 地址安装 gongwen-format-converter skill：
https://github.com/danxiaozhiren/gongwen-format-converter-skill/tree/main/skills/gongwen-format-converter
```

安装完成后，重启 Codex / TRAE / 你的 agent 环境，让新 Skill 被重新发现。

## 安装方式

### 方式一：让 agent 自动安装

如果你的 agent 支持 `skill-installer`，直接让它安装这个 URL：

```text
https://github.com/danxiaozhiren/gongwen-format-converter-skill/tree/main/skills/gongwen-format-converter
```

对于 Codex，可以这样描述：

```text
使用 skill-installer 从 GitHub 安装：
https://github.com/danxiaozhiren/gongwen-format-converter-skill/tree/main/skills/gongwen-format-converter
```

### 方式二：命令行安装

如果你本地有 Codex 的 `skill-installer` 脚本，可以运行类似命令：

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --url "https://github.com/danxiaozhiren/gongwen-format-converter-skill/tree/main/skills/gongwen-format-converter"
```

安装位置通常是：

```text
%USERPROFILE%\.codex\skills\gongwen-format-converter
```

### 方式三：手动安装

也可以手动下载或克隆本仓库，然后把下面这个目录复制到你的 Codex skills 目录：

```text
skills/gongwen-format-converter
```

复制后的目标路径通常是：

```text
%USERPROFILE%\.codex\skills\gongwen-format-converter
```

目录里必须能直接看到：

```text
SKILL.md
references/
scripts/
evals/
agents/
```

## 快速验证

重启 agent 后，可以用下面的测试请求确认 Skill 是否被正确触发：

```text
帮我把这份 Markdown 简报整理成内部信息简报 Word 版式，内容不要改，只调整标题、期号、正文、一级标题、行距和缩进。
```

或：

```text
这个 docx 内容已经定稿了，只帮我按正式公文习惯调格式，不要润色正文。
```

## 推荐触发指令

安装后建议显式点名 Skill：

```text
$gongwen-format-converter

模式：现文格式化
输入文件：D:\path\材料.docx
预设：brief
要求：内容不要改、不补写，只调整已有内容和对象的格式，并给出覆盖报告。
```

如果你只给了一篇文章但没说怎么处理，Skill 应该先让你选择：

```text
1. 现文格式化：内容不改、不补写，只调整已有文字、段落、页面、表格、图片、页眉页脚等格式。
2. 现文格式诊断：不改文档，完整识别已有内容的页面、页眉页脚、段落角色、字体字号、颜色/下划线、缩进、行距、表格图片和对象格式状态。
3. 现文模板套用：用一份范文/模板的格式去套另一份文档，只学习样式，不复述模板内容，也不补目标文档缺失内容。
```

### 现文格式化

```text
$gongwen-format-converter

模式：现文格式化
输入文件：D:\path\材料.docx
预设：formal
要求：保留原文和段落顺序，只调整已有标题、正文、表格、页眉页脚等格式。
```

### 现文格式诊断

```text
$gongwen-format-converter

模式：现文格式诊断
输入文件：D:\path\材料.docx
要求：先不要生成新文档，完整诊断整篇文档的页面、页眉页脚、段落角色、字体字号、颜色/下划线、行距缩进、表格图片、对象、样式系统，以及与 formal/brief 预设的差异。
```

### 现文模板套用

```text
$gongwen-format-converter

模式：现文模板套用
模板文件：D:\path\模板.docx
目标文件：D:\path\新材料.docx
要求：先提取模板格式清单并列出目标文档已有内容中模板未覆盖的格式项，不要直接生成。确认后再套用模板。
```

如果模板里没有某类段落或某种样式，Skill 应该先询问你是使用推荐公文/简报格式、保留目标文档原格式，还是指定自定义格式；如果你要求直接继续，默认使用对应的公文/内部简报预设并在报告中说明。

模板复刻推荐流程：

1. 提取模板格式清单：页面、页眉页脚、标题、正文、各级标题、颜色、下划线、行距、缩进、表格图片等。
2. 对照目标文档，列出模板没有覆盖的已有格式项。
3. 询问未覆盖项处理方式：使用推荐公文/简报格式、保留目标原格式、或用户自定义。
4. 确认后生成套用模板后的 `.docx`，不复制模板正文、不补目标缺失内容。

脚本层也支持先提取模板 profile：

```powershell
python skills\gongwen-format-converter\scripts\format_document.py `
  --extract-template D:\path\模板.docx `
  --target D:\path\新材料.docx `
  --preset brief `
  --report template_profile.json
```

## 格式范围清单

正式公文格式范围优先参考《党政机关公文处理工作条例》和 `GB/T 9704-2012 党政机关公文格式`。Skill 会把格式分成这些范围来确认和处理：

| 范围 | 示例 |
| --- | --- |
| 页面级格式 | A4、横竖版、页边距、装订线、分节、分栏、版心、行数/字数参考、页眉页脚距离 |
| 段落级格式 | 对齐、首行缩进、左右缩进、固定行距、段前段后、大纲级别、分页控制、制表位、编号 |
| 文字级格式 | 中文/西文字体、字号、颜色、加粗、斜体、下划线、删除线、上下标、字符间距、高亮 |
| 已有结构角色 | 份号、密级、紧急程度、发文机关标志、发文字号、签发人、标题、主送机关、正文、附件、落款、成文日期、附注、抄送、印发机关和日期等已存在元素 |
| 表格格式 | 表格宽度、列宽、行高、单元格边距、边框、底纹、表内字体字号、跨页表头 |
| 图片和对象 | 图片、印章图片、文本框、形状、线条、图表、公式、嵌入对象 |
| 页眉页脚/版记 | 已有页码、页眉线、页脚信息、版记分隔线、抄送/印发区域格式 |
| 样式系统 | 段落样式、字符样式、标题样式、直接格式冲突、样式变体 |
| 特殊状态 | 批注、修订、域代码、目录、超链接、脚注尾注、书签、隐藏文字 |

处理优先级：

1. 用户明确指定的格式。
2. 用户提供的正式模板。
3. `formal` 正式公文预设。
4. `brief` 内部简报预设。
5. 仍不明确时先询问；若用户要求直接继续，则使用最接近的推荐格式并在报告中说明。

## Skill 位置

```text
skills/gongwen-format-converter/
```

## 核心模式

- 现文格式化模式：内容已定稿，只调整已有内容和对象的格式。
- 现文格式诊断模式：诊断页面、段落、文字、表格、图片、页眉页脚、样式系统和特殊状态等格式范围。
- 现文模板套用模式：从内部范文中提取样式特征，只学习格式，不复述或外泄范文正文，也不补目标文档缺失内容。

## 支持输入

- `.docx`
- `.md`
- `.txt`
- 粘贴文本
- 模板 `.docx` + 目标文档

主输出为 `.docx`，因为公文版式依赖字体、字号、页边距、固定行距、缩进等 Word 样式能力。

## 隐私原则

默认将用户提供的公文、简报、会议、经营、财务、人事等材料视为内部敏感内容：

- 不联网处理用户文档正文。
- 不在报告中默认输出完整原文。
- 模板复刻只提取样式指纹，不提取正文语义。
- 格式化模式默认不改写、不润色、不补写正文，也不补写缺失公文要素。

## 脚本依赖

脚本依赖 `python-docx`。如本地没有安装，可运行：

```powershell
python -m pip install -r skills\gongwen-format-converter\scripts\requirements.txt
```

## 规则沉淀

比赛规则和开发检查清单见：

```text
docs/solo-skill-contest-rules.md
```
