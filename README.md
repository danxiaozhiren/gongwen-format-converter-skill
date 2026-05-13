# 公文/内部简报格式整理 Skill

这个仓库沉淀了一个面向 TRAE SOLO / Codex Skills 的公文格式处理 Skill：`gongwen-format-converter`。

它的核心目标是：在不改写正文内容的前提下，把 `.docx`、Markdown、纯文本或粘贴材料整理成符合公文/内部简报习惯的 Word 版式。

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

## Skill 位置

```text
skills/gongwen-format-converter/
```

## 核心模式

- 仅格式化模式：内容已定稿，只调整字体、字号、行距、缩进、页边距、标题层级等。
- 格式识别模式：识别主标题、期号、正文标题、一级标题、二级标题、正文、附件、落款、日期等段落角色。
- 模板复刻模式：从内部范文中提取样式特征，只学习格式，不复述或外泄范文正文。

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
- 格式化模式默认不改写、不润色、不补写正文。

## 规则沉淀

比赛规则和开发检查清单见：

```text
docs/solo-skill-contest-rules.md
```
