# Word 样例集

本目录用于维护最小 `.docx` 覆盖样例。样例服务于 `docs/word-format-coverage.md`，用于验证 `format-xingwen-word` 对常见 Word 格式面的诊断、保留和安全修改能力。

## 生成方式

从仓库根目录运行：

```bash
.venv/bin/python tests/fixtures/generate_word_samples.py
```

脚本会生成：

```text
tests/fixtures/word_samples/
```

也可以运行 smoke check，重新生成样例并用格式化器诊断报告断言关键结构：

```bash
.venv/bin/python tests/fixtures/check_word_samples.py
```

项目根目录还提供了标准测试入口：

```bash
make check
make skill-validate
make smoke-word
make evals
make coverage-matrix
make golden-word
make render-word
make test
```

`make evals` 会运行 `evals.json` 中的 formatter、formatter_mutation 和 mock-agent-policy 用例，确保 prompt 级行为和报告信号不悬空。

`make coverage-matrix` 会扫描 `docs/word-format-coverage.md`，要求所有 `supported` 行都有 fixture 或 unittest 证明。

`make golden-word` 会格式化固定输入，解包 `.docx` 并比对关键 XML 节点，防止字体、行距、缩进、页边距和大纲级别漂移。

`make render-word` 会在检测到可用 LibreOffice/OpenOffice 时把样例导出为 PDF，并做 PDF 有效性检查；安装 `pdfplumber` 时会额外提取每页文本行数。未检测到渲染器或当前环境中的渲染器无法启动时默认跳过。可以通过 `WORD_RENDERER=/path/to/soffice make render-word` 指定渲染器，或用 `make render-word-required` 强制失败。

等价的 Python 标准库测试命令：

```bash
.venv/bin/python -m unittest tests.test_word_fixtures
```

## 样例清单

| 文件 | 覆盖目标 |
| --- | --- |
| `01-basic-formal.docx` | 行文检查版式基础样例：标题、主送、正文、三级标题、附件、落款、日期。 |
| `02-table.docx` | 表格文字、表格边框、合并单元格诊断。 |
| `03-header-footer-page.docx` | 页眉、页脚、PAGE/NUMPAGES 字段、首页不同、多节文档。 |
| `04-image-seal.docx` | 图片/印章类 inline shape 的诊断和保留。 |
| `05-textbox-shape.docx` | 文本框/形状 OOXML 的存在性诊断。 |
| `06-auto-numbering.docx` | Word 自动编号和手写编号的区分。 |
| `07-toc-fields.docx` | TOC、DATE、REF、PAGEREF 字段，书签和标题结构。 |
| `08-comments-revisions.docx` | 批注、脚注、尾注关系，脚注/尾注引用，修订插入/删除标记。 |

## 验收原则

- 样例正文不追求业务真实性，只覆盖 Word 结构和格式边界。
- 格式化结果必须保持正文和表格文本顺序不变。
- 高风险对象先以诊断和保留为目标，不把样例存在性误当作自动修改能力。
- 后续新增 Word 能力时，先补样例，再补脚本行为和报告断言。
