#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Unittest entrypoint for Word fixture smoke checks."""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest

from docx import Document
from docx.oxml.ns import qn

from tests.fixtures import check_word_samples


ROOT = Path(__file__).resolve().parents[1]
FORMATTER = ROOT / "skills" / "gongwen-format-converter" / "scripts" / "format_document.py"


def run_font_attrs(run):
    rpr = run._element.rPr
    rfonts = rpr.rFonts if rpr is not None else None
    if rfonts is None:
        return {}
    return {
        "eastAsia": rfonts.get(qn("w:eastAsia")),
        "ascii": rfonts.get(qn("w:ascii")),
        "hAnsi": rfonts.get(qn("w:hAnsi")),
    }


class WordFixtureSmokeTest(unittest.TestCase):
    def test_generated_samples_emit_expected_diagnostics(self) -> None:
        self.assertEqual(check_word_samples.main(), 0)

    def test_formal_default_uses_checklist_template(self) -> None:
        with TemporaryDirectory(prefix="gongwen-default-template-") as temp_dir:
            temp_path = Path(temp_dir)
            source = temp_path / "default-template.txt"
            output = temp_path / "default-template.docx"
            report = temp_path / "default-template.json"
            source.write_text(
                "\n".join(
                    [
                        "关于开展默认模板检查工作的通知",
                        "各部门：",
                        "正文内容包含 ABC123，用于检查数字和字母字体。",
                        "一、一级标题",
                        "（一）二级标题",
                        "1. 三级标题",
                        "（1）四级标题",
                    ]
                ),
                encoding="utf-8",
            )

            subprocess.run(
                [
                    sys.executable,
                    str(FORMATTER),
                    str(source),
                    "--output",
                    str(output),
                    "--preset",
                    "formal",
                    "--add-page-numbers",
                    "--report",
                    str(report),
                ],
                cwd=ROOT,
                check=True,
            )

            data = json.loads(report.read_text(encoding="utf-8"))
            self.assertEqual(data["role_counts"]["heading_4"], 1)

            document = Document(str(output))
            section = document.sections[0]
            self.assertAlmostEqual(section.top_margin.mm, 37, places=1)
            self.assertAlmostEqual(section.bottom_margin.mm, 35, places=1)
            self.assertAlmostEqual(section.left_margin.mm, 27, places=1)
            self.assertAlmostEqual(section.right_margin.mm, 27, places=1)

            paragraphs = {paragraph.text: paragraph for paragraph in document.paragraphs if paragraph.text}
            title = paragraphs["关于开展默认模板检查工作的通知"]
            self.assertAlmostEqual(title.paragraph_format.line_spacing.pt, 36, places=1)
            self.assertAlmostEqual(title.paragraph_format.space_after.pt, 30, places=1)
            title_run = title.runs[0]
            self.assertEqual(title_run.font.size.pt, 22)
            self.assertFalse(title_run.font.bold)
            self.assertEqual(run_font_attrs(title_run)["eastAsia"], "方正小标宋简体")
            self.assertEqual(run_font_attrs(title_run)["ascii"], "Times New Roman")

            body_run = paragraphs["正文内容包含 ABC123，用于检查数字和字母字体。"].runs[0]
            self.assertAlmostEqual(
                paragraphs["正文内容包含 ABC123，用于检查数字和字母字体。"].paragraph_format.line_spacing.pt,
                30,
                places=1,
            )
            self.assertEqual(run_font_attrs(body_run)["eastAsia"], "仿宋_GB2312")
            self.assertEqual(run_font_attrs(body_run)["ascii"], "Times New Roman")

            heading_2_run = paragraphs["（一）二级标题"].runs[0]
            self.assertTrue(heading_2_run.font.bold)
            self.assertEqual(run_font_attrs(heading_2_run)["eastAsia"], "楷体_GB2312")

            heading_4 = paragraphs["（1）四级标题"]
            self.assertAlmostEqual(heading_4.paragraph_format.line_spacing.pt, 30, places=1)
            self.assertEqual(run_font_attrs(heading_4.runs[0])["eastAsia"], "仿宋_GB2312")

            footer_run = section.footer.paragraphs[0].runs[0]
            self.assertEqual(footer_run.font.size.pt, 14)
            self.assertEqual(run_font_attrs(footer_run)["eastAsia"], "宋体")


if __name__ == "__main__":
    unittest.main()
