#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Compare key WordprocessingML nodes against a golden baseline."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET

from docx import Document


ROOT = Path(__file__).resolve().parents[2]
FORMATTER = ROOT / "skills" / "format-xingwen-word" / "scripts" / "format_document.py"
GOLDEN = ROOT / "tests" / "golden" / "basic_formal_golden.json"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

SOURCE_TEXT = "\n".join(
    [
        "关于开展黄金样例检查工作的通知",
        "各部门：",
        "正文内容包含 ABC123，用于检查数字和字母字体。",
        "一、一级标题",
        "（一）二级标题",
        "1. 三级标题",
    ]
)


def w_attr(element: ET.Element | None, name: str) -> str | None:
    if element is None:
        return None
    return element.get(f"{{{W_NS}}}{name}")


def text_for_paragraph(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def first_run_props(paragraph: ET.Element) -> dict[str, str | bool | None]:
    run = paragraph.find("./w:r", NS)
    rpr = run.find("./w:rPr", NS) if run is not None else None
    rfonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
    sz = rpr.find("./w:sz", NS) if rpr is not None else None
    bold = rpr.find("./w:b", NS) if rpr is not None else None
    bold_val = w_attr(bold, "val")
    return {
        "eastAsia": w_attr(rfonts, "eastAsia"),
        "ascii": w_attr(rfonts, "ascii"),
        "hAnsi": w_attr(rfonts, "hAnsi"),
        "size_half_points": w_attr(sz, "val"),
        "bold": bool(bold is not None and bold_val not in {"0", "false", "False"}),
    }


def paragraph_props(paragraph: ET.Element) -> dict[str, str | bool | None]:
    ppr = paragraph.find("./w:pPr", NS)
    jc = ppr.find("./w:jc", NS) if ppr is not None else None
    spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
    ind = ppr.find("./w:ind", NS) if ppr is not None else None
    outline = ppr.find("./w:outlineLvl", NS) if ppr is not None else None
    return {
        "align": w_attr(jc, "val"),
        "line": w_attr(spacing, "line"),
        "space_after": w_attr(spacing, "after"),
        "first_line_indent": w_attr(ind, "firstLine"),
        "outline_level": w_attr(outline, "val"),
        "run": first_run_props(paragraph),
    }


def extract_docx_signature(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        document_xml = ET.fromstring(package.read("word/document.xml"))
        settings_xml = ET.fromstring(package.read("word/settings.xml"))
    sect_pr = document_xml.find(".//w:sectPr", NS)
    pg_sz = sect_pr.find("./w:pgSz", NS) if sect_pr is not None else None
    pg_mar = sect_pr.find("./w:pgMar", NS) if sect_pr is not None else None
    doc_grid = sect_pr.find("./w:docGrid", NS) if sect_pr is not None else None
    if doc_grid is None:
        doc_grid = settings_xml.find(".//w:docGrid", NS)
    paragraphs = []
    for paragraph in document_xml.findall(".//w:body/w:p", NS):
        text = text_for_paragraph(paragraph)
        if not text:
            continue
        row = {"text": text}
        row.update(paragraph_props(paragraph))
        paragraphs.append(row)
    return {
        "page": {
            "width": w_attr(pg_sz, "w"),
            "height": w_attr(pg_sz, "h"),
            "top": w_attr(pg_mar, "top"),
            "bottom": w_attr(pg_mar, "bottom"),
            "left": w_attr(pg_mar, "left"),
            "right": w_attr(pg_mar, "right"),
        },
        "document_grid": {
            "type": w_attr(doc_grid, "type"),
            "line_pitch": w_attr(doc_grid, "linePitch"),
            "char_space": w_attr(doc_grid, "charSpace"),
        },
        "paragraphs": paragraphs,
    }


def formatted_signature(temp_dir: Path) -> dict[str, object]:
    source = temp_dir / "basic-formal.docx"
    output = temp_dir / "basic-formal-output.docx"
    report = temp_dir / "basic-formal.json"
    document = Document()
    for paragraph in SOURCE_TEXT.splitlines():
        document.add_paragraph(paragraph)
    document.save(str(source))
    subprocess.run(
        [
            sys.executable,
            str(FORMATTER),
            str(source),
            "--output",
            str(output),
            "--report",
            str(report),
        ],
        cwd=ROOT,
        check=True,
    )
    return extract_docx_signature(output)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--golden", default=str(GOLDEN), help="Golden JSON baseline path.")
    parser.add_argument("--dump", action="store_true", help="Print the current extracted signature.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    with TemporaryDirectory(prefix="gongwen-golden-") as temp:
        actual = formatted_signature(Path(temp))
    if args.dump:
        print(json.dumps(actual, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    golden_path = Path(args.golden)
    expected = json.loads(golden_path.read_text(encoding="utf-8"))
    if actual != expected:
        print("Golden XML check failed.", file=sys.stderr)
        print("Expected:", json.dumps(expected, ensure_ascii=False, indent=2, sort_keys=True), file=sys.stderr)
        print("Actual:", json.dumps(actual, ensure_ascii=False, indent=2, sort_keys=True), file=sys.stderr)
        return 1
    print(f"Golden XML check passed: {golden_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
