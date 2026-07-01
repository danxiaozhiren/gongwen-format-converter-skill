#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Check that supported coverage-matrix rows have executable proof links."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
COVERAGE = ROOT / "docs" / "word-format-coverage.md"
TESTS = ROOT / "tests" / "test_word_fixtures.py"
FIXTURES = ROOT / "tests" / "fixtures" / "word_samples"


COVERAGE_PROOFS: dict[str, dict[str, list[str]]] = {
    "纸张大小": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "页边距": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "版心参考": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "段落对齐": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "缩进": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "行距": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "段前段后": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "手写编号": {"fixtures": ["01-basic-formal.docx"], "tests": ["test_generated_samples_emit_expected_diagnostics"]},
    "中文字体": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "字号": {"tests": ["test_cli_default_uses_checklist_layout"]},
    "表内文字": {"tests": ["test_format_tables_preserves_table_text"]},
    "内容保持校验": {
        "tests": [
            "test_content_preservation_failure_withholds_output",
            "test_format_tables_preserves_table_text",
        ]
    },
}


def parse_matrix(path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    row_re = re.compile(r"^\|\s*(?P<area>[^|]+?)\s*\|\s*`(?P<status>[^`]+)`\s*\|")
    for line in path.read_text(encoding="utf-8").splitlines():
        match = row_re.match(line)
        if not match:
            continue
        area = match.group("area").strip()
        status = match.group("status").strip()
        if area == "范围":
            continue
        rows.append({"area": area, "status": status})
    return rows


def proof_errors(area: str, proof: dict[str, list[str]], test_source: str) -> list[str]:
    errors: list[str] = []
    for fixture in proof.get("fixtures", []):
        if not (FIXTURES / fixture).exists():
            errors.append(f"{area}: fixture does not exist: {fixture}")
    for test_symbol in proof.get("tests", []):
        if test_symbol not in test_source:
            errors.append(f"{area}: test symbol not found: {test_symbol}")
    if not proof.get("fixtures") and not proof.get("tests"):
        errors.append(f"{area}: proof has no fixtures or tests")
    return errors


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--coverage", default=str(COVERAGE), help="Coverage matrix markdown path.")
    parser.add_argument(
        "--statuses",
        nargs="+",
        default=["supported"],
        help="Coverage statuses that must have proof links.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = parse_matrix(Path(args.coverage))
    required_statuses = set(args.statuses)
    required = [row["area"] for row in rows if row["status"] in required_statuses]
    test_source = TESTS.read_text(encoding="utf-8")
    failures: list[str] = []
    for area in required:
        proof = COVERAGE_PROOFS.get(area)
        if proof is None:
            failures.append(f"{area}: no coverage proof registered")
            continue
        failures.extend(proof_errors(area, proof, test_source))
    unknown = sorted(set(COVERAGE_PROOFS) - set(row["area"] for row in rows))
    for area in unknown:
        failures.append(f"{area}: proof registered but area is absent from coverage matrix")
    if failures:
        print("Coverage matrix proof check failed:", file=sys.stderr)
        for failure in failures:
            print(f"- {failure}", file=sys.stderr)
        return 1
    print(f"Checked {len(required)} coverage rows for statuses: {', '.join(sorted(required_statuses))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
