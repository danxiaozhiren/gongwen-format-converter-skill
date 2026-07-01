#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Optionally render generated Word fixtures to PDF with LibreOffice."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import pdfplumber
except ModuleNotFoundError:
    pdfplumber = None

try:
    from .generate_word_samples import OUTPUT_DIR, generate_all
except ImportError:
    from generate_word_samples import OUTPUT_DIR, generate_all


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_RENDER_DIR = Path("/private/tmp/gongwen-rendered-word-samples")
PDF_PAGE_RE = re.compile(rb"/Type\s*/Page\b")


def renderer_candidates() -> list[Path]:
    candidates = []
    env_renderer = os.environ.get("WORD_RENDERER")
    if env_renderer:
        candidates.append(Path(env_renderer))
    for binary in ["soffice", "libreoffice"]:
        found = shutil.which(binary)
        if found:
            candidates.append(Path(found))
    candidates.extend(
        [
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path("/Applications/OpenOffice.app/Contents/MacOS/soffice"),
        ]
    )
    seen: set[Path] = set()
    unique = []
    for candidate in candidates:
        resolved = candidate.expanduser()
        if resolved in seen:
            continue
        seen.add(resolved)
        if resolved.exists() or shutil.which(str(resolved)):
            unique.append(resolved)
    return unique


def find_renderer() -> Path | None:
    candidates = renderer_candidates()
    return candidates[0] if candidates else None


def render_pdf(renderer: Path, sample: Path, outdir: Path) -> Path:
    output = outdir / f"{sample.stem}.pdf"
    if output.exists():
        output.unlink()
    completed = subprocess.run(
        [
            str(renderer),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(sample),
        ],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"renderer failed for {sample.name}: {completed.stderr.strip() or completed.stdout.strip()}"
        )
    if not output.exists():
        raise RuntimeError(f"renderer did not create expected PDF: {output}")
    return output


def pdf_text_line_counts(path: Path) -> list[int]:
    if pdfplumber is None:
        return []
    counts = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [line for line in text.splitlines() if line.strip()]
            counts.append(len(lines))
    return counts


def validate_pdf(path: Path) -> dict[str, int | str | list[int]]:
    data = path.read_bytes()
    if not data.startswith(b"%PDF"):
        raise RuntimeError(f"{path.name} is not a PDF")
    if len(data) < 1000:
        raise RuntimeError(f"{path.name} is unexpectedly small: {len(data)} bytes")
    page_markers = len(PDF_PAGE_RE.findall(data))
    if page_markers < 1:
        raise RuntimeError(f"{path.name} has no detectable PDF page markers")
    line_counts = pdf_text_line_counts(path)
    if line_counts:
        if not any(count > 0 for count in line_counts):
            raise RuntimeError(f"{path.name} rendered without extractable text lines")
        if any(count > 35 for count in line_counts):
            raise RuntimeError(f"{path.name} has an unexpected rendered line count: {line_counts}")
    return {
        "file": path.name,
        "bytes": len(data),
        "page_markers": page_markers,
        "text_line_counts": line_counts,
        "layout_check": "pdfplumber" if pdfplumber is not None else "pdf-smoke-only",
    }


def renderer_bootstrap_failed(failures: list[str]) -> bool:
    if not failures:
        return False
    failure_text = "\n".join(failures)
    bootstrap_markers = [
        "Library not loaded",
        "cannot open shared object file",
        "libfontconfig",
        "dyld",
    ]
    return all(any(marker in failure for marker in bootstrap_markers) for failure in failures) and any(
        marker in failure_text for marker in bootstrap_markers
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-renderer", action="store_true", help="Fail when no renderer is installed.")
    parser.add_argument("--outdir", default=str(DEFAULT_RENDER_DIR), help="Directory for rendered PDFs.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    renderer = find_renderer()
    if renderer is None:
        message = (
            "SKIP: no LibreOffice/OpenOffice renderer found. Install LibreOffice or set WORD_RENDERER "
            "to enable render smoke checks."
        )
        if args.require_renderer:
            print(message, file=sys.stderr)
            return 2
        print(message)
        return 0

    samples = generate_all()
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    results = []
    failures = []
    for sample in samples:
        try:
            pdf = render_pdf(renderer, sample, outdir)
            results.append(validate_pdf(pdf))
        except Exception as exc:
            failures.append(f"{sample.name}: {exc}")

    if failures:
        if not args.require_renderer and renderer_bootstrap_failed(failures):
            print(
                "SKIP: LibreOffice/OpenOffice renderer was found but cannot start in this environment. "
                "Use make render-word-required to force this failure."
            )
            for failure in failures[:2]:
                print(f"- {failure}")
            return 0
        print("Word render smoke check failed:", file=sys.stderr)
        for failure in failures:
            print(f"- {failure}", file=sys.stderr)
        return 1

    print(f"Rendered and checked {len(results)} Word fixtures with {renderer}")
    print(f"Rendered PDFs: {outdir}")
    if pdfplumber is None:
        print("PDF text-line checks skipped because pdfplumber is not installed.")
    else:
        print("PDF text-line checks: " + ", ".join(f"{row['file']}={row['text_line_counts']}" for row in results))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
