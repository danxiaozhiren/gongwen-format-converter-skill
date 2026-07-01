"""Conservative paragraph role detection for existing Word documents."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


CHINESE_NUM = "一二三四五六七八九十"
DATE_RE = re.compile(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")
ISSUE_RE = re.compile(r"^第\s*[\d%s]+\s*期$" % CHINESE_NUM)
DOC_NUMBER_RE = re.compile(r"^[\u4e00-\u9fffA-Za-z0-9]{1,12}〔\d{4}〕\d{1,4}号$")
H1_RE = re.compile(r"^[%s]+、" % CHINESE_NUM)
H2_RE = re.compile(r"^（[%s]+）" % CHINESE_NUM)
H3_RE = re.compile(r"^\d+\s*[.．、]")
H4_RE = re.compile(r"^（\d+）")
SEPARATOR_RE = re.compile(r"^[\-_—=─━]{3,}$")
BULLET_RE = re.compile(r"^(?:[•·●○■□◆◇]\s*|-\s+)")
MANUAL_NUMBERING_PATTERNS = [
    ("chinese_level_1", H1_RE),
    ("chinese_level_2", H2_RE),
    ("arabic_level_3", H3_RE),
    ("paren_arabic", H4_RE),
    ("circled_arabic", re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩]")),
]
MANUAL_NUMBERING_LEVELS = {
    "chinese_level_1": 1,
    "chinese_level_2": 2,
    "arabic_level_3": 3,
    "paren_arabic": 4,
    "circled_arabic": 3,
}


@dataclass
class ParagraphItem:
    text: str
    role_hint: str | None = None
    paragraph: Any | None = None


@dataclass
class RoleCandidate:
    role: str
    confidence: float
    reason_codes: list[str]
    warnings: list[str]


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def manual_numbering_kind(text: str) -> str | None:
    normalized = normalize_text(text)
    for kind, pattern in MANUAL_NUMBERING_PATTERNS:
        if pattern.match(normalized):
            return kind
    return None


def manual_numbering_level(kind: str | None) -> int | None:
    if kind is None:
        return None
    return MANUAL_NUMBERING_LEVELS.get(kind)


def is_short_title_like(text: str) -> bool:
    if len(text) > 40:
        return False
    return not text.endswith(("。", "；", ";", "，", ","))


def is_recipient_like(text: str, index: int, total: int) -> bool:
    return text.endswith("：") and index < max(6, total // 4) and len(text) <= 30


def is_metadata_like(text: str) -> bool:
    return bool(
        DATE_RE.fullmatch(text)
        or ISSUE_RE.match(text)
        or DOC_NUMBER_RE.match(text)
        or re.search(r"签发人[:：]", text)
    )


def is_signature_like(text: str) -> bool:
    return bool(
        len(text) <= 30
        and re.search(r"(部|处|办|办公室|中心|公司|委员会|局|厅|院|所)$", text)
    )


def role_candidate(
    role: str,
    confidence: float,
    reason_codes: list[str],
    warnings: list[str] | None = None,
) -> RoleCandidate:
    return RoleCandidate(role, confidence, reason_codes, warnings or [])


def paragraph_role_candidates(
    item: ParagraphItem,
    index: int,
    total: int,
    preset: str,
    context: dict[str, Any],
) -> list[RoleCandidate]:
    text = normalize_text(item.text)

    if item.role_hint:
        return [role_candidate(item.role_hint, 1.0, ["explicit_role_hint"])]

    if not text:
        return [role_candidate("empty", 1.0, ["empty_paragraph"])]

    if SEPARATOR_RE.match(text):
        return [role_candidate("separator", 0.99, ["separator_pattern"])]

    if H1_RE.match(text):
        return [role_candidate("heading_1", 0.98, ["manual_numbering_chinese_level_1"])]
    if H2_RE.match(text):
        return [role_candidate("heading_2", 0.98, ["manual_numbering_chinese_level_2"])]
    if H3_RE.match(text):
        return [role_candidate("heading_3", 0.96, ["manual_numbering_arabic_level_3"])]
    if H4_RE.match(text):
        return [role_candidate("heading_4", 0.96, ["manual_numbering_parenthesized_arabic"])]

    if text.startswith("附件"):
        return [role_candidate("attachment", 0.96, ["attachment_prefix"])]

    candidates: list[RoleCandidate] = []

    if is_recipient_like(text, index, total):
        candidates.append(role_candidate("recipient", 0.93, ["early_full_width_colon_recipient"]))

    if index >= total - 3 and DATE_RE.fullmatch(text):
        candidates.append(role_candidate("date", 0.97, ["date_fullmatch_near_end"]))

    has_trailing_date_after = bool(context.get("has_trailing_date_after"))
    if index >= total - 4 and is_signature_like(text) and (has_trailing_date_after or index == total - 1):
        candidates.append(role_candidate("signature", 0.9, ["org_suffix_near_end_with_trailing_date"]))

    title_window = index < max(6, total // 4)
    if title_window and is_short_title_like(text) and not is_metadata_like(text) and not is_recipient_like(text, index, total):
        confidence = 0.88 if index <= 1 else 0.78
        candidates.append(role_candidate("main_title", confidence, ["early_title_like_non_metadata"]))

    if is_metadata_like(text):
        candidates.append(role_candidate("metadata", 0.88, ["metadata_pattern"]))

    if not candidates and is_short_title_like(text) and len(text) <= 18 and index < 6:
        candidates.append(
            role_candidate(
                "needs_review",
                0.74,
                ["low_confidence_short_title_like"],
                ["short title-like paragraph; treated as needs_review unless confirmed"],
            )
        )

    candidates.append(role_candidate("body", 0.55, ["default_body"]))
    return candidates


def select_role_candidate(candidates: list[RoleCandidate]) -> RoleCandidate:
    return max(candidates, key=lambda candidate: candidate.confidence)


def decide_role(
    candidates: list[RoleCandidate],
    selected_main_title_index: int | None,
    selected_article_title_index: int | None,
    index: int,
) -> RoleCandidate:
    filtered = []
    for candidate in candidates:
        if candidate.role == "main_title" and selected_main_title_index != index:
            continue
        if candidate.role == "article_title" and selected_article_title_index != index:
            continue
        filtered.append(candidate)
    return select_role_candidate(filtered or candidates)
