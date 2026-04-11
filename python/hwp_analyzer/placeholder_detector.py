"""Placeholder detector — rule 기반 양식 placeholder 자동 감지 (Phase 5B).

사용자 핵심 원칙 (v0.7.10+):
> "수천개의 양식을 모두 학습할 수 없기에 그때그때 양식에 따라서 달라짐을 알아야합니다"

특정 양식 hardcode 없이, rule 만으로 임의 양식의 placeholder 를 감지.

v0.8.2 CRITICAL FIX:
이전 버전은 R3/R4/R5 를 placeholder 로 분류했으나, 실제로 이들은 양식의 **section heading**.
v0.8.1 E2E 에서 cleanup_all_placeholders 가 `(1) 산업의 특성`, `1) 투자계획`, `가. 추진배경`,
`1. 사업 개요` 같은 **양식의 핵심 heading 을 대량 삭제** → 문서 구조 파괴.

v0.8.2 재분류:
  **진짜 placeholder** (cleanup 대상 — 비어있는 template):
  - R1 bare_marker  : 단독 marker (`◦`) — cleanup OK
  - R2 colon_label  : `◦ 안정성 :` (label + 빈 값) — cleanup OK
  - R6 blank        : 빈 paragraph — 유지 (fill_empty_line 이 사용)

  **section heading** (cleanup 금지 — 양식의 구조):
  - R3 bracket_heading  : `[과학⋅기술적 측면]` — 양식 sub-section
  - R4 numbered_section : `(1) 산업의 특성`, `1) 투자계획`, `가. 추진배경`
  - R5 marker_section   : `◦ 매출증대 및 비용 절감 효과` — 양식 sub-bullet

  모두 감지되지만 type 이 다르게 분류되어 cleanup 에서 **default 로 제외됨**.

Returns structured list — {type, line, text, raw_line}.
"""
from __future__ import annotations

import re
from typing import List, Dict, Any


# 양식에 등장 가능한 모든 marker char (pattern 기반 감지용)
# 특정 양식 hardcode 아닌 "bullet 계열 char 집합"
BULLET_MARKERS = "◦○●○-*▪∙·"

# Number prefix patterns (rule 기반, 양식별 variants 커버)
_NUMBER_PREFIX = re.compile(
    r"^\s*(?:"
    r"\(\d+\)|"          # (1) (2)
    r"\d+\)|"             # 1) 2)
    r"\d+\.|"             # 1. 2.
    r"[①-⑳]|"           # ① ②
    r"[가-힣]\.|"       # 가. 나.
    r"\([가-힣]\)|"     # (가) (나)
    r"[A-Za-z]\.|"       # A. a.
    r"[A-Za-z]\)"         # A) a)
    r")\s*"
)

# Bracket heading patterns — []/【】/〈〉/《》 variants
_BRACKET_OPEN = "[【〈《『「"
_BRACKET_CLOSE = "]】〉》』」"

# Label 구분자 — halfwidth/fullwidth colon + dash
_LABEL_SEP = "::-"

# 단일 marker only (R1)
_R1_RE = re.compile(rf"^\s*[{re.escape(BULLET_MARKERS)}]\s*$")

# R3 bracket heading: ≤30 char, bracket 으로 감싸인 짧은 한글/영문/기호 label
_R3_RE = re.compile(
    rf"^\s*[{re.escape(_BRACKET_OPEN)}]"
    r"[\w가-힣·⋅․\-\s]{1,30}"
    rf"[{re.escape(_BRACKET_CLOSE)}]\s*$"
)


def _is_colon_label(stripped: str, markers: str = BULLET_MARKERS) -> bool:
    """R2: marker + 짧은 label + `:` 끝"""
    if not stripped:
        return False
    if stripped[0] not in markers:
        return False
    # marker + 공백 이후 label ≤15 + `:`/`：` 로 끝나는 형태
    rest = stripped[1:].strip()
    if not rest:
        return False
    if not (rest.endswith(":") or rest.endswith("：")):
        return False
    label = rest.rstrip(":：").strip()
    if len(label) > 15:
        return False
    return True


def _is_number_short(stripped: str, max_label_len: int = 15) -> bool:
    """R4: number prefix + 짧은 label (≤15)"""
    m = _NUMBER_PREFIX.match(stripped)
    if not m:
        return False
    rest = stripped[m.end():].strip()
    return 0 < len(rest) <= max_label_len


def _is_marker_short(stripped: str, markers: str = BULLET_MARKERS, max_len: int = 30) -> bool:
    """R5: marker + 짧은 text (≤30), colon 없이"""
    if not stripped or stripped[0] not in markers:
        return False
    if stripped.endswith(":") or stripped.endswith("："):
        return False  # R2 에서 처리
    rest = stripped[1:].strip()
    return 0 < len(rest) <= max_len


def detect_placeholders(
    full_text: str,
    primary_marker: str = "◦",
) -> List[Dict[str, Any]]:
    """Rule 기반 placeholder 자동 감지.

    Args:
        full_text: 양식의 전체 텍스트 (hwp.get_text_file / analyze_document 의 full_text)
        primary_marker: 양식의 주 bullet marker (form_profile.markers.primary 에서 결정)

    Returns:
        List of placeholder dicts — {type, line, text, raw_line}
    """
    placeholders: List[Dict[str, Any]] = []

    for line_no, raw_line in enumerate(full_text.split("\n"), 1):
        stripped = raw_line.strip()

        # R6 blank
        if not stripped:
            placeholders.append({
                "type": "blank",
                "line": line_no,
                "text": "",
                "raw_line": raw_line,
            })
            continue

        # R1 bare_marker (단독 marker)
        if _R1_RE.match(raw_line):
            placeholders.append({
                "type": "bare_marker",
                "line": line_no,
                "text": stripped,
                "raw_line": raw_line,
            })
            continue

        # R2 colon_label
        if _is_colon_label(stripped):
            placeholders.append({
                "type": "colon_label",
                "line": line_no,
                "text": stripped,
                "raw_line": raw_line,
            })
            continue

        # R3 bracket_heading (section heading, NOT placeholder — v0.8.2 fix)
        if _R3_RE.match(raw_line):
            placeholders.append({
                "type": "bracket_heading",  # 감지만, cleanup 대상 아님
                "line": line_no,
                "text": stripped,
                "raw_line": raw_line,
            })
            continue

        # R4 numbered_section (section heading, NOT placeholder — v0.8.2 fix)
        if _is_number_short(stripped):
            placeholders.append({
                "type": "numbered_section",  # v0.8.1: number_short → v0.8.2: numbered_section
                "line": line_no,
                "text": stripped,
                "raw_line": raw_line,
            })
            continue

        # R5 marker_section (section heading / sub-bullet, NOT placeholder — v0.8.2 fix)
        if _is_marker_short(stripped):
            placeholders.append({
                "type": "marker_section",  # v0.8.1: marker_short → v0.8.2: marker_section
                "line": line_no,
                "text": stripped,
                "raw_line": raw_line,
            })
            continue

    return placeholders


_PRIMARY_MARKER_RE = re.compile(
    rf"^[{re.escape(BULLET_MARKERS)}]\s+\S"
)


def detect_primary_marker(full_text: str) -> Dict[str, Any]:
    """양식의 주 bullet marker 를 빈도 분석으로 결정.

    v0.8.1 P3: 첫 글자만 check 에서 **marker + 공백 + 실제 content** pattern 으로 강화.
    `"◦ 본 사업"` 같이 marker 가 실제 bullet item 으로 사용되는 경우만 count.
    `"◦"` 단독 (bare_marker) 또는 `"---"` 같은 구분자는 primary marker 통계에서 제외.

    Returns:
        {primary, counts} — primary 는 가장 많이 사용된 marker, counts 는 각 marker 빈도
    """
    counts: Dict[str, int] = {}
    for line in full_text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        if _PRIMARY_MARKER_RE.match(stripped):
            first = stripped[0]
            counts[first] = counts.get(first, 0) + 1

    # Sort by count desc
    sorted_markers = sorted(counts.items(), key=lambda x: -x[1])
    primary = sorted_markers[0][0] if sorted_markers else "◦"
    return {
        "primary": primary,
        "counts": counts,
    }
