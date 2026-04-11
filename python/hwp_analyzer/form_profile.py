"""Form profile — 임의 HWP 양식을 rule 기반으로 전체 분석 (Phase 5A).

사용자 핵심 원칙 (v0.7.10+):
> "수천개의 양식을 모두 학습할 수 없기에 그때그때 양식에 따라서 달라짐을 알아야합니다"

특정 양식 hardcode 없이, 임의 양식에서:
- 페이지별 sections
- 모든 표 (index, dimensions, type, header/data cells)
- 작성요령 (max_pages, required, format_hints)
- 마커 종류 + 빈도 (primary bullet 자동 결정)
- Placeholder 자동 감지 (R1-R6 rule)
- 안내문 (* 현장적용 같은 가이드)

결과: `form_profile` dict (또는 JSON 저장). `form_fill_auto` 의 입력으로 사용.

사용 예:
    from hwp_analyzer.form_profile import build_form_profile
    profile = build_form_profile(hwp, "양식.hwp")
    # {sections, tables, guides, markers, placeholders, guidance_texts, pages}
"""
from __future__ import annotations

import os
import re
import sys
from typing import Any, Dict, List, Optional

from hwp_analyzer.placeholder_detector import (
    detect_placeholders,
    detect_primary_marker,
)


def _safe_get_text(hwp) -> str:
    """양식의 full text 추출 (analyze_document 과 다른 실시간 경로)."""
    try:
        return hwp.get_text_file("TEXT", "") or ""
    except Exception as e:
        print(f"[WARN] form_profile get_text_file: {e}", file=sys.stderr)
        return ""


def _extract_sections(full_text: str) -> List[Dict[str, Any]]:
    """v0.8.1 P1 fix: full_text line 기반 heading level 감지 (rule 기반).

    v0.8.0 `analyze_document` 는 `paragraphs` 를 반환하지 않음 (full_text 만).
    → full_text 를 line 단위로 순회하면서 `_heading_level` rule 적용.

    Returns: [{heading, level, line}, ...]
    """
    sections: List[Dict[str, Any]] = []
    if not full_text:
        return sections

    for line_no, raw_line in enumerate(full_text.split("\n"), 1):
        text = raw_line.strip()
        if not text:
            continue
        level = _heading_level(text)
        if level >= 1:
            sections.append({
                "heading": text,
                "level": level,
                "line": line_no,
            })
    return sections


# Heading level rule — v0.7.7 의 _detect_heading_depth 와 일치
_HEADING_PATTERNS = [
    (re.compile(r"^\s*\d+\.\s"), 1),              # "1. 사업 개요"
    (re.compile(r"^\s*[가-힣]\.\s"), 2),           # "가. 추진배경"
    (re.compile(r"^\s*\(\d+\)\s"), 3),             # "(1) 산업의 특성"
    (re.compile(r"^\s*\([가-힣]\)\s"), 3),         # "(가) 주시장"
    (re.compile(r"^\s*[①-⑳]"), 4),                 # ①
]


def _heading_level(text: str) -> int:
    """Rule 기반 heading level 감지. 0 = heading 아님."""
    for pat, lvl in _HEADING_PATTERNS:
        if pat.match(text):
            return lvl
    return 0


def _extract_tables(hwp, file_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """모든 표의 구조 분석 (rule 기반 type 분류).

    Returns: [{index, rows, cols, type, header_text, cell_count}, ...]
    """
    tables: List[Dict[str, Any]] = []

    # 기존 도구 재사용
    try:
        from hwp_analyzer.document import analyze_document as _analyze
        from hwp_analyzer.label import classify_table_type
    except Exception as e:
        print(f"[WARN] form_profile _extract_tables import: {e}", file=sys.stderr)
        return []

    try:
        result = _analyze(hwp, file_path or "", already_open=True)
    except Exception as e:
        print(f"[WARN] form_profile _extract_tables analyze: {e}", file=sys.stderr)
        return []

    raw_tables = result.get("tables", []) if isinstance(result, dict) else []
    for t in raw_tables:
        if not isinstance(t, dict):
            continue
        headers = t.get("headers", []) or []
        header_text = " | ".join(str(h) for h in headers[:5]) if headers else ""

        # v0.8.1 P2 fix: classify_table_type 은 table_info dict 를 받음 (list 아님).
        # 이전 코드는 `classify_table_type(headers)` (list 전달) → 내부 get("headers")
        # 가 list method 호출 실패 → unknown 반환.
        try:
            ttype = classify_table_type(t)
        except Exception as e:
            print(f"[WARN] classify_table_type: {e}", file=sys.stderr)
            ttype = "unknown"

        tables.append({
            "index": t.get("index", -1),
            "rows": t.get("rows", 0),
            "cols": t.get("cols", 0),
            "type": ttype,
            "header_text": header_text[:80],
        })

    return tables


def _extract_guidance_texts(full_text: str) -> List[Dict[str, Any]]:
    """양식의 본문 가이드 (`* 현장적용...`, `※ 주의사항...`) 자동 감지.

    Rule:
    - `*` 또는 `※` 로 시작하는 line
    - 한글/영문 포함, 10 char 이상
    """
    guidance: List[Dict[str, Any]] = []
    for line_no, line in enumerate(full_text.split("\n"), 1):
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("*") or stripped.startswith("※"):
            if len(stripped) >= 10:
                guidance.append({
                    "line": line_no,
                    "text": stripped[:200],
                })
    return guidance


def build_form_profile(
    hwp,
    file_path: str,
    include_placeholders: bool = True,
    include_guidance: bool = True,
) -> Dict[str, Any]:
    """임의 양식에서 `form_profile` 전체 생성.

    Args:
        hwp: pyhwpx Hwp 인스턴스 (이미 open_document 된 상태)
        file_path: 양식 파일 경로
        include_placeholders: placeholder 감지 여부
        include_guidance: 양식 안내문 감지 여부

    Returns:
        profile dict — sections/tables/guides/markers/placeholders/guidance_texts/pages
    """
    profile: Dict[str, Any] = {
        "form_file": os.path.basename(file_path) if file_path else "",
        "form_path": os.path.abspath(file_path) if file_path else "",
        "pages": 0,
        "sections": [],
        "tables": [],
        "guides": [],
        "markers": {},
        "placeholders": [],
        "guidance_texts": [],
    }

    # pages
    try:
        profile["pages"] = int(hwp.PageCount)
    except Exception as e:
        print(f"[WARN] form_profile pages: {e}", file=sys.stderr)

    # full_text (실시간 get_text_file)
    full_text = _safe_get_text(hwp)

    # v0.8.1 P1: sections — full_text line 단위로 rule 기반 heading 감지
    try:
        profile["sections"] = _extract_sections(full_text)
    except Exception as e:
        print(f"[WARN] form_profile sections: {e}", file=sys.stderr)

    # tables (classify_table_type 활용)
    try:
        profile["tables"] = _extract_tables(hwp, file_path)
    except Exception as e:
        print(f"[WARN] form_profile tables: {e}", file=sys.stderr)

    # guides (extract_guide_text 활용 — 기존 도구)
    try:
        from hwp_core.content import extract_guide_text as _eg_handler
        eg_result = _eg_handler(hwp, {"file_path": file_path})
        profile["guides"] = eg_result.get("guides", []) if isinstance(eg_result, dict) else []
    except Exception as e:
        print(f"[WARN] form_profile guides: {e}", file=sys.stderr)

    # markers (primary bullet 자동 결정)
    if full_text:
        profile["markers"] = detect_primary_marker(full_text)

    # placeholders (rule 기반 R1-R6)
    if include_placeholders and full_text:
        primary = profile["markers"].get("primary", "◦")
        profile["placeholders"] = detect_placeholders(full_text, primary_marker=primary)

    # guidance_texts (* / ※ 로 시작하는 안내문)
    if include_guidance and full_text:
        profile["guidance_texts"] = _extract_guidance_texts(full_text)

    return profile


def summarize_profile(profile: Dict[str, Any]) -> Dict[str, Any]:
    """`form_profile` 요약 — 사용자 / LLM 에게 보여줄 간략 정보.

    Returns: 통계 (pages/sections/tables/placeholders 수 + 주요 type)
    """
    tables = profile.get("tables", [])
    table_types: Dict[str, int] = {}
    for t in tables:
        ttype = t.get("type", "unknown")
        table_types[ttype] = table_types.get(ttype, 0) + 1

    placeholders = profile.get("placeholders", [])
    placeholder_types: Dict[str, int] = {}
    for p in placeholders:
        ptype = p.get("type", "unknown")
        placeholder_types[ptype] = placeholder_types.get(ptype, 0) + 1

    return {
        "pages": profile.get("pages", 0),
        "sections_count": len(profile.get("sections", [])),
        "tables_count": len(tables),
        "table_types": table_types,
        "guides_count": len(profile.get("guides", [])),
        "primary_marker": profile.get("markers", {}).get("primary", "?"),
        "marker_counts": profile.get("markers", {}).get("counts", {}),
        "placeholders_count": len(placeholders),
        "placeholder_types": placeholder_types,
        "guidance_texts_count": len(profile.get("guidance_texts", [])),
    }
