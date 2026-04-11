"""Form profile RPC handlers (Phase 5A/5B/5F).

- analyze_form : build_form_profile 호출 (rule 기반 양식 분석)
- detect_placeholders : 단독 placeholder 감지 (full_text 입력)
- mark_review_required : 표 type 기반 검토 필요 영역 자동 마킹
"""
from __future__ import annotations

import os
import sys
from typing import Any, Dict

from hwp_core import register
from hwp_core._helpers import validate_params, validate_file_path
from hwp_core._state import set_current_doc_path, get_current_doc_path


@register("analyze_form")
def analyze_form(hwp, params: Dict[str, Any]) -> Dict[str, Any]:
    """Phase 5A: 양식을 rule 기반으로 전체 분석 → form_profile.

    Params:
        file_path: 양식 파일 경로 (필수)
        include_placeholders: bool (default True)
        include_guidance: bool (default True)
        summary_only: bool (default False) — 요약만 반환 (profile 전체 생략)

    Returns:
        full profile (summary_only=False) 또는 summary dict
    """
    validate_params(params, ["file_path"], "analyze_form")
    file_path = validate_file_path(params["file_path"], must_exist=True)
    include_placeholders = bool(params.get("include_placeholders", True))
    include_guidance = bool(params.get("include_guidance", True))
    summary_only = bool(params.get("summary_only", False))

    # 양식이 이미 열려있는지 확인 — 아니면 open
    current = get_current_doc_path()
    need_open = current != os.path.abspath(file_path)
    if need_open:
        try:
            hwp.open(os.path.abspath(file_path))
            set_current_doc_path(os.path.abspath(file_path))
        except Exception as e:
            print(f"[WARN] analyze_form open: {e}", file=sys.stderr)

    from hwp_analyzer.form_profile import build_form_profile, summarize_profile

    profile = build_form_profile(
        hwp,
        file_path,
        include_placeholders=include_placeholders,
        include_guidance=include_guidance,
    )

    result: Dict[str, Any] = {
        "status": "ok",
        "summary": summarize_profile(profile),
    }
    if not summary_only:
        result["profile"] = profile
    return result


@register("detect_placeholders")
def detect_placeholders_handler(hwp, params: Dict[str, Any]) -> Dict[str, Any]:
    """Phase 5B: 양식의 placeholder 만 자동 감지 (full_text 기반).

    Params:
        file_path: optional (생략 시 현재 열린 문서)

    Returns:
        {status, placeholders, placeholders_count, primary_marker}
    """
    from hwp_analyzer.placeholder_detector import detect_placeholders, detect_primary_marker

    file_path = params.get("file_path")
    if file_path:
        file_path = validate_file_path(file_path, must_exist=True)
        current = get_current_doc_path()
        if current != os.path.abspath(file_path):
            hwp.open(os.path.abspath(file_path))
            set_current_doc_path(os.path.abspath(file_path))

    try:
        full_text = hwp.get_text_file("TEXT", "") or ""
    except Exception as e:
        print(f"[WARN] detect_placeholders get_text: {e}", file=sys.stderr)
        full_text = ""

    markers = detect_primary_marker(full_text)
    placeholders = detect_placeholders(full_text, primary_marker=markers["primary"])

    # type별 count
    type_counts: Dict[str, int] = {}
    for p in placeholders:
        t = p.get("type", "unknown")
        type_counts[t] = type_counts.get(t, 0) + 1

    return {
        "status": "ok",
        "placeholders": placeholders,
        "placeholders_count": len(placeholders),
        "type_counts": type_counts,
        "primary_marker": markers["primary"],
        "marker_counts": markers["counts"],
    }


# Phase 5F: mark_review_required — 실제 데이터 필요 영역 자동 감지
_REVIEW_REQUIRED_TYPES = {"patent", "market_size", "sales_plan", "equipment", "revenue"}


@register("mark_review_required")
def mark_review_required(hwp, params: Dict[str, Any]) -> Dict[str, Any]:
    """Phase 5F: 표 type 기반 실제 데이터 필요 영역 식별.

    `form_profile` 의 tables list 를 받아서 type 이 _REVIEW_REQUIRED_TYPES 인
    표들을 식별. cell 에 `[!검토필요]` prefix + 빨간색 marker 추가 (별도 단계).

    Params:
        profile: form_profile dict (analyze_form 결과)

    Returns:
        {review_required_tables: [...], skipped_tables: [...]}
    """
    validate_params(params, ["profile"], "mark_review_required")
    profile = params["profile"]
    tables = profile.get("tables", []) if isinstance(profile, dict) else []

    required: list = []
    skipped: list = []
    for t in tables:
        ttype = t.get("type", "unknown")
        if ttype in _REVIEW_REQUIRED_TYPES:
            required.append({
                "index": t.get("index"),
                "type": ttype,
                "header_text": t.get("header_text", ""),
                "reason": f"{ttype} 표는 실제 사용자 데이터 필요",
            })
        else:
            skipped.append({
                "index": t.get("index"),
                "type": ttype,
            })

    return {
        "status": "ok",
        "review_required_tables": required,
        "review_required_count": len(required),
        "skipped_tables": skipped,
        "skipped_count": len(skipped),
    }
