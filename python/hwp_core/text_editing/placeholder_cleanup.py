"""Placeholder cleanup — rule 기반 동적 cleanup (Phase 5C).

v5 hotfix 에서 발견한 잔존 2 문제 해결:
1. 빈 `◦` × 8 — 한컴 find_replace regex 모드 미지원 → paragraph iterate + DeletePara
2. `[경제적․사회적 측면]` × 1 — 한컴 normalize 와 Python NFKC 차이 → multi-method fallback

전략 (일반화 중심, hardcode 없음):
1. `form_profile.placeholders` 를 입력 받음 (Phase 5B 의 rule 결과)
2. 각 placeholder type 별 cleanup strategy:
   - blank: 그대로 유지 (양식의 빈 줄 placeholder 는 fill_empty_line 이 사용)
   - bare_marker: paragraph 단위 text 비우기 (SelectLine + Delete)
   - colon_label / bracket_heading / number_short / marker_short: find_replace multi-method

Multi-method fallback:
  Method 1: literal find_replace
  Method 2: NFKC variant (`․` → `.`)
  Method 3: regex-based (한컴 미지원 시 skip)
  Method 4: paragraph iterate + Select + Delete
"""
from __future__ import annotations

import re
import sys
import unicodedata
from typing import Any, Dict, List

from hwp_core import register
from hwp_core._helpers import validate_params


def _try_find_replace(hwp, find: str, replace: str = "", use_regex: bool = False) -> int:
    """한컴 find_replace 호출, 치환 여부 반환. 실패 시 0.

    v0.8.2: `_execute_all_replace` 는 **bool** 반환 (전/후 text 비교로 변경 여부만).
    Python 에서 `bool` 은 `int` subtype 이므로 isinstance check 순서 주의.
    bool True → 1 (변경 있음), bool False → 0 (변경 없음).
    실제 치환 횟수는 알 수 없음 (한컴 AllReplace 반환값 신뢰 불가).
    """
    try:
        # 기존 _execute_all_replace helper 활용 (from hwp_core._helpers)
        from hwp_core._helpers import _execute_all_replace
        result = _execute_all_replace(hwp, find, replace, use_regex=use_regex)
        # bool 이 int subtype 이라 isinstance(result, bool) 을 먼저 체크
        if isinstance(result, bool):
            return 1 if result else 0
        if isinstance(result, int):
            return result
        return int(result or 0)
    except Exception as e:
        print(f"[WARN] _try_find_replace '{find[:30]}': {e}", file=sys.stderr)
        return 0


def _cleanup_literal(hwp, text: str) -> int:
    """Method 1: 그대로 find_replace."""
    return _try_find_replace(hwp, text, "")


def _cleanup_nfkc(hwp, text: str) -> int:
    """Method 2: NFKC normalize 후 시도 (`․` U+2024 → `.`, `⋅` U+22C5 → `⋅` 유지)."""
    nfkc_text = unicodedata.normalize("NFKC", text)
    if nfkc_text == text:
        return 0  # 차이 없으면 skip
    return _try_find_replace(hwp, nfkc_text, "")


def _cleanup_paragraph_iterate(hwp, target_text: str, max_pages: int = 50) -> int:
    """Method 4: paragraph 단위 iterate + text 비교 + Delete.

    한컴 API 로 cursor 를 MoveDocBegin 부터 paragraph 씩 순회하면서
    현재 paragraph 의 text 가 target_text 와 일치하면 SelectLine + Delete.

    Fallback 전용 — find_replace 가 모두 실패한 경우에만 사용.
    """
    try:
        hwp.MovePos(2)  # DocBegin
        max_iter = max_pages * 100  # 안전 상한
        deleted = 0

        for _ in range(max_iter):
            try:
                # 현재 paragraph 의 텍스트 얻기
                hwp.HAction.Run("MoveLineBegin")
                hwp.HAction.Run("SelectLineEnd")
                try:
                    current = hwp.GetTextFile("TEXT", "saveblock") or ""
                except Exception:
                    current = ""
                try:
                    hwp.HAction.Run("Cancel")
                except Exception:
                    pass

                if current.strip() == target_text.strip():
                    # 이 line 을 선택 후 Delete
                    hwp.HAction.Run("MoveLineBegin")
                    hwp.HAction.Run("SelectLineEnd")
                    try:
                        hwp.HAction.Run("Delete")
                    except Exception:
                        try:
                            hwp.HAction.Run("DeleteBack")
                        except Exception:
                            pass
                    deleted += 1

                # 다음 paragraph 로 이동
                prev_pos = hwp.GetPos()
                hwp.HAction.Run("MoveDown")
                if hwp.GetPos() == prev_pos:
                    break  # 더 이상 이동 안 됨 → 문서 끝
            except Exception as e:
                print(f"[WARN] _cleanup_paragraph_iterate loop: {e}", file=sys.stderr)
                break

        return deleted
    except Exception as e:
        print(f"[WARN] _cleanup_paragraph_iterate: {e}", file=sys.stderr)
        return 0


def cleanup_one_placeholder_safe(hwp, text: str) -> Dict[str, Any]:
    """단일 placeholder 를 multi-method fallback 으로 cleanup.

    Returns:
        {method_used, replaced_count, success}
    """
    # Method 1: literal
    n = _cleanup_literal(hwp, text)
    if n > 0:
        return {"method": "literal", "replaced": n, "success": True}

    # Method 2: NFKC variant
    n = _cleanup_nfkc(hwp, text)
    if n > 0:
        return {"method": "nfkc_variant", "replaced": n, "success": True}

    # Method 4: paragraph iterate (fallback — 무거움)
    # 짧은 text (<50 char) 만 paragraph iterate 시도
    if len(text) < 50:
        n = _cleanup_paragraph_iterate(hwp, text)
        if n > 0:
            return {"method": "paragraph_iterate", "replaced": n, "success": True}

    return {"method": "none", "replaced": 0, "success": False}


@register("cleanup_all_placeholders")
def cleanup_all_placeholders(hwp, params: Dict[str, Any]) -> Dict[str, Any]:
    """Phase 5C: 양식의 placeholder 를 자동 cleanup.

    v0.8.2 CRITICAL fix:
    - 이전 default 가 bracket_heading/number_short/marker_short 포함 →
      `(1) 산업의 특성`, `1) 투자계획` 같은 양식 section heading 을 find_replace 로 삭제 →
      문서 구조 파괴 (v0.8.1 E2E 에서 발견).
    - v0.8.2: default = **`bare_marker` + `colon_label` 만** (진짜 비어있는 template).
    - section heading type (bracket_heading, numbered_section, marker_section) 은
      명시적으로 params["types"] 에 포함해야만 cleanup.

    Params:
        placeholders: list of {type, line, text} (Phase 5B 결과)
        types: list of placeholder types to cleanup
               (default: ["bare_marker", "colon_label"] — 안전)

    Returns:
        {status, cleaned, failed, method_stats, details}
    """
    validate_params(params, ["placeholders"], "cleanup_all_placeholders")
    placeholders = params["placeholders"]
    # v0.8.2: default 에서 section heading type 제거 (양식 구조 보호)
    allowed_types = set(params.get("types") or [
        "bare_marker", "colon_label"
    ])

    cleaned = 0
    failed = 0
    method_stats: Dict[str, int] = {}
    details: List[Dict[str, Any]] = []

    # 중복 text 제거 (같은 text 여러 번 시도 방지)
    seen_texts = set()
    for p in placeholders:
        ptype = p.get("type")
        if ptype not in allowed_types:
            continue
        text = (p.get("text") or "").strip()
        if not text or text in seen_texts:
            continue
        seen_texts.add(text)

        result = cleanup_one_placeholder_safe(hwp, text)
        if result["success"]:
            cleaned += result["replaced"]
            method = result["method"]
            method_stats[method] = method_stats.get(method, 0) + 1
        else:
            failed += 1

        details.append({
            "type": ptype,
            "text": text[:80],
            **result,
        })

    return {
        "status": "ok",
        "cleaned_count": cleaned,
        "failed_count": failed,
        "method_stats": method_stats,
        "details": details,
    }
