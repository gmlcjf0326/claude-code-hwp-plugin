"""Form fill auto — rule 기반 양식 독립 통합 pipeline (Phase 5G, v0.8.0).

사용자 핵심 원칙:
> "수천개의 양식을 모두 학습할 수 없기에 그때그때 양식에 따라서 달라짐을 알아야합니다"

`korean_business_fill` 의 일반화 버전. 임의 HWP 양식에 대해:

1. **analyze_form** — form_profile 자동 생성 (sections/tables/markers/placeholders)
2. **mark_review_required** — 실제 데이터 필요 영역 식별 (type 기반)
3. **fill_tables** (선택) — smart_fill_table_auto 로 type 기반 cell 매핑 + auto-style
4. **insert_bodies** (선택) — insert_body_after_heading 로 본문 channel
5. **cleanup_all_placeholders** — Phase 5B 감지 + Phase 5C multi-method cleanup
6. **save_as** — 원본 보존 + 새 경로 저장

사용자 입력 최소:
- 양식 파일 경로
- 본문 fills (heading → body_text, optional)
- 표 fills (table_index → cells, optional)
- output_path
"""
from __future__ import annotations

import os
import sys
from typing import Any, Dict, List, Optional

from hwp_core import register
from hwp_core._helpers import validate_params, validate_file_path


@register("form_fill_auto")
def form_fill_auto(hwp, params: Dict[str, Any]) -> Dict[str, Any]:
    """v0.8.0 Phase 5G: rule 기반 양식 독립 통합 pipeline.

    Params:
        form_file: 양식 파일 경로 (필수)
        output_path: 저장 경로 (필수)
        body_fills: [{heading, body_text}, ...] (optional)
        table_fills: [{table_index, cells, table_type?}, ...] (optional)
        cleanup_placeholders: bool (default True)
        mark_review: bool (default True)
        return_profile: bool (default False)

    Returns:
        전체 파이프라인 결과 — {profile_summary, review_required, body_result, table_result, cleanup_result, save_result}
    """
    validate_params(params, ["form_file", "output_path"], "form_fill_auto")
    form_file = validate_file_path(params["form_file"], must_exist=True)
    output_path = validate_file_path(params["output_path"], must_exist=False)
    body_fills = params.get("body_fills", []) or []
    table_fills = params.get("table_fills", []) or []
    cleanup_placeholders = bool(params.get("cleanup_placeholders", True))
    mark_review = bool(params.get("mark_review", True))
    return_profile = bool(params.get("return_profile", False))

    result: Dict[str, Any] = {"status": "ok", "steps": {}}

    # === Step 1: analyze_form (양식 독립 분석) ===
    try:
        from hwp_core.analysis.form_handler import analyze_form
        af_result = analyze_form(hwp, {
            "file_path": form_file,
            "include_placeholders": True,
            "include_guidance": True,
            "summary_only": False,
        })
        result["steps"]["analyze_form"] = {"status": "ok", "summary": af_result.get("summary", {})}
        profile = af_result.get("profile", {})
    except Exception as e:
        print(f"[ERROR] form_fill_auto analyze_form: {e}", file=sys.stderr)
        result["status"] = "error"
        result["steps"]["analyze_form"] = {"status": "error", "error": str(e)}
        return result

    # === Step 2: mark_review_required (type 기반 식별) ===
    if mark_review:
        try:
            from hwp_core.analysis.form_handler import mark_review_required
            mr_result = mark_review_required(hwp, {"profile": profile})
            result["steps"]["mark_review_required"] = {
                "status": "ok",
                "review_required_count": mr_result.get("review_required_count", 0),
                "review_required_tables": mr_result.get("review_required_tables", [])[:10],
            }
        except Exception as e:
            print(f"[WARN] form_fill_auto mark_review: {e}", file=sys.stderr)
            result["steps"]["mark_review_required"] = {"status": "error", "error": str(e)}

    # === Step 3: fill_tables (smart_fill_table_auto + type) ===
    table_result: Dict[str, Any] = {"total": 0, "filled": 0, "failed": 0, "details": []}
    if table_fills:
        try:
            from hwp_editor import smart_fill_table_auto
            profile_tables = profile.get("tables", [])
            table_types_by_idx = {t.get("index"): t.get("type") for t in profile_tables}

            for tf in table_fills:
                table_index = tf.get("table_index")
                cells = tf.get("cells", [])
                ttype = tf.get("table_type") or table_types_by_idx.get(table_index)
                r = smart_fill_table_auto(hwp, table_index, cells, table_type=ttype)
                table_result["total"] += 1
                table_result["filled"] += r.get("filled", 0)
                table_result["failed"] += r.get("failed", 0)
                table_result["details"].append({
                    "table_index": table_index,
                    "table_type": ttype,
                    **r,
                })
        except Exception as e:
            print(f"[WARN] form_fill_auto fill_tables: {e}", file=sys.stderr)
            table_result["error"] = str(e)
    result["steps"]["fill_tables"] = table_result

    # === Step 4: insert_bodies ===
    body_result: Dict[str, Any] = {"total": 0, "ok": 0, "failed": 0, "details": []}
    if body_fills:
        try:
            from hwp_core.text_editing.insertions import insert_body_after_heading
            for bf in body_fills:
                heading = bf.get("heading", "")
                body_text = bf.get("body_text", "")
                if not heading or not body_text:
                    continue
                r = insert_body_after_heading(hwp, {"heading": heading, "body_text": body_text})
                body_result["total"] += 1
                st = r.get("status") if isinstance(r, dict) else "?"
                if st == "ok":
                    body_result["ok"] += 1
                else:
                    body_result["failed"] += 1
                body_result["details"].append({
                    "heading": heading[:40],
                    "status": st,
                    "match_tier": r.get("match_tier") if isinstance(r, dict) else None,
                })
        except Exception as e:
            print(f"[WARN] form_fill_auto insert_bodies: {e}", file=sys.stderr)
            body_result["error"] = str(e)
    result["steps"]["insert_bodies"] = body_result

    # === Step 5: cleanup_all_placeholders (Phase 5C 일반화) ===
    if cleanup_placeholders:
        try:
            from hwp_core.text_editing.placeholder_cleanup import cleanup_all_placeholders
            # profile 의 placeholders 를 refresh (본문 insert 후 양식 변화)
            # 단순히 Phase 5A 의 결과 사용
            placeholders = profile.get("placeholders", [])
            # 본문 insert 후 일부 placeholder 가 사라졌을 수 있으므로 re-detect 고려
            cr = cleanup_all_placeholders(hwp, {"placeholders": placeholders})
            result["steps"]["cleanup_placeholders"] = {
                "status": "ok",
                "cleaned_count": cr.get("cleaned_count", 0),
                "failed_count": cr.get("failed_count", 0),
                "method_stats": cr.get("method_stats", {}),
            }
        except Exception as e:
            print(f"[WARN] form_fill_auto cleanup: {e}", file=sys.stderr)
            result["steps"]["cleanup_placeholders"] = {"status": "error", "error": str(e)}

    # === Step 6: save_as ===
    try:
        from hwp_core.document import save_as
        sr = save_as(hwp, {"path": output_path})
        result["steps"]["save_as"] = sr
    except Exception as e:
        print(f"[ERROR] form_fill_auto save_as: {e}", file=sys.stderr)
        result["status"] = "error"
        result["steps"]["save_as"] = {"status": "error", "error": str(e)}
        return result

    # === Summary ===
    orig_exists = os.path.exists(form_file)
    orig_size = os.path.getsize(form_file) if orig_exists else 0
    out_exists = os.path.exists(output_path)
    out_size = os.path.getsize(output_path) if out_exists else 0

    result["summary"] = {
        "form_file": os.path.basename(form_file),
        "output_path": os.path.basename(output_path),
        "original_size": orig_size,
        "original_preserved": True,  # save_as 는 원본 변경 X
        "output_size": out_size,
        "body_ok": body_result.get("ok", 0),
        "body_total": body_result.get("total", 0),
        "table_filled": table_result.get("filled", 0),
        "placeholder_cleaned": result["steps"].get("cleanup_placeholders", {}).get("cleaned_count", 0),
    }

    if return_profile:
        result["profile"] = profile

    return result
