"""hwp_core.analysis.detection — 문서 타입 감지 + 패턴 분석 + 작업량 추정 + 양식 필드 감지.

Handlers:
- detect_document_type    : 5 타입 감지 (business_plan/official/form/report/general) + 프리셋 매핑
- analyze_writing_patterns: 양식 서식 패턴 학습 (XML 직접 파싱 fallback 포함)
- estimate_workload       : 복잡도 분류 + 모델 권장 + 토큰/시간 추정
- form_detect             : 빈칸/괄호/밑줄 패턴 자동 감지
"""
import os
import re
import sys

from .. import register  # 두 점!
from .._helpers import validate_params  # 두 점!


@register("detect_document_type")
def detect_document_type(hwp, params):
    """문서 타입 자동 감지 (v0.7.5.4 P3-1).

    5가지 타입: business_plan / official_document / form / report / general
    각 heuristic 점수 합산 후 최고 점수 선택.
    """
    try:
        text = hwp.get_text_file("TEXT", "") or ""
    except Exception as e:
        print(f"[WARN] detect_document_type get_text: {e}", file=sys.stderr)
        text = ""
    try:
        total_pages = hwp.PageCount
    except Exception:
        total_pages = 0

    features = {
        "has_toc": False,
        "has_heading_hierarchy": False,
        "has_form_tables": False,
        "has_guide_text": False,
        "has_official_header": False,
        "estimated_pages": total_pages,
        "text_length": len(text),
    }

    # 1. TOC 감지
    if re.search(r"목\s*차", text[:500]):
        features["has_toc"] = True
    toc_line_count = len(re.findall(r"[가-힣\w\s]+\t\d+", text))
    if toc_line_count >= 5:
        features["has_toc"] = True

    # 2. 헤딩 계층
    has_level_1 = bool(re.search(r"^\s*\d+\.\s+\S", text, re.MULTILINE))
    has_level_2 = bool(re.search(r"^\s*[가-힣]\.\s+\S", text, re.MULTILINE))
    has_level_3 = bool(re.search(r"^\s*\(\d+\)\s+\S", text, re.MULTILINE))
    if sum([has_level_1, has_level_2, has_level_3]) >= 2:
        features["has_heading_hierarchy"] = True

    # 3. 작성요령 박스
    if "작성요령" in text or "유의사항" in text or "※" in text:
        features["has_guide_text"] = True

    # 4. 공문 헤더
    if re.search(r"문서번호|시행일자|수신자|발신명의", text[:2000]):
        features["has_official_header"] = True

    # 5. 양식 표
    empty_field_count = len(re.findall(r"\([ 　]*\)|_{3,}", text))
    if empty_field_count >= 3:
        features["has_form_tables"] = True

    # 타입 판정 (점수 기반)
    scores = {
        "business_plan": 0,
        "official_document": 0,
        "form": 0,
        "report": 0,
        "general": 1,
    }
    if features["has_toc"] and features["has_heading_hierarchy"]:
        scores["business_plan"] += 3
        scores["report"] += 1
    if features["has_guide_text"] and features["has_heading_hierarchy"]:
        scores["business_plan"] += 2
    if total_pages >= 10 and features["has_heading_hierarchy"]:
        scores["business_plan"] += 2
    if features["has_official_header"]:
        scores["official_document"] += 4
    if features["has_form_tables"] and not features["has_heading_hierarchy"]:
        scores["form"] += 3
    if features["has_heading_hierarchy"] and not features["has_toc"] and total_pages < 10:
        scores["report"] += 2

    best_type = max(scores, key=scores.get)
    best_score = scores[best_type]
    total_score = sum(scores.values())
    confidence = round(best_score / total_score, 2) if total_score > 0 else 0.0

    # 프리셋 매칭 — type_override 가 있으면 감지 타입 대신 사용 (v0.7.9 fix)
    preset_type = params.get("type_override") or best_type
    try:
        from presets import get_korean_business_default
        preset = get_korean_business_default(preset_type)
    except Exception as e:
        print(f"[WARN] preset load: {e}", file=sys.stderr)
        preset = None

    return {
        "status": "ok",
        "type": best_type,
        "confidence": confidence,
        "scores": scores,
        "features": features,
        "recommended_preset": preset,
        "hint": (
            f"감지된 문서 타입: {best_type} (confidence {confidence}). "
            f"recommended_preset 을 insert_body_after_heading 의 body_style 에 전달하면 "
            f"공무원 양식 표준 서식으로 본문 삽입됩니다."
        ),
    }


@register("analyze_writing_patterns")
def analyze_writing_patterns(hwp, params):
    """양식의 서식 패턴 학습 (v0.7.1). XML 직접 파싱 fallback 포함."""
    validate_params(params, ["file_path"], "analyze_writing_patterns")
    file_path_abs = os.path.abspath(params["file_path"])

    # v0.7.2.13: .hwpx 는 XML 직접 파싱
    if file_path_abs.lower().endswith(".hwpx"):
        try:
            from hwpx_reader import read_body_style
            xml_body = read_body_style(file_path_abs)
            try:
                hwp.open(file_path_abs)
                hwp.HAction.Run("MoveDocBegin")
                page_d = hwp.get_pagedef_as_dict()
            except Exception as e:
                print(f"[WARN] page_setup COM read failed: {e}", file=sys.stderr)
                page_d = {}
            return {
                "status": "ok",
                "file_path": file_path_abs,
                "page_setup": page_d,
                "body_style": {
                    "char": xml_body.get("char", {}),
                    "para": xml_body.get("para", {}),
                },
                "title_styles": {},
                "table_styles": [],
                "numbering_pattern": "decimal_dot",
                "consistency_score": 100,
                "deviations_sample": [],
                "source": "xml",
                "para_pr_id_ref": xml_body.get("paraPrIDRef"),
                "char_pr_id_ref": xml_body.get("charPrIDRef"),
            }
        except Exception as e:
            print(f"[WARN] XML-based read failed, fallback to COM: {e}", file=sys.stderr)

    # .hwp 또는 XML 실패 → COM 경로
    from hwp_editor import get_para_shape, get_char_shape
    try:
        hwp.open(file_path_abs)
        hwp.HAction.Run("MoveDocBegin")
    except Exception as e:
        print(f"[WARN] analyze_writing_patterns open: {e}", file=sys.stderr)
    try:
        page_d = hwp.get_pagedef_as_dict()
    except Exception:
        page_d = {}
    try:
        body_para = get_para_shape(hwp)
    except Exception:
        body_para = {}
    try:
        body_char = get_char_shape(hwp)
    except Exception:
        body_char = {}

    return {
        "status": "ok",
        "file_path": params["file_path"],
        "page_setup": page_d,
        "body_style": {"char": body_char, "para": body_para},
        "title_styles": {},
        "table_styles": [],
        "numbering_pattern": "decimal_dot",
        "consistency_score": 100,
        "deviations_sample": [],
    }


@register("estimate_workload")
def estimate_workload(hwp, params):
    """작업 로드 추정 + 복잡도 분류 + 모델 권장 (v0.7.1 / v0.7.4.8 Part 3.2)."""
    validate_params(params, ["user_request"], "estimate_workload")
    user_request = params["user_request"]
    constraints = params.get("constraints", {}) or {}
    max_ref_files = int(constraints.get("max_reference_files", 5))
    max_ref_mb = int(constraints.get("max_reference_mb", 10))
    context_window = int(constraints.get("context_window_tokens", 200000))

    # 1. 양식 분석
    estimated_pages = 10
    estimated_sections = 5
    estimated_tables = 2
    analysis_data = None
    if params.get("file_path"):
        try:
            from hwp_analyzer import analyze_document as _analyze
            analysis_data = _analyze(hwp, params["file_path"])
            estimated_pages = analysis_data.get("pages", estimated_pages)
            estimated_tables = len(analysis_data.get("tables", []))
        except Exception as e:
            print(f"[WARN] estimate_workload analyze: {e}", file=sys.stderr)

    # 2. user_request 휴리스틱
    page_match = re.search(r'(\d+)\s*(페이지|쪽|장|page)', user_request, re.IGNORECASE)
    if page_match:
        estimated_pages = int(page_match.group(1))
    section_match = re.search(r'(\d+)\s*(섹션|section|chapter|챕터|단락)', user_request, re.IGNORECASE)
    if section_match:
        estimated_sections = int(section_match.group(1))

    # 3. 추정 공식
    chars_per_page = 1100
    tokens_per_char = 1.0 / 3.5
    output_chars = estimated_pages * chars_per_page
    output_tokens = int(output_chars * tokens_per_char * 1.6)

    # 입력 토큰
    input_chars = 0
    ref_summary = {"files": 0, "total_chars": 0, "tables_seen": 0, "skipped": []}
    ref_files = params.get("reference_files", []) or []
    if ref_files:
        from ref_reader import read_reference
        for i, rf in enumerate(ref_files):
            if i >= max_ref_files:
                ref_summary["skipped"].append({"file": rf, "reason": f"exceeds max_reference_files={max_ref_files}"})
                continue
            try:
                rf_size_mb = os.path.getsize(rf) / (1024 * 1024)
                if rf_size_mb > max_ref_mb:
                    ref_summary["skipped"].append({"file": rf, "reason": f"size {rf_size_mb:.1f}MB exceeds max {max_ref_mb}MB"})
                    continue
                rf_data = read_reference(rf, max_chars=20000)
                rf_chars = len(rf_data.get("content", "") or str(rf_data))
                input_chars += rf_chars
                ref_summary["files"] += 1
                ref_summary["total_chars"] += rf_chars
            except Exception as e:
                ref_summary["skipped"].append({"file": rf, "reason": f"read error: {e}"})

    if analysis_data:
        input_chars += len(analysis_data.get("full_text", "") or "")

    input_tokens = int(input_chars * tokens_per_char)
    total_tokens = input_tokens + output_tokens
    context_usage_percent = round(total_tokens / context_window * 100, 2)

    # 4. 시간 예측
    seconds_per_output_token = 0.011
    writing_seconds = int(output_tokens * seconds_per_output_token)
    analysis_seconds = 5 + estimated_tables * 2
    verification_seconds = estimated_pages * 3
    save_seconds = 30
    total_seconds = writing_seconds + analysis_seconds + verification_seconds + save_seconds

    # 5. 위험 평가
    risks = []
    if analysis_data and analysis_data.get("controls_by_type", {}).get("tbl", 0) > 5:
        risks.append({"type": "many_tables", "severity": "medium", "description": f"표 {analysis_data['controls_by_type']['tbl']}개"})
    if input_tokens > 0.4 * context_window:
        risks.append({"type": "long_context", "severity": "high", "description": f"입력 토큰 {input_tokens} > context window 40%"})
    if output_tokens > 60000:
        risks.append({"type": "output_overflow", "severity": "high", "description": f"출력 토큰 {output_tokens} > 60k"})
    if total_tokens > 0.8 * context_window:
        risks.append({"type": "context_window_overflow", "severity": "critical", "description": "전체 토큰이 context window 80% 초과"})

    # 6. recommended_action
    high_risks = sum(1 for r in risks if r["severity"] in ("high", "critical"))
    if high_risks >= 2:
        recommended = "reduce_scope"
    elif total_tokens > 0.5 * context_window or estimated_pages > 20:
        recommended = "split_into_sessions"
    else:
        recommended = "proceed"

    # 7. split suggestion
    split_suggestion = []
    if recommended == "split_into_sessions" and estimated_sections > 0:
        half = max(1, estimated_sections // 2)
        split_suggestion = [
            {"section_range": f"1-{half}", "estimated_pages": estimated_pages // 2, "estimated_tokens": total_tokens // 2},
            {"section_range": f"{half + 1}-{estimated_sections}", "estimated_pages": estimated_pages // 2, "estimated_tokens": total_tokens // 2},
        ]

    # 복잡도 분류
    num_refs = len(ref_summary) if isinstance(ref_summary, list) else 0
    if estimated_sections <= 3 and estimated_tables <= 2 and num_refs <= 1:
        complexity_class = "simple_fill"
    elif estimated_sections > 10 or estimated_tables > 5:
        complexity_class = "structured_analysis"
    else:
        complexity_class = "text_generation"

    # 모델 권장
    if complexity_class == "simple_fill":
        recommended_model = {
            "speed": "claude-haiku-4-5",
            "balanced": "claude-sonnet-4-6",
            "quality": "claude-sonnet-4-6",
            "default": "claude-haiku-4-5",
            "reason": "단순 양식/표준 문서 — Haiku 로 빠른 처리 권장",
        }
    elif complexity_class == "text_generation":
        recommended_model = {
            "speed": "claude-sonnet-4-6",
            "balanced": "claude-sonnet-4-6",
            "quality": "claude-opus-4-6",
            "default": "claude-sonnet-4-6",
            "reason": "일반 업무 문서 — Sonnet 이 균형잡힌 선택",
        }
    else:
        recommended_model = {
            "speed": "claude-sonnet-4-6",
            "balanced": "claude-opus-4-6",
            "quality": "claude-opus-4-6",
            "default": "claude-opus-4-6",
            "reason": "복잡한 구조 분석 — Opus 로 정확도 보장 권장",
        }

    duration_by_model = {
        "haiku": int(round(total_seconds * 0.5)),
        "sonnet": int(round(total_seconds * 1.0)),
        "opus": int(round(total_seconds * 1.8)),
    }

    input_ratio = input_tokens / context_window if context_window > 0 else 0
    if input_ratio < 0.15:
        ce_recommendation = "여유 충분"
    elif input_ratio < 0.35:
        ce_recommendation = "적정"
    elif input_ratio < 0.6:
        ce_recommendation = "주의 (Sonnet/Opus 권장)"
    else:
        ce_recommendation = "경고 (참고자료 축소 또는 split 권장)"

    return {
        "status": "ok",
        "estimated_pages": estimated_pages,
        "estimated_sections": estimated_sections,
        "estimated_tables": estimated_tables,
        "tokens": {
            "input_tokens": input_tokens,
            "output_tokens_estimate": output_tokens,
            "total_tokens_estimate": total_tokens,
            "context_window_usage_percent": context_usage_percent,
        },
        "duration_seconds_estimate": total_seconds,
        "duration_breakdown": {
            "analysis": analysis_seconds,
            "writing": writing_seconds,
            "verification": verification_seconds,
            "save": save_seconds,
        },
        "risks": risks,
        "recommended_action": recommended,
        "split_suggestion": split_suggestion,
        "reference_summary": ref_summary,
        "constraints_applied": {
            "max_reference_files": max_ref_files,
            "max_reference_mb": max_ref_mb,
            "context_window_tokens": context_window,
        },
        "suggested_workflow": complexity_class,
        "recommended_model_for_complexity": recommended_model,
        "duration_by_model": duration_by_model,
        "context_efficiency": {
            "input_tokens_to_context_ratio": round(input_ratio, 3),
            "recommendation": ce_recommendation,
        },
    }


@register("form_detect")
def form_detect(hwp, params):
    """문서 텍스트에서 빈칸/괄호/밑줄 패턴으로 양식 필드 자동 감지.

    v0.6.6 B3: extract_all_text 사용 (ReleaseScan finally 보장).
    """
    from hwp_editor import extract_all_text
    text = ""
    try:
        text = extract_all_text(hwp, max_iters=10000, strip_each=False, separator="\n")
    except Exception as e:
        print(f"[WARN] form_detect extract_all_text: {e}", file=sys.stderr)

    # 패턴 감지: ( ), [ ], ___, ☐, □, ○, ◯, 빈칸+콜론
    patterns = [
        (r'\(\s*\)', 'bracket_empty', '빈 괄호'),
        (r'\[\s*\]', 'square_empty', '빈 대괄호'),
        (r'_{3,}', 'underline', '밑줄 빈칸'),
        (r'[☐□]', 'checkbox', '체크박스'),
        (r'[○◯]', 'circle', '빈 원'),
        (r':\s*$', 'colon_empty', '콜론 뒤 빈칸'),
    ]
    fields = []
    for pattern, field_type, description in patterns:
        for m in re.finditer(pattern, text, re.MULTILINE):
            context = text[max(0, m.start() - 20):m.end() + 20].strip()
            fields.append({
                "type": field_type,
                "description": description,
                "position": m.start(),
                "context": context[:50],
            })
    return {"status": "ok", "total_fields": len(fields), "fields": fields[:50]}
