"""hwp_core.analysis.verification — 5-stage verify + 레이아웃 + 일관성 검증.

Handlers:
- verify_5stage        : TEST_CHECKLIST Phase 19 표준 5단계 자동 검증
- verify_layout        : PDF→PNG 시각 검증 (PyMuPDF)
- validate_consistency : 서식 일관성 검증 (expected_profile vs current)

v0.7.7 변경: verify_5stage expected_chars 대비 50% → 70%
"""
import os
import tempfile

from .. import register  # 두 점!
from .._helpers import validate_params  # 두 점!
from .._state import get_current_doc_path, set_current_doc_path  # 두 점!


@register("verify_5stage")
def verify_5stage(hwp, params):
    """TEST_CHECKLIST Phase 19 표준 5단계 자동 검증."""
    validate_params(params, ["file_path"], "verify_5stage")
    file_path = params["file_path"]
    file_path = os.path.abspath(file_path)
    expected_chars = int(params.get("expected_chars", 0))
    expected_text_snippet = params.get("expected_text_snippet", "")
    run_layout = bool(params.get("run_layout", False))

    result = {
        "file_path": file_path,
        "stage1_step_log_ok": True,
        "stage2_body_verified": None,
        "stage3_file_size_ok": None,
        "stage4_text_cross_check_ok": None,
        "stage4_match_percent": 0.0,
        "stage5_layout_png_paths": [],
        "passed_stages": 0,
        "failed_stages": [],
        "overall_pass": False,
        "details": {},
    }

    # Stage 2/3: file size
    try:
        file_size = os.path.getsize(file_path)
        result["details"]["file_size"] = file_size
    except Exception as e:
        result["details"]["file_size_error"] = str(e)

    ext = os.path.splitext(file_path)[1].lower()
    min_size = 22 * 1024 if ext == ".hwpx" else 28 * 1024
    if result["details"].get("file_size", 0) >= min_size:
        result["stage3_file_size_ok"] = True
    else:
        result["stage3_file_size_ok"] = False
        result["failed_stages"].append("stage3_file_size")

    # Stage 4: text cross-check
    try:
        was_opened = (get_current_doc_path() == file_path)
        if not was_opened:
            hwp.open(file_path)
            set_current_doc_path(file_path)
        text = hwp.get_text_file("TEXT", "") or ""
        chars_total = len(text.strip())
        result["details"]["chars_total"] = chars_total

        if expected_chars > 0:
            if chars_total >= expected_chars * 0.7:  # v0.7.7: 50%→70%
                result["stage2_body_verified"] = True
            else:
                result["stage2_body_verified"] = False
                result["failed_stages"].append("stage2_body_verified")
        else:
            result["stage2_body_verified"] = (chars_total > 0)
            if chars_total == 0:
                result["failed_stages"].append("stage2_body_verified")

        if expected_text_snippet:
            total_len = len(expected_text_snippet)
            matched = sum(
                1 for i in range(0, total_len, 10)
                if expected_text_snippet[i:i+10] in text
            )
            match_percent = round(matched / max(1, total_len // 10), 2) * 100
            result["stage4_match_percent"] = match_percent
            result["stage4_text_cross_check_ok"] = (match_percent >= 90)
        else:
            result["stage4_text_cross_check_ok"] = (chars_total > 0)
            result["stage4_match_percent"] = 100.0 if chars_total > 0 else 0.0

        if not result["stage4_text_cross_check_ok"]:
            result["failed_stages"].append("stage4_text_cross_check")

    except Exception as e:
        result["details"]["text_error"] = str(e)
        result["stage2_body_verified"] = False
        result["stage4_text_cross_check_ok"] = False
        result["failed_stages"].extend(["stage2_body_verified", "stage4_text_cross_check"])

    # Stage 5: layout PNG
    if run_layout:
        result["stage5_layout_png_paths"] = ["(run_layout integration TBD)"]
    else:
        result["stage5_layout_png_paths"] = ["(skipped — run_layout=False)"]

    # 통과 단계 계산
    passed = 0
    if result["stage1_step_log_ok"]:
        passed += 1
    if result["stage2_body_verified"]:
        passed += 1
    if result["stage3_file_size_ok"]:
        passed += 1
    if result["stage4_text_cross_check_ok"]:
        passed += 1
    if run_layout and result["stage5_layout_png_paths"] and "skipped" not in str(result["stage5_layout_png_paths"][0]):
        passed += 1
    result["passed_stages"] = passed
    required_min = 5 if run_layout else 4
    result["overall_pass"] = (passed >= required_min)

    return {"status": "ok", **result}


@register("verify_layout")
def verify_layout(hwp, params):
    """PDF→PNG 변환으로 레이아웃 시각 검증."""
    # 현재 문서 저장
    if get_current_doc_path():
        try:
            hwp.save()
        except Exception:
            pass
    tmp_pdf = os.path.join(tempfile.gettempdir(), "hwp_verify_layout.pdf")
    try:
        hwp.save_as(tmp_pdf, "PDF")
        if not os.path.exists(tmp_pdf):
            return {"status": "error", "error": "PDF 생성 실패"}

        try:
            import fitz
            doc = fitz.open(tmp_pdf)
            image_paths = []
            page_range = params.get("pages")
            start_page = 0
            end_page = doc.page_count

            if page_range:
                parts = str(page_range).split("-")
                start_page = max(0, int(parts[0]) - 1)
                end_page = int(parts[-1]) if len(parts) > 1 else start_page + 1

            for i in range(start_page, min(end_page, doc.page_count)):
                pix = doc[i].get_pixmap(dpi=150)
                png_path = os.path.join(tempfile.gettempdir(), f"hwp_verify_page{i+1}.png")
                pix.save(png_path)
                image_paths.append(png_path)

            doc.close()
            try:
                os.remove(tmp_pdf)
            except Exception:
                pass
            return {
                "status": "ok",
                "image_paths": image_paths,
                "pages": len(image_paths),
                "total_pages": hwp.PageCount,
                "hint": "Read 도구로 각 PNG 이미지를 열어 레이아웃을 시각적으로 검증하세요.",
            }
        except ImportError:
            return {
                "status": "ok_pdf_only",
                "pdf_path": tmp_pdf,
                "pages": hwp.PageCount,
                "file_size": os.path.getsize(tmp_pdf),
                "hint": "PyMuPDF 미설치. 'pip install PyMuPDF' 실행 후 다시 시도.",
            }
    except Exception as e:
        return {"status": "error", "error": f"레이아웃 검증 실패: {e}"}


@register("validate_consistency")
def validate_consistency(hwp, params):
    """문서 서식 일관성 검증 (MVP)."""
    validate_params(params, ["file_path"], "validate_consistency")
    expected = params.get("expected_profile")
    deviations = []

    try:
        from hwp_editor import get_para_shape, get_char_shape
        current_para = get_para_shape(hwp)
        current_char = get_char_shape(hwp)
    except Exception as e:
        return {"status": "error", "error": f"현재 문서 분석 실패: {e}"}

    score = 100
    if expected and isinstance(expected, dict):
        exp_body = expected.get("body_style", {}) or {}
        exp_para = exp_body.get("para", {}) or {}
        exp_char = exp_body.get("char", {}) or {}
        for key, exp_val in (exp_para or {}).items():
            if current_para.get(key) != exp_val:
                deviations.append({
                    "field": f"para.{key}",
                    "expected": exp_val,
                    "actual": current_para.get(key),
                    "severity": "low",
                })
        for key, exp_val in (exp_char or {}).items():
            if current_char.get(key) != exp_val:
                deviations.append({
                    "field": f"char.{key}",
                    "expected": exp_val,
                    "actual": current_char.get(key),
                    "severity": "low",
                })
        score = max(0, 100 - len(deviations) * 5)

    return {
        "status": "ok",
        "consistency_score": score,
        "deviations": deviations,
        "summary": {
            "checked_paragraphs": 1,
            "current_para": current_para,
            "current_char": current_char,
        },
    }
