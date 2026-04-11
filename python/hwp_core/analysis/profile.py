"""hwp_core.analysis.profile — 서식 프로파일 추출 + 템플릿 구조 분석 + 스냅샷.

Handlers:
- extract_style_profile       : 본문 + 첫 표 셀 서식 프로파일
- extract_full_profile        : 용지 + 본문 + 최대 5개 표 치수
- extract_template_structure  : heading 계층 구조 추출 (최대 4 depth)
- snapshot_template_style     : 양식 원본 서식 전체 스냅샷 (JSON 저장)
"""
import json as _json
import os
import re
import sys
import time as _time

from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely  # 두 점!


@register("extract_style_profile")
def extract_style_profile(hwp, params):
    """양식 문서에서 본문 + 표 셀 서식 프로파일 추출."""
    from hwp_editor import get_char_shape, get_para_shape
    profiles = {}
    # 본문 서식 (문서 시작 위치)
    hwp.MovePos(2)
    profiles["body"] = {"char": get_char_shape(hwp), "para": get_para_shape(hwp)}
    # 표 셀 서식 (첫 번째 표)
    try:
        hwp.get_into_nth_table(0)
        profiles["table_cell"] = {"char": get_char_shape(hwp), "para": get_para_shape(hwp)}
        _exit_table_safely(hwp)
    except Exception:
        profiles["table_cell"] = None
    return {"status": "ok", "profiles": profiles}


@register("extract_full_profile")
def extract_full_profile(hwp, params):
    """양식 종합 프로파일 — 용지 + 본문 + 표 치수 (최대 5개).

    내부적으로 get_page_setup 을 직접 호출 (dispatch 경유 X).
    get_table_dimensions 는 REGISTRY lookup 으로 dynamic.
    """
    from hwp_editor import get_char_shape, get_para_shape
    from .metadata import get_page_setup as _get_page_setup  # same sub-package
    profile = {"status": "ok"}

    # 1. 용지 설정 (직접 호출)
    try:
        profile["page_setup"] = _get_page_setup(hwp, {})
    except Exception as e:
        profile["page_setup"] = {"error": str(e)}

    # 2. 본문 서식 (커서가 본문에 있을 때)
    try:
        hwp.MovePos(2)
    except Exception:
        pass
    try:
        profile["body_char"] = get_char_shape(hwp)
    except Exception as e:
        profile["body_char"] = {"error": str(e)}
    try:
        profile["body_para"] = get_para_shape(hwp)
    except Exception as e:
        profile["body_para"] = {"error": str(e)}

    # 3. 표 치수 (최대 5개 표) — REGISTRY lookup
    profile["tables"] = []
    try:
        from hwp_core import REGISTRY
        dims_handler = REGISTRY.get("get_table_dimensions")
        if dims_handler:
            for i in range(5):
                try:
                    dims = dims_handler(hwp, {"table_index": i})
                    if dims.get("status") == "ok":
                        profile["tables"].append(dims)
                except Exception:
                    break
    except Exception as e:
        profile["tables_error"] = str(e)

    return profile


@register("extract_template_structure")
def extract_template_structure(hwp, params):
    """양식의 heading 계층 구조 추출 (최대 4 depth).

    패턴: 제N장/조/절, I./II., 1./1.1/1.1.1, 가./나., (1)/(가)
    반환: sections[] — {id, title, level, para_index}
    """
    validate_params(params, ["file_path"], "extract_template_structure")
    from hwp_analyzer import analyze_document as _analyze
    max_depth = int(params.get("max_depth", 4))

    # 1. 기존 analyze_document 재활용
    analysis = _analyze(hwp, params["file_path"])

    # 2. heading 인식 정규식 (full_text 단락 단위 분석)
    _heading_patterns = [
        (re.compile(r'^제\s*(\d+)\s*[장조절]\s'), 1),
        (re.compile(r'^([IVX]+)\.\s'), 1),
        (re.compile(r'^(\d+)\.\s'), 1),
        (re.compile(r'^(\d+)\.(\d+)\s'), 2),
        (re.compile(r'^(\d+)\.(\d+)\.(\d+)\s'), 3),
        (re.compile(r'^([가-힣])\.\s'), 2),
        (re.compile(r'^\(([가-힣\d])\)\s'), 3),
    ]
    full_text = analysis.get("full_text", "") or ""
    paragraphs = full_text.split("\n")
    sections = []
    section_id = 0
    for idx, para in enumerate(paragraphs):
        stripped = para.strip()
        if not stripped:
            continue
        for pat, level in _heading_patterns:
            m = pat.match(stripped)
            if m and level <= max_depth:
                section_id += 1
                sections.append({
                    "id": f"sec_{section_id}",
                    "title": stripped[:80],
                    "level": level,
                    "para_index": idx,
                })
                break

    return {
        "status": "ok",
        "file_path": analysis.get("file_path"),
        "total_pages": analysis.get("pages", 0),
        "sections": sections,
        "section_count": len(sections),
        "global_tables_count": len(analysis.get("tables", [])),
        "global_fields_count": len(analysis.get("fields", [])),
        "controls_by_type": analysis.get("controls_by_type", {}),
    }


@register("snapshot_template_style")
def snapshot_template_style(hwp, params):
    """양식 원본 서식 프로파일 스냅샷 (v0.7.5.4 P2-3)."""
    from hwp_editor import get_char_shape, get_para_shape

    snapshot = {
        "version": "v0.7.5.4",
        "body_default": None,
        "heading_samples": [],
        "table_cell_samples": [],
        "warnings": [],
    }

    # 1) 본문 기본 서식
    try:
        hwp.MovePos(2)  # MoveDocBegin
        snapshot["body_default"] = {
            "char": get_char_shape(hwp),
            "para": get_para_shape(hwp),
        }
    except Exception as e:
        snapshot["warnings"].append(f"body_default: {e}")

    # 2) heading 샘플
    heading_patterns = [
        (1, r"^\s*\d+\.\s+\S"),
        (2, r"^\s*[가-힣]\.\s+\S"),
        (3, r"^\s*\(\d+\)\s+\S"),
    ]
    try:
        full_text = hwp.get_text_file("TEXT", "") or ""
    except Exception:
        full_text = ""

    for level, pattern in heading_patterns:
        try:
            match = re.search(pattern, full_text, re.MULTILINE)
            if not match:
                continue
            heading_text = match.group(0).strip()
            search_key = heading_text[:20]
            hwp.MovePos(2)
            act = hwp.HAction
            pset = hwp.HParameterSet.HFindReplace
            act.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = search_key
            pset.Direction = 0
            pset.IgnoreMessage = 1
            if act.Execute("RepeatFind", pset.HSet):
                snapshot["heading_samples"].append({
                    "level": level,
                    "sample_text": heading_text[:50],
                    "char": get_char_shape(hwp),
                    "para": get_para_shape(hwp),
                })
        except Exception as e:
            snapshot["warnings"].append(f"heading_level_{level}: {e}")

    # 3) 표 셀 샘플
    try:
        hwp.MovePos(2)
        hwp.get_into_nth_table(0)
        snapshot["table_cell_samples"].append({
            "table_index": 0,
            "cell_tab": 0,
            "char": get_char_shape(hwp),
            "para": get_para_shape(hwp),
        })
        _exit_table_safely(hwp)
    except Exception as e:
        snapshot["warnings"].append(f"table_cell_sample: {e}")

    # 4) 저장
    snapshot_id = f"snap_{int(_time.time())}"
    snapshot["snapshot_id"] = snapshot_id
    try:
        state_dir = os.path.expanduser("~/.hwp_studio_state")
        os.makedirs(state_dir, exist_ok=True)
        snapshot_path = os.path.join(state_dir, f"{snapshot_id}.json")
        with open(snapshot_path, "w", encoding="utf-8") as fp:
            _json.dump(snapshot, fp, ensure_ascii=False, indent=2)
        snapshot["saved_path"] = snapshot_path
    except Exception as e:
        snapshot["warnings"].append(f"save: {e}")

    return {"status": "ok", **snapshot}
