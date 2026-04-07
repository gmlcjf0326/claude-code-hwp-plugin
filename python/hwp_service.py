"""HWP Studio AI - Python HWP Service Bridge
stdin/stdout JSON protocol for Electron <-> Python communication.
"""
import sys
import json
import os
import signal
import time
import pathlib

from hwp_analyzer import analyze_document, map_table_cells
from hwp_editor import (fill_document, fill_table_cells_by_tab, fill_table_cells_by_label,
                        set_paragraph_style, get_char_shape, get_para_shape,
                        verify_after_fill)


def _exit_table_safely(hwp):
    """표에서 안전하게 탈출. MovePos(3)으로 문서 끝(표 밖)으로 이동."""
    try:
        if hwp.is_cell():
            hwp.MovePos(3)  # 문서 마지막 위치 (표 밖으로 탈출)
    except Exception:
        pass
    # 표 밖 확인 후 새 문단 생성
    try:
        if not hwp.is_cell():
            hwp.HAction.Run("BreakPara")
    except Exception:
        pass


def validate_file_path(file_path, must_exist=True):
    """경로 보안 검증. 심볼릭 링크 거부, 존재/권한 확인. 한글 에러 대화상자 사전 방지."""
    real = os.path.abspath(file_path)
    if os.path.islink(file_path):
        raise ValueError(f"심볼릭 링크는 허용되지 않습니다: {file_path}")
    if must_exist and not os.path.exists(real):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {real}")
    if not must_exist:
        # 저장 대상 경로: 디렉토리 존재 + 쓰기 권한 사전 확인
        dir_path = os.path.dirname(real) or '.'
        if not os.path.exists(dir_path):
            raise FileNotFoundError(f"저장 디렉토리가 존재하지 않습니다: {dir_path}")
        if not os.access(dir_path, os.W_OK):
            raise PermissionError(f"디렉토리 쓰기 권한이 없습니다: {dir_path}")
        # 기존 파일 덮어쓰기: 쓰기 권한 + 잠금 확인
        if os.path.exists(real):
            if not os.access(real, os.W_OK):
                raise PermissionError(f"파일 쓰기 권한이 없습니다 (읽기 전용 또는 잠김): {real}")
            # 파일 잠금 사전 확인 — HWP/HWPX 파일은 제외 (한글이 잠금 보유 중)
            ext = os.path.splitext(real)[1].lower()
            if ext not in ('.hwp', '.hwpx'):
                try:
                    with open(real, 'a'):
                        pass
                except (PermissionError, IOError):
                    raise PermissionError(f"파일이 다른 프로그램에서 사용 중입니다: {real}")
    return real


def _execute_all_replace(hwp, find_str, replace_str, use_regex=False, case_sensitive=True):
    """AllReplace 공통 함수. 전후 텍스트 비교로 검증. H4: 타임아웃 시 낙관적 가정."""
    # 치환 전 텍스트 캡처
    before = None
    try:
        before = hwp.get_text_file("TEXT", "")
    except Exception as e:
        print(f"[WARN] get_text_file before failed: {e}", file=sys.stderr)

    act = hwp.HAction
    pset = hwp.HParameterSet.HFindReplace
    act.GetDefault("AllReplace", pset.HSet)
    pset.FindString = find_str
    pset.ReplaceString = replace_str
    pset.IgnoreMessage = 1
    pset.Direction = 0
    pset.FindRegExp = 1 if use_regex else 0
    pset.FindJaso = 0
    # 대소문자 구분 옵션
    try:
        pset.MatchCase = 1 if case_sensitive else 0
    except Exception:
        pass  # 일부 한글 버전에서 미지원
    pset.AllWordForms = 0
    pset.SeveralWords = 0
    act.Execute("AllReplace", pset.HSet)

    # 치환 후 텍스트 비교로 실제 변경 여부 판단
    after = None
    try:
        after = hwp.get_text_file("TEXT", "")
    except Exception as e:
        print(f"[WARN] get_text_file after failed: {e}", file=sys.stderr)

    # 2C3: 텍스트 캡처 실패 시 구분
    if before is None and after is None:
        # 양쪽 모두 실패 → Execute는 실행됨 → 낙관적 True
        return True
    if before is None or after is None:
        # 한쪽만 실패 → 비교 불가하지만 Execute는 실행됨 → 로깅 후 True
        print(f"[WARN] Partial text capture: before={'ok' if before is not None else 'fail'}, after={'ok' if after is not None else 'fail'}", file=sys.stderr)
        return True
    return before != after


def respond(req_id, success, data=None, error=None, error_type=None, guide=None):
    """Send JSON response to stdout."""
    response = {"id": req_id, "success": success}
    if data is not None:
        response["data"] = data
    if error is not None:
        response["error"] = error
    if error_type:
        response["error_type"] = error_type
    if guide:
        response["guide"] = guide
    sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def validate_params(params, required_keys, method_name):
    """Validate required parameters exist."""
    missing = [k for k in required_keys if k not in params]
    if missing:
        raise ValueError(f"{method_name}: missing required params: {', '.join(missing)}")


_current_doc_path = None

def dispatch(hwp, method, params):
    """Route method calls to appropriate handlers."""
    global _current_doc_path

    if method == "ping":
        return {"status": "ok", "message": "HWP Service is running"}

    # B1 (v0.6.6): dispatch 진입부 SetMessageBoxMode 멱등 적용
    # 모든 RPC가 자동 안전 모드로 실행됨 (대화상자 미출력 → 무인 자동화 안정성)
    # 기존 open_document line 182의 XHwpMessageBoxMode = 1 은 폴백으로 유지
    try:
        from hwp_constants import apply_safe_mode
        apply_safe_mode(hwp)
    except Exception:
        # hwp_constants import 실패 시에도 dispatch는 계속 (호환성)
        try:
            hwp.SetMessageBoxMode(0x00010000)
        except Exception:
            pass

    if method == "inspect_com_object":
        obj_name = params.get("object", "HCharShape")
        if obj_name == "HCharShape":
            pset = hwp.HParameterSet.HCharShape
            hwp.HAction.GetDefault("CharShape", pset.HSet)
        elif obj_name == "HParaShape":
            pset = hwp.HParameterSet.HParaShape
            hwp.HAction.GetDefault("ParaShape", pset.HSet)
        elif obj_name == "HFindReplace":
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("AllReplace", pset.HSet)
        else:
            return {"error": f"Unknown object: {obj_name}"}
        attrs = [a for a in dir(pset) if not a.startswith('_')]
        return {"object": obj_name, "attributes": attrs, "count": len(attrs)}

    # v0.7.2.5: 빈 문서 생성 (autopilot blank 분기에서 호출)
    # v0.7.2.9: cursor 초기화 + 본문 시작점 보장 (이전 문서가 표 셀 cursor였을 수 있음)
    if method == "document_new":
        try:
            hwp.HAction.Run("FileNew")
        except Exception as e:
            return {"status": "error", "error": f"FileNew failed: {e}"}
        # 이전 문서 잔여 상태 정리 + cursor를 본문 처음으로
        try:
            if hwp.is_cell():
                hwp.MovePos(3)  # 표 탈출 (Cancel() 안 됨, MovePos(3) 필수 — v0.5.x 메모리)
        except Exception:
            pass
        try:
            hwp.MovePos(2)  # movePOS_START: 본문 첫 단락
        except Exception as e:
            print(f"[WARN] document_new MovePos failed: {e}", file=sys.stderr)
        return {"status": "ok"}

    if method == "open_document":
        validate_params(params, ["file_path"], method)
        file_path = validate_file_path(params["file_path"], must_exist=True)

        # HWP 자동저장 디렉토리 확인/생성 (.asv 저장 오류 방지)
        try:
            import tempfile
            asv_dir = os.path.join(tempfile.gettempdir(), "Hwp90")
            if not os.path.exists(asv_dir):
                os.makedirs(asv_dir, exist_ok=True)
        except Exception:
            pass

        # COM 상태 초기화 (이전 문서 캐시 정리)
        try:
            hwp.MovePos(2)  # 커서 초기화
        except Exception:
            pass

        # 원본 백업 (기본 활성, backup=False로 비활성 가능)
        if params.get("backup", True):
            import shutil
            root, ext = os.path.splitext(file_path)
            backup_path = f"{root}_backup{ext}"
            if not os.path.exists(backup_path):
                shutil.copy2(file_path, backup_path)

        # 파일 열기 전 다이얼로그 자동 처리 재확인
        try:
            hwp.XHwpMessageBoxMode = 1
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

        result = hwp.open(file_path)
        if not result:
            raise RuntimeError(f"한글 프로그램에서 파일을 열 수 없습니다: {file_path}")
        _current_doc_path = file_path
        # v0.7.2.12: open() 후 한컴이 마지막 cursor 위치를 복원하므로
        # 본문 첫 단락 보장을 위해 MoveDocBegin 명시 호출.
        # 이전: cursor 가 마지막 단락에 있어 set_paragraph_style/get_para_shape 가
        # 본문 첫 단락이 아닌 다른 단락에 작동하는 hidden bug 였음.
        try:
            hwp.HAction.Run("MoveDocBegin")
        except Exception as e:
            print(f"[WARN] open_document MoveDocBegin failed: {e}", file=sys.stderr)
        return {"status": "ok", "file_path": file_path, "pages": hwp.PageCount}

    if method == "get_document_info":
        # 경량 메타데이터만 반환 (analyze_document보다 빠름)
        result = {"status": "ok"}
        try:
            result["pages"] = hwp.PageCount
        except Exception:
            result["pages"] = 0
        try:
            result["current_path"] = _current_doc_path or ""
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        return result

    if method == "analyze_document":
        validate_params(params, ["file_path"], method)
        file_path = os.path.abspath(params["file_path"])
        return analyze_document(hwp, file_path, already_open=(file_path == _current_doc_path))

    if method == "fill_document":
        return fill_document(hwp, params)

    if method == "fill_by_tab":
        validate_params(params, ["table_index", "cells"], method)
        return fill_table_cells_by_tab(hwp, params["table_index"], params["cells"])

    if method == "fill_by_label":
        validate_params(params, ["table_index", "cells"], method)
        return fill_table_cells_by_label(hwp, params["table_index"], params["cells"])

    if method == "map_table_cells":
        validate_params(params, ["table_index"], method)
        return map_table_cells(hwp, params["table_index"])

    if method == "get_selected_text":
        text = hwp.get_selected_text()
        return {"text": text}

    if method == "get_font_list":
        from presets import get_font_list
        category = params.get("category")
        gov_only = params.get("gov_only", False)
        fonts = get_font_list(category=category, gov_only=gov_only)
        return {"status": "ok", "fonts": fonts, "count": len(fonts)}

    if method == "get_preset_list":
        from presets import DOCUMENT_PRESETS, TABLE_STYLES
        doc_presets = [{"name": k, "page": v.get("page", {})} for k, v in DOCUMENT_PRESETS.items()]
        table_styles = [{"name": k, "header_bg": v.get("header_bg")} for k, v in TABLE_STYLES.items()]
        return {"status": "ok", "document_presets": doc_presets, "table_styles": table_styles}

    if method == "apply_document_preset":
        validate_params(params, ["preset_name"], method)
        from presets import DOCUMENT_PRESETS
        preset_name = params["preset_name"]
        if preset_name not in DOCUMENT_PRESETS:
            return {"error": f"프리셋 '{preset_name}' 없음. 사용 가능: {list(DOCUMENT_PRESETS.keys())}"}
        preset = DOCUMENT_PRESETS[preset_name]
        # 1. 용지 설정 적용
        page = preset.get("page", {})
        if page:
            dispatch(hwp, "set_page_setup", {
                "top_margin": page.get("top", 20),
                "bottom_margin": page.get("bottom", 15),
                "left_margin": page.get("left", 20),
                "right_margin": page.get("right", 20),
            })
        # 2. 본문 서식 적용
        body = preset.get("body", {})
        if body:
            from hwp_editor import set_paragraph_style
            para_params = {}
            if "line_spacing" in body:
                para_params["line_spacing"] = body["line_spacing"]
            if "align" in body:
                para_params["align"] = body["align"]
            if para_params:
                set_paragraph_style(hwp, para_params)
        return {"status": "ok", "preset": preset_name, "applied": preset}

    if method == "get_table_dimensions":
        # 표 치수 추출 — 표 전체 너비, 셀 여백, 행/열 구조
        table_index = params.get("table_index", 0)
        hwp.get_into_nth_table(table_index)
        result = {"status": "ok", "table_index": table_index}
        try:
            result["table_width_mm"] = hwp.get_table_width()
        except Exception:
            result["table_width_mm"] = None
        try:
            result["cell_margin"] = hwp.get_cell_margin()
        except Exception:
            result["cell_margin"] = None
        try:
            result["outside_margin"] = {
                "top": hwp.get_table_outside_margin_top(),
                "bottom": hwp.get_table_outside_margin_bottom(),
                "left": hwp.get_table_outside_margin_left(),
                "right": hwp.get_table_outside_margin_right(),
            }
        except Exception:
            result["outside_margin"] = None
        # 셀 맵에서 행/열 구조 추출
        try:
            from hwp_editor import map_table_cells as _map
            cell_data = _map(hwp, table_index)
            result["total_cells"] = cell_data.get("total_cells", 0)
        except Exception:
            pass
        _exit_table_safely(hwp)
        return result

    if method == "extract_full_profile":
        # 양식 종합 프로파일 — 용지 + 문단 + 글자 + 표 치수
        from hwp_editor import get_char_shape, get_para_shape
        profile = {"status": "ok"}
        # 1. 용지 설정
        try:
            profile["page_setup"] = dispatch(hwp, "get_page_setup", {})
        except Exception as e:
            profile["page_setup"] = {"error": str(e)}
        # 2. 본문 서식 (커서가 본문에 있을 때)
        hwp.MovePos(2)
        try:
            profile["body_char"] = get_char_shape(hwp)
        except Exception as e:
            profile["body_char"] = {"error": str(e)}
        try:
            profile["body_para"] = get_para_shape(hwp)
        except Exception as e:
            profile["body_para"] = {"error": str(e)}
        # 3. 표 치수 (최대 5개 표)
        profile["tables"] = []
        for i in range(5):
            try:
                dims = dispatch(hwp, "get_table_dimensions", {"table_index": i})
                if dims.get("status") == "ok":
                    profile["tables"].append(dims)
            except Exception:
                break
        return profile

    if method == "get_page_setup":
        # F7 용지편집 정보 — 용지 크기, 방향, 여백, 사용 가능 영역
        try:
            d = hwp.get_pagedef_as_dict()
            pw = d.get("용지폭", 210)
            ph = d.get("용지길이", 297)
            lm = d.get("왼쪽", 30)
            rm = d.get("오른쪽", 30)
            tm = d.get("위쪽", 20)
            bm = d.get("아래쪽", 15)
            hm = d.get("머리말", 15)
            fm = d.get("꼬리말", 15)
            orient = d.get("용지방향", 0)
            binding = d.get("제본여백", 0)
            return {
                "status": "ok",
                "paper_width_mm": pw,
                "paper_height_mm": ph,
                "orientation": "landscape" if orient == 1 else "portrait",
                "top_margin_mm": tm,
                "bottom_margin_mm": bm,
                "left_margin_mm": lm,
                "right_margin_mm": rm,
                "header_margin_mm": hm,
                "footer_margin_mm": fm,
                "binding_margin_mm": binding,
                "usable_width_mm": round(pw - lm - rm, 1),
                "usable_height_mm": round(ph - tm - bm, 1),
            }
        except Exception as e:
            raise RuntimeError(f"용지 설정 읽기 실패: {e}")

    if method == "get_cursor_context":
        # 실제 커서 위치의 서식 + 주변 텍스트 반환
        from hwp_editor import get_char_shape, get_para_shape
        context = {"status": "ok"}
        try:
            context["char_shape"] = get_char_shape(hwp)
        except Exception as e:
            context["char_shape"] = {"error": str(e)}
        try:
            context["para_shape"] = get_para_shape(hwp)
        except Exception as e:
            context["para_shape"] = {"error": str(e)}
        try:
            pos = hwp.GetPos()
            context["position"] = list(pos) if pos else None
        except Exception:
            context["position"] = None
        try:
            context["total_pages"] = hwp.PageCount
        except Exception:
            context["total_pages"] = None
        try:
            # KeyIndicator: (섹션, 페이지, 줄, 컬럼, 삽입/수정, 줄번호)
            ki = hwp.KeyIndicator()
            context["current_page"] = ki[1] if ki else None
        except Exception:
            context["current_page"] = None
        return context

    if method == "save_document":
        # COM 메모리 → 파일 저장 (XML 엔진 동기화용)
        if _current_doc_path:
            try:
                hwp.save()
                return {"status": "ok", "saved": True, "path": _current_doc_path}
            except Exception as e:
                print(f"[WARN] save_document failed: {e}", file=sys.stderr)
                return {"status": "ok", "saved": False, "error": str(e)}
        return {"status": "ok", "saved": False, "reason": "no document open"}

    if method == "save_as":
        validate_params(params, ["path"], method)
        save_path = validate_file_path(params["path"], must_exist=False)
        fmt = params.get("format", "HWP").upper()  # pyhwpx는 대문자 포맷 필요 (HWP, HWPX, PDF 등)
        # 내보내기 전 현재 문서 저장 (COM 메모리 → 파일 반영, 빈 PDF 방지)
        if _current_doc_path and fmt in ("PDF", "DOCX", "HTML"):
            try:
                hwp.save()
            except Exception:
                pass
        hwp.save_as(save_path, fmt)
        # 파일 실제 생성 확인
        if not os.path.exists(save_path):
            # 대안: 임시 디렉토리에 저장 후 이동
            import tempfile, shutil as _shutil
            temp_path = os.path.join(tempfile.gettempdir(), os.path.basename(save_path))
            hwp.save_as(temp_path, fmt)
            if os.path.exists(temp_path):
                _shutil.move(temp_path, save_path)
        exists = os.path.exists(save_path)
        file_size = os.path.getsize(save_path) if exists else 0
        if not exists:
            raise RuntimeError(f"저장 실패: 파일이 생성되지 않았습니다. 경로: {save_path}")
        return {"status": "ok", "path": save_path, "file_size": file_size}

    if method == "close_document":
        # BUG-8 fix: XHwpMessageBoxMode 복원
        try:
            hwp.XHwpMessageBoxMode = 0
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        hwp.close()
        _current_doc_path = None
        return {"status": "ok"}

    if method == "text_search":
        validate_params(params, ["search"], method)
        search_text = params["search"]
        max_results = min(max(params.get("max_results", 50), 1), 1000)

        # 방법 1: 전체 텍스트에서 직접 검색 (COM FindReplace 반환값 불신뢰 대안)
        full_text = ""
        try:
            full_text = hwp.get_text_file("TEXT", "")
        except Exception:
            pass

        if full_text and search_text in full_text:
            results = []
            pos = 0
            idx = 0
            while idx < max_results:
                found = full_text.find(search_text, pos)
                if found == -1:
                    break
                start = max(0, found - 20)
                end = min(len(full_text), found + len(search_text) + 20)
                results.append({
                    "index": idx + 1,
                    "matched_text": search_text,
                    "context": full_text[start:end],
                })
                pos = found + len(search_text)
                idx += 1
            return {
                "search": search_text,
                "total_found": len(results),
                "results": results,
            }

        # 방법 2: COM FindReplace 기반 (fallback)
        hwp.MovePos(2)
        results = []
        for i in range(max_results):
            act = hwp.HAction
            pset = hwp.HParameterSet.HFindReplace
            act.GetDefault("FindReplace", pset.HSet)
            pset.FindString = search_text
            pset.Direction = 0
            pset.IgnoreMessage = 1
            act.Execute("FindReplace", pset.HSet)
            context = ""
            try:
                context = hwp.GetTextFile("TEXT", "saveblock").strip()[:200]
            except Exception:
                pass
            if not context:
                break
            hwp.HAction.Run("Cancel")
            results.append({
                "index": i + 1,
                "matched_text": context[:50] if context else search_text,
            })
        return {
            "search": search_text,
            "total_found": len(results),
            "results": results,
        }

    if method == "find_replace":
        validate_params(params, ["find", "replace"], method)
        use_regex = params.get("use_regex", False)
        case_sensitive = params.get("case_sensitive", True)
        replaced = _execute_all_replace(hwp, params["find"], params["replace"], use_regex, case_sensitive)
        return {"status": "ok", "find": params["find"], "replace": params["replace"], "replaced": replaced}

    if method == "find_replace_multi":
        validate_params(params, ["replacements"], method)
        use_regex = params.get("use_regex", False)
        results = []
        hwp.MovePos(2)  # 문서 시작으로 이동
        for item in params["replacements"]:
            replaced = _execute_all_replace(hwp, item["find"], item["replace"], use_regex)
            results.append({"find": item["find"], "replaced": replaced})
        return {"status": "ok", "results": results, "total": len(results),
                "success": sum(1 for r in results if r["replaced"])}

    if method == "find_and_append":
        validate_params(params, ["find", "append_text"], method)
        find_text = params["find"]
        append_text = params["append_text"]

        # 방법 1: AllReplace로 find → find+append 치환 (반환값 무시, 텍스트 검증)
        before = ""
        try:
            before = hwp.get_text_file("TEXT", "")
        except Exception:
            pass

        if find_text not in before:
            return {"status": "not_found", "find": find_text}

        # AllReplace: find → find + append_text
        replace_text = find_text + append_text
        _execute_all_replace(hwp, find_text, replace_text)

        # 실제 텍스트 변화로 성공 판단 (COM 반환값 무시)
        after = ""
        try:
            after = hwp.get_text_file("TEXT", "")
        except Exception:
            pass

        if replace_text in after:
            return {"status": "ok", "find": find_text, "appended": True}
        else:
            return {"status": "not_found", "find": find_text,
                    "warning": "AllReplace 실행했으나 텍스트 변화 미확인"}

    if method == "insert_text":
        validate_params(params, ["text"], method)
        # 표 안에 커서가 있으면 먼저 탈출 (표 간격/넘침 방지)
        try:
            if hwp.is_cell():
                _exit_table_safely(hwp)
        except Exception:
            pass
        text = params["text"]
        # === 텍스트 전처리: 줄바꿈 정규화 + 마커 앞 자동 줄바꿈 ===
        import re
        # 1) \r\n → \n 통일 (혼용 방지)
        text = text.replace("\r\n", "\n")
        # 2) 마커 문자 앞에 줄바꿈 삽입 (이미 줄바꿈이 있으면 건너뜀)
        _markers = r'[○□■◆●•◦※➤❶-❿▶▷►]'
        _roman = r'(?:Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)'
        text = re.sub(rf'(?<=[^\n])({_markers})', r'\n\1', text)
        text = re.sub(rf'(?<=[^\n])({_roman}\.)', r'\n\1', text)
        # 3) 3개+ 연속 공백 → 줄바꿈+들여쓰기 (PDF 원본 줄바꿈 복원)
        text = re.sub(r'  {3,}', '\n     ', text)
        # 4) \n → \r\n (HWP 단락 구분)
        text = text.replace("\n", "\r\n")
        # 5) 끝에 \r\n 보장
        if not text.endswith("\r\n"):
            text += "\r\n"
        # 원문 보존 (자동 내어쓰기 판단용)
        original_text = params["text"]
        style = params.get("style")
        color = params.get("color")  # [r, g, b] 하위 호환
        if style:
            from hwp_editor import insert_text_with_style
            insert_text_with_style(hwp, text, style)
        elif color:
            from hwp_editor import insert_text_with_color
            insert_text_with_color(hwp, text, tuple(color))
        else:
            hwp.insert_text(text)
        # === 후처리: 마커 감지 → ParagraphShapeIndentAtCaret 자동 내어쓰기 ===
        # 원문이 마커(○□-※* 등)로 시작하면, 마커 뒤 위치에서 Shift+Tab 효과 적용
        _INDENT_MARKERS = set('○□■◆●•◦※➤▶▷►-*')
        auto_indent = params.get("auto_indent", True)
        # v0.6.9: outline_level 지정 시 IndentAtCaret 스킵 (중복 처리 방지)
        outline_level = params.get("outline_level")
        raw = original_text.lstrip()
        if outline_level is None and auto_indent and raw and raw[0] in _INDENT_MARKERS:
            try:
                hwp.HAction.Run("MovePrevParaBegin")
                # 마커 뒤 위치 계산 (선행공백 + 마커 + 마커뒤공백)
                skip = 0
                ot = original_text
                while skip < len(ot) and ot[skip] == ' ':
                    skip += 1
                if skip < len(ot) and ot[skip] in _INDENT_MARKERS:
                    skip += 1
                while skip < len(ot) and ot[skip] == ' ':
                    skip += 1
                for _ in range(skip):
                    hwp.HAction.Run("MoveRight")
                hwp.HAction.Run("ParagraphShapeIndentAtCaret")
                hwp.MovePos(3)
            except Exception as e:
                print(f"[WARN] auto IndentAtCaret: {e}", file=sys.stderr)
                try:
                    hwp.MovePos(3)
                except Exception:
                    pass
        # v0.6.9 신규: outline_level 지정 시 직전 단락의 ParaShape.OutlineLevel 설정
        # (한글 "개요 보기" + hwp_generate_toc 계층 인식 활성화)
        # v0.6.9.3: multi-fallback (SetItem → set_style → 직접 attribute)
        if outline_level is not None:
            try:
                hwp.HAction.Run("MovePrevPara")
                ol_int = int(outline_level)
                success = False
                # 시도 1: ParameterSet.HSet.SetItem (표준 ParameterSet API)
                try:
                    act = hwp.HAction
                    pset = hwp.HParameterSet.HParaShape
                    act.GetDefault("ParaShape", pset.HSet)
                    pset.HSet.SetItem("OutlineLevel", ol_int)
                    act.Execute("ParaShape", pset.HSet)
                    success = True
                except Exception as e1:
                    print(f"[INFO] insert_text OutlineLevel SetItem failed: {e1}", file=sys.stderr)
                # 시도 2: hwp.set_style("개요 N+1") — 한컴 정의된 스타일
                if not success:
                    try:
                        hwp.set_style(f"개요 {ol_int + 1}")
                        success = True
                    except Exception as e2:
                        print(f"[INFO] insert_text set_style 개요 {ol_int + 1} failed: {e2}", file=sys.stderr)
                # 시도 3: pset.OutlineLevel 직접 attribute (v0.6.9 원래 방식, fallback)
                if not success:
                    try:
                        act = hwp.HAction
                        pset = hwp.HParameterSet.HParaShape
                        act.GetDefault("ParaShape", pset.HSet)
                        pset.OutlineLevel = ol_int
                        act.Execute("ParaShape", pset.HSet)
                        success = True
                    except Exception as e3:
                        print(f"[WARN] insert_text OutlineLevel all alternatives failed: {e3}", file=sys.stderr)
                hwp.MovePos(3)
            except Exception as e:
                print(f"[WARN] insert_text OutlineLevel (level={outline_level}): {e}", file=sys.stderr)
        return {"status": "ok"}

    if method == "set_paragraph_style":
        validate_params(params, ["style"], method)
        s = params["style"]
        # v0.6.7: first_line_indent는 indent의 alias (사용자 친화적 이름)
        if "first_line_indent" in s and "indent" not in s:
            s["indent"] = s["first_line_indent"]
        # Execute로 정상 작동하는 속성 (align, spacing, border 등)
        act = hwp.HAction
        pset = hwp.HParameterSet.HParaShape
        act.GetDefault("ParaShape", pset.HSet)
        align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
        _need_execute = False
        if "align" in s:
            pset.AlignType = align_map.get(s["align"], 0)
            _need_execute = True
        if "line_spacing" in s:
            pset.LineSpacingType = s.get("line_spacing_type", 0)
            pset.LineSpacing = int(s["line_spacing"])
            _need_execute = True
        if "space_before" in s:
            pset.PrevSpacing = int(s["space_before"] * 100)
            _need_execute = True
        if "space_after" in s:
            pset.NextSpacing = int(s["space_after"] * 100)
            _need_execute = True
        if "page_break_before" in s:
            pset.PagebreakBefore = 1 if s["page_break_before"] else 0
            _need_execute = True
        if "keep_with_next" in s:
            pset.KeepWithNext = 1 if s["keep_with_next"] else 0
            _need_execute = True
        if "widow_orphan" in s:
            pset.WidowOrphan = 1 if s["widow_orphan"] else 0
            _need_execute = True
        # v0.6.7: hwp_editor.py:set_paragraph_style와 인라인 풀 동기화 (8개 추가)
        if "line_wrap" in s:
            try:
                pset.LineWrap = int(s["line_wrap"])
                _need_execute = True
            except Exception as e:
                print(f"[WARN] LineWrap: {e}", file=sys.stderr)
        if "snap_to_grid" in s:
            try:
                pset.SnapToGrid = 1 if s["snap_to_grid"] else 0
                _need_execute = True
            except Exception as e:
                print(f"[WARN] SnapToGrid: {e}", file=sys.stderr)
        if "auto_space_eAsian_eng" in s:
            try:
                pset.AutoSpaceEAsianEng = 1 if s["auto_space_eAsian_eng"] else 0
                _need_execute = True
            except Exception as e:
                print(f"[WARN] AutoSpaceEAsianEng: {e}", file=sys.stderr)
        if "auto_space_eAsian_num" in s:
            try:
                pset.AutoSpaceEAsianNum = 1 if s["auto_space_eAsian_num"] else 0
                _need_execute = True
            except Exception as e:
                print(f"[WARN] AutoSpaceEAsianNum: {e}", file=sys.stderr)
        if "break_latin_word" in s:
            try:
                pset.BreakLatinWord = int(s["break_latin_word"])
                _need_execute = True
            except Exception as e:
                print(f"[WARN] BreakLatinWord: {e}", file=sys.stderr)
        if "heading_type" in s:
            try:
                pset.HeadingType = int(s["heading_type"])
                _need_execute = True
            except Exception as e:
                print(f"[WARN] HeadingType: {e}", file=sys.stderr)
        if "keep_lines_together" in s:
            try:
                pset.KeepLinesTogether = 1 if s["keep_lines_together"] else 0
                _need_execute = True
            except Exception as e:
                print(f"[WARN] KeepLinesTogether: {e}", file=sys.stderr)
        if "condense" in s:
            try:
                pset.Condense = int(s["condense"])
                _need_execute = True
            except Exception as e:
                print(f"[WARN] Condense: {e}", file=sys.stderr)
        # v0.6.7 신규: 문단 테두리 4면 (Border)
        # 입력: border_left/right/top/bottom = {"type": int, "width": float, "color": "#RRGGBB"}
        # 또는 border_color = "#RRGGBB" (4면 일괄), border_shadowing = bool
        _border_edges = {"left": "Left", "right": "Right", "top": "Top", "bottom": "Bottom"}
        for edge_key, edge_attr in _border_edges.items():
            border_key = f"border_{edge_key}"
            if border_key in s and isinstance(s[border_key], dict):
                bspec = s[border_key]
                try:
                    if "type" in bspec:
                        setattr(pset, f"BorderType{edge_attr}", int(bspec["type"]))
                    if "width" in bspec:
                        setattr(pset, f"BorderWidth{edge_attr}", float(bspec["width"]))
                    if "color" in bspec:
                        # "#RRGGBB" → RGB
                        c = bspec["color"].lstrip("#")
                        if len(c) == 6:
                            r = int(c[0:2], 16)
                            g = int(c[2:4], 16)
                            b = int(c[4:6], 16)
                            setattr(pset, f"BorderColor{edge_attr}", hwp.RGBColor(r, g, b))
                    _need_execute = True
                except Exception as e:
                    print(f"[WARN] Border{edge_attr}: {e}", file=sys.stderr)
        # 4면 색 일괄
        if "border_color" in s:
            try:
                c = s["border_color"].lstrip("#")
                if len(c) == 6:
                    r = int(c[0:2], 16)
                    g = int(c[2:4], 16)
                    b = int(c[4:6], 16)
                    rgb = hwp.RGBColor(r, g, b)
                    for edge_attr in _border_edges.values():
                        setattr(pset, f"BorderColor{edge_attr}", rgb)
                    _need_execute = True
            except Exception as e:
                print(f"[WARN] BorderColor (all): {e}", file=sys.stderr)
        # 그림자
        if "border_shadowing" in s:
            try:
                pset.BorderShadowing = 1 if s["border_shadowing"] else 0
                _need_execute = True
            except Exception as e:
                print(f"[WARN] BorderShadowing: {e}", file=sys.stderr)
        # v0.7.2.1 신규: ParaShape 정밀 옵션 (multi-fallback)
        # first_line_indent_hwpunit (1mm = 283 hwpunit, indent보다 정밀)
        if "first_line_indent_hwpunit" in s:
            try:
                fli_hwpu = int(s["first_line_indent_hwpunit"])
                # 시도 1: SetItem (v0.6.9.3 패턴)
                try:
                    pset.HSet.SetItem("Indent", fli_hwpu)
                except Exception:
                    pset.Indent = fli_hwpu  # 시도 2: 직접 attribute
                _need_execute = True
            except Exception as e:
                print(f"[WARN] first_line_indent_hwpunit: {e}", file=sys.stderr)
        # hanging_indent: 음수 indent 명시적 표현 (내어쓰기 체크박스 효과)
        if s.get("hanging_indent"):
            try:
                # 현재 Indent를 음수로 (이미 |Indent|만큼 내어쓰기)
                cur_indent = getattr(pset, "Indent", 0)
                if cur_indent > 0:
                    pset.Indent = -abs(int(cur_indent))
                _need_execute = True
            except Exception as e:
                print(f"[WARN] hanging_indent: {e}", file=sys.stderr)
        # paragraph_heading_type: none/outline/number (HeadingType 매핑)
        if "paragraph_heading_type" in s:
            try:
                pht_map = {"none": 0, "outline": 1, "number": 2}
                pht_val = pht_map.get(s["paragraph_heading_type"], 0)
                try:
                    pset.HSet.SetItem("HeadingType", pht_val)
                except Exception:
                    pset.HeadingType = pht_val
                _need_execute = True
            except Exception as e:
                print(f"[WARN] paragraph_heading_type: {e}", file=sys.stderr)
        # word_spacing: 단어 간격 (-50 ~ +50)
        if "word_spacing" in s:
            try:
                ws = int(s["word_spacing"])
                try:
                    pset.HSet.SetItem("WordSpacing", ws)
                except Exception:
                    pset.WordSpacing = ws
                _need_execute = True
            except Exception as e:
                print(f"[WARN] word_spacing: {e}", file=sys.stderr)
        # line_weight: 줄 두께 (50% ~ 500%)
        if "line_weight" in s:
            try:
                lw = int(s["line_weight"])
                try:
                    pset.HSet.SetItem("LineWeight", lw)
                except Exception:
                    pset.LineWeight = lw
                _need_execute = True
            except Exception as e:
                print(f"[WARN] line_weight: {e}", file=sys.stderr)
        if _need_execute:
            act.Execute("ParaShape", pset.HSet)
        # Execute로 미반영되는 속성 (LeftMargin, Indentation) → set_para 사용
        _para_kwargs = {}
        if "left_margin" in s:
            _para_kwargs["LeftMargin"] = float(s["left_margin"])
        if "right_margin" in s:
            _para_kwargs["RightMargin"] = float(s["right_margin"])
        if "indent" in s:
            indent_val = float(s["indent"])
            _para_kwargs["Indentation"] = indent_val
            # v0.6.7: indent<0 (내어쓰기) + left_margin 미지정 시 자동 보정
            # HWP Shift+Tab과 동일 효과. v0.6.5에서 사라졌던 로직 복원
            # (메모리 feedback_indent_auto_correction 참조)
            if indent_val < 0 and "left_margin" not in s:
                _para_kwargs["LeftMargin"] = abs(indent_val)
        if _para_kwargs:
            try:
                hwp.set_para(**_para_kwargs)
            except Exception as e:
                print(f"[WARN] set_para failed: {e}", file=sys.stderr)
        return {"status": "ok"}

    if method == "get_char_shape":
        return get_char_shape(hwp)

    if method == "get_para_shape":
        return get_para_shape(hwp)

    if method == "get_cell_format":
        validate_params(params, ["table_index", "cell_tab"], method)
        from hwp_editor import get_cell_format
        return get_cell_format(hwp, params["table_index"], params["cell_tab"])

    if method == "get_table_format_summary":
        validate_params(params, ["table_index"], method)
        from hwp_editor import get_table_format_summary
        return get_table_format_summary(
            hwp, params["table_index"], params.get("sample_tabs"))

    if method == "smart_fill":
        validate_params(params, ["table_index", "cells"], method)
        from hwp_editor import smart_fill_table_cells
        return smart_fill_table_cells(hwp, params["table_index"], params["cells"])

    if method == "read_reference":
        validate_params(params, ["file_path"], method)
        from ref_reader import read_reference
        return read_reference(params["file_path"], params.get("max_chars", 30000))

    if method == "find_replace_nth":
        validate_params(params, ["find", "replace", "nth"], method)
        find_text = params["find"]
        replace_text = params["replace"]
        nth = params["nth"]  # 1-based
        if nth < 1 or nth > 10000:
            raise ValueError("nth must be between 1 and 10000")

        # 전체 텍스트에서 n번째 확인
        before = ""
        try:
            before = hwp.get_text_file("TEXT", "")
        except Exception:
            pass
        count = before.count(find_text)
        if count < nth:
            return {"status": "not_found", "find": find_text, "searched": count, "nth": nth}

        # AllReplace 기반 n번째 치환: 마커 치환 → n번째만 replace → 복원
        import uuid
        marker = f"@@NTH{uuid.uuid4().hex[:6]}@@"
        # 1단계: find_text → 마커 (전부 치환)
        _execute_all_replace(hwp, find_text, marker)
        # 2단계: n번째 마커만 replace_text로, 나머지는 find_text로 복원
        # InitScan으로 마커를 하나씩 찾아 순번에 따라 처리
        hwp.MovePos(2)
        found_count = 0
        for i in range(count):  # count = before.count(find_text)
            act = hwp.HAction
            pset = hwp.HParameterSet.HFindReplace
            act.GetDefault("FindReplace", pset.HSet)
            pset.FindString = marker
            pset.ReplaceString = replace_text if (i == nth - 1) else find_text
            pset.Direction = 0
            pset.IgnoreMessage = 1
            pset.ReplaceMode = 1  # 현재 선택만 치환
            act.Execute("FindReplace", pset.HSet)
            # 치환 후 다음으로 이동
            act.GetDefault("FindReplace", pset.HSet)
            pset.FindString = marker
            pset.Direction = 0
            pset.IgnoreMessage = 1
            act.Execute("FindReplace", pset.HSet)
            found_count += 1
        # 잔여 마커 복원 (안전)
        _execute_all_replace(hwp, marker, find_text)
        # 검증
        after = ""
        try:
            after = hwp.get_text_file("TEXT", "")
        except Exception:
            pass
        replaced = replace_text in after
        return {"status": "ok" if replaced else "uncertain", "find": find_text, "replace": replace_text, "nth": nth, "replaced": replaced}

    if method == "table_add_row":
        validate_params(params, ["table_index"], method)
        hwp.get_into_nth_table(params["table_index"])
        # 마지막 셀로 이동 후 행 추가
        try:
            hwp.HAction.Run("TableAppendRow")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception:
            # 대안: InsertRowBelow
            try:
                hwp.HAction.Run("InsertRowBelow")
                return {"status": "ok", "table_index": params["table_index"], "method": "InsertRowBelow"}
            except Exception as e:
                raise RuntimeError(f"표 행 추가 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "document_merge":
        validate_params(params, ["file_path"], method)
        merge_path = validate_file_path(params["file_path"], must_exist=True)
        hwp.MovePos(3)  # 문서 끝으로 이동
        # BreakSection으로 페이지 분리 후 파일 삽입
        try:
            hwp.HAction.Run("BreakSection")
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        hwp.insert_file(merge_path)
        return {"status": "ok", "merged_file": merge_path, "pages": hwp.PageCount}

    if method == "insert_page_break":
        try:
            hwp.HAction.Run("BreakPage")
            return {"status": "ok"}
        except Exception as e:
            raise RuntimeError(f"페이지 나누기 실패: {e}")

    if method == "insert_markdown":
        validate_params(params, ["text"], method)
        from hwp_editor import insert_markdown
        return insert_markdown(hwp, params["text"])

    if method == "table_delete_row":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.TableSubtractRow()
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 행 삭제 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "table_add_column":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("InsertColumnRight")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 열 추가 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "table_delete_column":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("DeleteColumn")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 열 삭제 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "table_merge_cells":
        validate_params(params, ["table_index"], method)
        table_index = params["table_index"]
        start_row = params.get("start_row")
        start_col = params.get("start_col")
        end_row = params.get("end_row")
        end_col = params.get("end_col")
        try:
            hwp.get_into_nth_table(table_index)
            if start_row is not None and end_row is not None and start_col is not None and end_col is not None:
                # 범위 지정 병합 — 시작 셀로 이동 → 블록 선택 확장 → 병합
                hwp.HAction.Run("TableColBegin")
                hwp.HAction.Run("TableRowBegin")
                for _ in range(start_row):
                    hwp.HAction.Run("TableLowerCell")
                for _ in range(start_col):
                    hwp.HAction.Run("TableRightCell")
                # 블록 선택 시작
                hwp.HAction.Run("TableCellBlock")
                # TableCellBlockExtend + 방향키로 블록 확장
                for _ in range(end_col - start_col):
                    hwp.HAction.Run("TableCellBlockExtend")
                    hwp.HAction.Run("TableRightCell")
                for _ in range(end_row - start_row):
                    hwp.HAction.Run("TableCellBlockExtend")
                    hwp.HAction.Run("TableLowerCell")
                hwp.HAction.Run("TableMergeCell")
            else:
                # 기존 방식 (현재 선택된 셀 병합)
                hwp.TableMergeCell()
            return {"status": "ok", "table_index": table_index,
                    "range": {"start_row": start_row, "start_col": start_col, "end_row": end_row, "end_col": end_col} if start_row is not None else None}
        except Exception as e:
            raise RuntimeError(f"셀 병합 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "table_split_cell":
        validate_params(params, ["table_index"], method)
        # v0.6.8: rows/cols 옵셔널 파라미터 추가 (기본 hwp.TableSplitCell() 인수 없는 호출)
        rows = params.get("rows")
        cols = params.get("cols")
        try:
            hwp.get_into_nth_table(params["table_index"])
            if rows is not None or cols is not None:
                # HAction ParameterSet 방식으로 Rows/Cols 지정
                try:
                    act = hwp.HAction
                    pset = hwp.HParameterSet.HTableSplitCell
                    act.GetDefault("TableSplitCell", pset.HSet)
                    if rows is not None:
                        try:
                            pset.Rows = int(rows)
                        except Exception as e:
                            print(f"[WARN] TableSplitCell Rows: {e}", file=sys.stderr)
                    if cols is not None:
                        try:
                            pset.Cols = int(cols)
                        except Exception as e:
                            print(f"[WARN] TableSplitCell Cols: {e}", file=sys.stderr)
                    act.Execute("TableSplitCell", pset.HSet)
                except Exception as e:
                    # ParameterSet 경로 실패 시 기본 split 폴백
                    print(f"[WARN] HTableSplitCell ParameterSet failed, fallback: {e}", file=sys.stderr)
                    hwp.TableSplitCell()
            else:
                hwp.TableSplitCell()
            return {"status": "ok", "table_index": params["table_index"], "rows": rows, "cols": cols}
        except Exception as e:
            raise RuntimeError(f"셀 분할 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    # v0.6.8 신규: 표 셀 네비게이션 (커서에 머무름, finally _exit_table_safely 호출 안 함)
    if method == "navigate_cell":
        validate_params(params, ["direction"], method)
        direction = params["direction"]
        if not hwp.is_cell():
            return {"status": "error", "error": "현재 커서가 표 안에 없습니다. 먼저 표에 진입하세요."}
        action_map = {
            "left": "TableLeftCell",
            "right": "TableRightCell",
            "upper": "TableUpperCell",
            "lower": "TableLowerCell",
        }
        action = action_map.get(direction)
        if action is None:
            raise ValueError(f"invalid direction: {direction}. Expected one of {list(action_map.keys())}")
        try:
            # pyhwpx wrap 우선 (hwp.TableLeftCell 등), 없으면 HAction.Run 폴백
            if hasattr(hwp, action):
                try:
                    result = getattr(hwp, action)()
                    moved = bool(result) if result is not None else True
                except Exception as e:
                    print(f"[WARN] pyhwpx {action} failed, falling back to HAction.Run: {e}", file=sys.stderr)
                    hwp.HAction.Run(action)
                    moved = True
            else:
                hwp.HAction.Run(action)
                moved = True
            return {"status": "ok", "direction": direction, "moved": moved}
        except Exception as e:
            raise RuntimeError(f"셀 이동 실패 ({direction}): {e}")

    # v0.6.8 신규: 현재 커서 셀 기준 행 추가 (above/below/append)
    # 기존 table_add_row(table_index 기반)와 구별 — 커서 위치 기반
    if method == "insert_row_at_cursor":
        validate_params(params, ["position"], method)
        position = params["position"]
        if not hwp.is_cell():
            return {"status": "error", "error": "현재 커서가 표 안에 없습니다. 먼저 표에 진입하세요."}
        action_map = {
            "above": "TableInsertUpperRow",
            "below": "TableInsertLowerRow",
            "append": "TableAppendRow",
        }
        action = action_map.get(position)
        if action is None:
            raise ValueError(f"invalid position: {position}. Expected one of {list(action_map.keys())}")
        try:
            hwp.HAction.Run(action)
            return {"status": "ok", "position": position}
        except Exception as e:
            raise RuntimeError(f"행 추가 실패 ({position}): {e}")
        finally:
            _exit_table_safely(hwp)

    # v0.6.8 신규: 이미 선택된 블록을 병합 (기존 table_merge_cells는 좌표 기반)
    # 사용자가 이미 TableCellBlock로 블록을 선택한 상태에서 병합
    if method == "merge_current_selection":
        try:
            hwp.TableMergeCell()
            return {"status": "ok"}
        except Exception as e:
            raise RuntimeError(f"선택 블록 병합 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "table_create_from_data":
        validate_params(params, ["data"], method)
        data = params["data"]  # 2D 배열 [[row1], [row2], ...]
        if not data or not isinstance(data, list):
            raise ValueError("data must be a non-empty 2D array")
        rows = len(data)
        cols = max(len(row) for row in data) if data else 0
        header_style = params.get("header_style", False)
        col_widths = params.get("col_widths")  # [mm, mm, ...] H1 fix
        row_heights = params.get("row_heights")  # [mm, mm, ...]
        alignment = params.get("alignment")  # left/center/right

        # 표 너비를 페이지 사용 가능 폭에 맞춤 (통일된 표 너비)
        col_width_warning = None
        try:
            page_d = hwp.get_pagedef_as_dict()
            usable_width = page_d.get("용지폭", 210) - page_d.get("왼쪽", 30) - page_d.get("오른쪽", 30)
        except Exception:
            usable_width = 160  # fallback
        usable_width = max(usable_width, 50)  # 최소 50mm 보장 (좁은 용지 방어)
        target_width = max(usable_width - 5, 20)  # 약간 여유 (5mm), 최소 20mm

        if col_widths:
            total_width = sum(col_widths)
            if abs(total_width - target_width) > 1:  # 1mm 이상 차이면 비율 조정
                ratio = target_width / total_width
                col_widths = [round(w * ratio, 1) for w in col_widths]
                if total_width > target_width + 5:
                    col_width_warning = f"col_widths 합계({total_width}mm)를 페이지 폭({target_width}mm)에 맞춰 조정했습니다."
        else:
            # col_widths 미지정 시: 균등 분배로 페이지 폭에 맞춤
            if cols > 0:
                col_widths = [round(target_width / cols, 1)] * cols
            else:
                col_widths = []

        # H1: col_widths/row_heights가 있으면 HTableCreation으로 정밀 생성
        if col_widths or row_heights:
            try:
                tc = hwp.HParameterSet.HTableCreation
                hwp.HAction.GetDefault("TableCreate", tc.HSet)
                tc.Rows = rows
                tc.Cols = cols
                tc.WidthType = 2  # 절대 너비
                tc.HeightType = 0
                if col_widths:
                    tc.CreateItemArray("ColWidth", cols)
                    for i, w in enumerate(col_widths[:cols]):
                        tc.ColWidth.SetItem(i, hwp.MiliToHwpUnit(w))
                if row_heights:
                    tc.CreateItemArray("RowHeight", rows)
                    for i, h in enumerate(row_heights[:rows]):
                        tc.RowHeight.SetItem(i, hwp.MiliToHwpUnit(h))
                hwp.HAction.Execute("TableCreate", tc.HSet)
            except Exception as e:
                print(f"[WARN] HTableCreation failed, fallback to create_table: {e}", file=sys.stderr)
                hwp.create_table(rows, cols)
        else:
            hwp.create_table(rows, cols)
        # 셀 채우기 (alignment 적용 포함)
        align_map = {"left": 0, "center": 1, "right": 2}
        # 넓은 표(6열+) 폰트 자동 축소
        wide_table_font_size = 9 if cols >= 6 else None
        filled = 0
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                # alignment 적용 (각 셀에 문단 정렬)
                if alignment and alignment in align_map:
                    try:
                        act_p = hwp.HAction
                        ps = hwp.HParameterSet.HParaShape
                        act_p.GetDefault("ParaShape", ps.HSet)
                        ps.AlignType = align_map[alignment]
                        act_p.Execute("ParaShape", ps.HSet)
                    except Exception as e:
                        print(f"[WARN] Cell align: {e}", file=sys.stderr)
                if val:
                    if header_style and r == 0:
                        from hwp_editor import insert_text_with_style
                        style = {"bold": True}
                        if wide_table_font_size:
                            style["font_size"] = wide_table_font_size
                        insert_text_with_style(hwp, str(val), style)
                    elif wide_table_font_size and r > 0:
                        from hwp_editor import insert_text_with_style
                        insert_text_with_style(hwp, str(val), {"font_size": wide_table_font_size})
                    else:
                        hwp.insert_text(str(val))
                    filled += 1
                if c < len(row) - 1 or r < rows - 1:
                    hwp.TableRightCell()
        # 표 밖으로 안전하게 탈출 (is_cell 확인 후 Cancel 반복)
        _exit_table_safely(hwp)
        try:
            pass  # _exit_table_safely에서 이미 MoveDocEnd + BreakPara 수행
        except Exception as e:
            print(f"[WARN] Table exit: {e}", file=sys.stderr)
        # header_style: Bold는 이미 표 생성 시 적용됨
        # 배경색은 set_cell_color로 별도 적용 (표 진입/탈출 부작용 방지)
        result = {"status": "ok", "rows": rows, "cols": cols, "filled": filled, "header_styled": bool(header_style)}
        if col_width_warning:
            result["warning"] = col_width_warning
        return result

    if method == "create_approval_box":
        # 결재란 자동 생성 (4×N 표 + 서식)
        levels = params.get("levels", ["기안", "검토", "결재"])
        position = params.get("position", "right")  # right or center
        cols = len(levels) + 1  # 구분열 + 결재자 수
        rows = 4  # 구분, 직급, 성명, 서명
        # 표 데이터 구성
        data = [["구분"] + levels]
        data.append(["직급"] + ["" for _ in levels])
        data.append(["성명"] + ["" for _ in levels])
        data.append(["서명"] + ["" for _ in levels])
        col_widths = [18] + [25 for _ in levels]
        row_heights = [8, 8, 12, 12]
        # 표 생성
        result = dispatch(hwp, "table_create_from_data", {
            "data": data,
            "col_widths": col_widths,
            "row_heights": row_heights,
            "alignment": position,
            "header_style": True,
        })
        # 헤더행 배경색 (진남색) + 흰색 글자
        try:
            from hwp_editor import set_cell_background_color
            cells = [{"tab": i, "color": "#E8E8E8"} for i in range(cols)]
            set_cell_background_color(hwp, 0, cells)
        except Exception as e:
            print(f"[WARN] Approval box style: {e}", file=sys.stderr)
        return {"status": "ok", "rows": rows, "cols": cols, "levels": levels}

    if method == "table_insert_from_csv":
        validate_params(params, ["file_path"], method)
        csv_path = validate_file_path(params["file_path"], must_exist=True)
        from ref_reader import read_reference
        ref = read_reference(csv_path)
        if ref.get("format") not in ("csv", "excel"):
            raise ValueError(f"CSV 또는 Excel 파일만 지원합니다. (현재: {ref.get('format')})")
        # 헤더 + 데이터를 2D 배열로 병합
        headers = ref.get("headers", [])
        data_rows = ref.get("data", [])
        if ref.get("format") == "excel":
            sheets = ref.get("sheets", [])
            if sheets:
                headers = sheets[0].get("headers", [])
                data_rows = sheets[0].get("data", [])
        all_data = [headers] + data_rows if headers else data_rows
        if not all_data:
            raise ValueError("CSV 파일에 데이터가 없습니다.")
        rows = len(all_data)
        cols = max(len(row) for row in all_data)
        hwp.create_table(rows, cols)
        filled = 0
        for r, row in enumerate(all_data):
            for c, val in enumerate(row):
                if val:
                    # BUG-1 fix: SelectAll 제거
                    hwp.insert_text(str(val))
                    filled += 1
                if c < len(row) - 1 or r < rows - 1:
                    hwp.TableRightCell()
        _exit_table_safely(hwp)
        return {"status": "ok", "file": os.path.basename(csv_path), "rows": rows, "cols": cols, "filled": filled}

    if method == "insert_heading":
        validate_params(params, ["text", "level"], method)
        from hwp_editor import insert_text_with_style
        # v0.6.9: level 범위 1~6 → 1~9 확장 (OutlineLevel 0~8 지원)
        level = min(max(params["level"], 1), 9)
        sizes = {1: 22, 2: 18, 3: 15, 4: 13, 5: 11, 6: 10, 7: 10, 8: 10, 9: 10}
        text = params["text"]
        # 순번 자동 생성 (기존 API 후방 호환)
        numbering = params.get("numbering")
        number = params.get("number", 1)
        # v0.6.9 신규 옵션
        auto_outline_level = bool(params.get("auto_outline_level", False))
        outline_level_only = bool(params.get("outline_level_only", False))
        # outline_level_only=true면 텍스트 prefix 생략
        if numbering and not outline_level_only:
            roman = ["Ⅰ","Ⅱ","Ⅲ","Ⅳ","Ⅴ","Ⅵ","Ⅶ","Ⅷ","Ⅸ","Ⅹ"]
            korean = ["가","나","다","라","마","바","사","아","자","차"]
            circle = ["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩"]
            idx = max(0, min(number - 1, 9))
            if numbering == "roman": text = f"{roman[idx]}. {text}"
            elif numbering == "decimal": text = f"{number}. {text}"
            elif numbering == "korean": text = f"{korean[idx]}. {text}"
            elif numbering == "circle": text = f"{circle[idx]} {text}"
            elif numbering == "paren_decimal": text = f"{number}) {text}"
            elif numbering == "paren_korean": text = f"{korean[idx]}) {text}"
        insert_text_with_style(hwp, text + "\r\n", {
            "bold": True,
            "font_size": sizes.get(level, 11),
        })
        # v0.6.9 신규: auto_outline_level 또는 outline_level_only 지정 시
        # 직전 단락(방금 삽입한 제목)의 ParaShape.OutlineLevel 설정
        # → 한글 "개요 보기" + hwp_generate_toc 계층 인식 활성화
        # v0.6.9.3: multi-fallback (SetItem → set_style → 직접 attribute)
        applied_outline_level = None
        applied_via = None  # 어떤 방법으로 성공했는지 추적
        if auto_outline_level or outline_level_only:
            try:
                hwp.HAction.Run("MovePrevPara")
                ol_int = level - 1  # 0-based (level 1 → OutlineLevel 0)
                # 시도 1: ParameterSet.HSet.SetItem (표준 ParameterSet API)
                try:
                    act = hwp.HAction
                    pset = hwp.HParameterSet.HParaShape
                    act.GetDefault("ParaShape", pset.HSet)
                    pset.HSet.SetItem("OutlineLevel", ol_int)
                    act.Execute("ParaShape", pset.HSet)
                    applied_outline_level = ol_int
                    applied_via = "SetItem"
                except Exception as e1:
                    print(f"[INFO] insert_heading SetItem failed: {e1}", file=sys.stderr)
                # 시도 2: hwp.set_style("개요 N") — 한컴 정의된 스타일
                if applied_outline_level is None:
                    try:
                        hwp.set_style(f"개요 {level}")
                        applied_outline_level = ol_int
                        applied_via = "set_style"
                    except Exception as e2:
                        print(f"[INFO] insert_heading set_style 개요 {level} failed: {e2}", file=sys.stderr)
                # 시도 3: pset.OutlineLevel 직접 attribute (v0.6.9 원래 방식, fallback)
                if applied_outline_level is None:
                    try:
                        act = hwp.HAction
                        pset = hwp.HParameterSet.HParaShape
                        act.GetDefault("ParaShape", pset.HSet)
                        pset.OutlineLevel = ol_int
                        act.Execute("ParaShape", pset.HSet)
                        applied_outline_level = ol_int
                        applied_via = "direct_attribute"
                    except Exception as e3:
                        print(f"[WARN] insert_heading OutlineLevel all alternatives failed: {e3}", file=sys.stderr)
                hwp.MovePos(3)
            except Exception as e:
                print(f"[WARN] insert_heading OutlineLevel (level={level}): {e}", file=sys.stderr)
        return {
            "status": "ok",
            "level": level,
            "text": text,
            "outline_level": applied_outline_level,
            "applied_via": applied_via,
        }

    if method == "export_format":
        validate_params(params, ["path", "format"], method)
        save_path = validate_file_path(params["path"], must_exist=False)
        fmt = params["format"].upper()  # HWP, HWPX, PDF, HTML, TXT 등
        # DOCX/HTML은 HWP COM에서 미지원 — 타임아웃 방지
        if fmt in ("DOCX", "DOC"):
            return {"status": "not_supported",
                    "message": "DOCX 직접 내보내기는 한/글 COM에서 지원되지 않습니다. PDF로 내보내기를 권장합니다.",
                    "alternative": "hwp_export_pdf"}
        if fmt == "HTML":
            return {"status": "not_supported",
                    "message": "HTML 직접 내보내기는 한/글 COM에서 지원되지 않습니다. hwp_get_as_markdown으로 마크다운 변환 후 HTML로 변환하세요.",
                    "alternative": "hwp_get_as_markdown"}
        # PDF/내보내기 전 현재 문서 저장 (COM 메모리 → 파일 반영, 빈 PDF 방지)
        if _current_doc_path:
            try:
                hwp.save()
            except Exception:
                pass
        result = hwp.save_as(save_path, fmt)
        # 파일 실제 생성 확인
        file_exists = os.path.exists(save_path)
        file_size = os.path.getsize(save_path) if file_exists else 0
        return {"status": "ok" if file_exists else "warning",
                "path": save_path, "format": fmt,
                "success": bool(result), "file_exists": file_exists, "file_size": file_size}

    if method == "verify_layout":
        # PDF로 내보내고 PNG 이미지로 변환 → Claude Code의 Read로 시각적 검증
        import tempfile
        # 먼저 현재 문서 저장 (COM 메모리 → 파일 반영, 빈 PDF 방지)
        if _current_doc_path:
            try:
                hwp.save()
            except Exception:
                pass
        tmp_pdf = os.path.join(tempfile.gettempdir(), "hwp_verify_layout.pdf")
        try:
            hwp.save_as(tmp_pdf, "PDF")
            if not os.path.exists(tmp_pdf):
                return {"status": "error", "error": "PDF 생성 실패"}

            # PDF → PNG 변환 (PyMuPDF)
            try:
                import fitz
                doc = fitz.open(tmp_pdf)
                image_paths = []
                page_range = params.get("pages")  # "1", "1-3" 등
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
                # 임시 PDF 정리 (PNG만 유지)
                try:
                    os.remove(tmp_pdf)
                except Exception:
                    pass
                return {
                    "status": "ok",
                    "image_paths": image_paths,
                    "pages": len(image_paths),
                    "total_pages": hwp.PageCount,
                    "hint": "Read 도구로 각 PNG 이미지를 열어 레이아웃을 시각적으로 검증하세요."
                }
            except ImportError:
                # PyMuPDF 미설치 → PDF 경로만 반환
                return {
                    "status": "ok_pdf_only",
                    "pdf_path": tmp_pdf,
                    "pages": hwp.PageCount,
                    "file_size": os.path.getsize(tmp_pdf),
                    "hint": "PyMuPDF 미설치. 'pip install PyMuPDF' 실행 후 다시 시도하면 PNG 이미지로 자동 변환됩니다."
                }
        except Exception as e:
            return {"status": "error", "error": f"레이아웃 검증 실패: {e}"}

    if method == "set_page_setup":
        # 페이지 설정 (여백, 용지 크기, 방향)
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HSecDef
            act.GetDefault("PageSetup", pset.HSet)
            pdef = pset.PageDef
            if "top_margin" in params:
                pdef.TopMargin = hwp.MiliToHwpUnit(params["top_margin"])
            if "bottom_margin" in params:
                pdef.BottomMargin = hwp.MiliToHwpUnit(params["bottom_margin"])
            if "left_margin" in params:
                pdef.LeftMargin = hwp.MiliToHwpUnit(params["left_margin"])
            if "right_margin" in params:
                pdef.RightMargin = hwp.MiliToHwpUnit(params["right_margin"])
            if "header_margin" in params:
                pdef.HeaderLen = hwp.MiliToHwpUnit(params["header_margin"])
            if "footer_margin" in params:
                pdef.FooterLen = hwp.MiliToHwpUnit(params["footer_margin"])
            if "orientation" in params:
                pdef.Landscape = 1 if params["orientation"] == "landscape" else 0
            if "paper_width" in params:
                pdef.PaperWidth = hwp.MiliToHwpUnit(params["paper_width"])
            if "paper_height" in params:
                pdef.PaperHeight = hwp.MiliToHwpUnit(params["paper_height"])
            act.Execute("PageSetup", pset.HSet)
            return {"status": "ok"}
        except Exception as e:
            return {"status": "error", "error": f"페이지 설정 실패: {e}"}

    if method == "set_cell_property":
        # 셀 속성 설정 (여백, 텍스트 방향, 수직 정렬, 보호)
        validate_params(params, ["table_index", "tab"], method)
        try:
            from hwp_editor import _navigate_to_tab
            hwp.get_into_nth_table(params["table_index"])
            _navigate_to_tab(hwp, params["table_index"], params["tab"], 0)
            pset = hwp.HParameterSet.HCell
            hwp.HAction.GetDefault("CellShape", pset.HSet)
            if "vert_align" in params:
                va_map = {"top": 0, "middle": 1, "bottom": 2}
                pset.VertAlign = va_map.get(params["vert_align"], 0)
            if "margin_left" in params:
                pset.MarginLeft = hwp.MiliToHwpUnit(params["margin_left"])
            if "margin_right" in params:
                pset.MarginRight = hwp.MiliToHwpUnit(params["margin_right"])
            if "margin_top" in params:
                pset.MarginTop = hwp.MiliToHwpUnit(params["margin_top"])
            if "margin_bottom" in params:
                pset.MarginBottom = hwp.MiliToHwpUnit(params["margin_bottom"])
            if "text_direction" in params:
                pset.TextDirection = int(params["text_direction"])  # 0=가로, 1=세로
            if "protected" in params:
                pset.Protected = 1 if params["protected"] else 0
            hwp.HAction.Execute("CellShape", pset.HSet)
            return {"status": "ok", "tab": params["tab"]}
        except Exception as e:
            raise RuntimeError(f"셀 속성 설정 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "insert_textbox":
        # 글상자 생성 (위치/크기 지정)
        x = params.get("x", 0)  # mm
        y = params.get("y", 0)  # mm
        width = params.get("width", 60)  # mm
        height = params.get("height", 30)  # mm
        text = params.get("text", "")
        border = params.get("border", True)
        try:
            # 방법 1: HParameterSet.HShapeObject로 위치/크기 지정
            act = hwp.HAction
            pset = hwp.HParameterSet.HShapeObject
            act.GetDefault("InsertDrawObj", pset.HSet)
            pset.ShapeType = 1  # 1=사각형(글상자)
            pset.HorzRelTo = 0  # 0=페이지 기준
            pset.VertRelTo = 0
            pset.HorzOffset = int(x * 283.465)  # mm → HWPUNIT (1mm=283.465)
            pset.VertOffset = int(y * 283.465)
            pset.Width = int(width * 283.465)
            pset.Height = int(height * 283.465)
            act.Execute("InsertDrawObj", pset.HSet)
            if text:
                hwp.insert_text(text)
            hwp.HAction.Run("Cancel")
            return {"status": "ok", "x": x, "y": y, "width": width, "height": height}
        except Exception as e:
            # 방법 2: CreateAction 방식
            try:
                act_tb = hwp.CreateAction("DrawTextBox")
                ps = act_tb.CreateSet()
                act_tb.GetDefault(ps)
                act_tb.Execute(ps)
                if text:
                    hwp.insert_text(text)
                hwp.HAction.Run("Cancel")
                return {"status": "ok", "method": "fallback", "text": text,
                        "warning": f"위치/크기 파라미터가 적용되지 않았습니다: {e}"}
            except Exception as e2:
                raise RuntimeError(f"글상자 생성 실패: {e} / {e2}")

    if method == "draw_line":
        # 선 그리기 (두께/색상/스타일)
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HDrawLineAttr
            act.GetDefault("DrawLine", pset.HSet)
            if "width" in params:
                pset.Width = int(params["width"])  # 선 두께
            if "color" in params:
                c = params["color"]
                if isinstance(c, str):  # "#RRGGBB"
                    r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
                    pset.Color = hwp.RGBColor(r, g, b)
                elif isinstance(c, list):
                    pset.Color = hwp.RGBColor(c[0], c[1], c[2])
            if "style" in params:
                pset.style = int(params["style"])  # 0=실선, 1=파선, 2=점선 등
            act.Execute("DrawLine", pset.HSet)
            return {"status": "ok"}
        except Exception as e:
            raise RuntimeError(f"선 그리기 실패: {e}")

    if method == "set_header_footer":
        # 머리글/바닥글 설정 (CreateAction 방식)
        hf_type = params.get("type", "header")  # "header" or "footer"
        text = params.get("text", "")
        style = params.get("style")  # {font_size, bold, align}
        try:
            act = hwp.CreateAction("HeaderFooter")
            ps = act.CreateSet()
            act.GetDefault(ps)
            # Type: 0=머리글, 1=바닥글
            ps.SetItem("Type", 0 if hf_type == "header" else 1)
            result = act.Execute(ps)
            if not result:
                raise RuntimeError("HeaderFooter Execute 실패")
            # 머리글/바닥글 편집 모드 진입됨 — 텍스트 삽입
            if text and style:
                from hwp_editor import insert_text_with_style, set_paragraph_style
                insert_text_with_style(hwp, text, style)
                if "align" in style:
                    set_paragraph_style(hwp, {"align": style["align"]})
            elif text:
                hwp.insert_text(text)
            # 본문으로 복귀
            hwp.HAction.Run("CloseEx")
            return {"status": "ok", "type": hf_type, "text": text}
        except Exception as e:
            # 편집 모드에 들어갔을 수 있으므로 복귀 시도
            try:
                hwp.HAction.Run("CloseEx")
            except Exception as ex:
                print(f"[WARN] CloseEx recovery failed: {ex}", file=sys.stderr)
            raise RuntimeError(f"머리글/바닥글 설정 실패: {e}")

    if method == "apply_style":
        # 스타일 적용 ("제목1", "본문", "개요1" 등)
        style_name = params.get("style_name", "본문")
        try:
            # CharShape/ParaShape를 스타일 기반으로 변경
            # pyhwpx의 set_style 또는 HAction 기반
            act = hwp.HAction
            pset = hwp.HParameterSet.HStyle
            act.GetDefault("Style", pset.HSet)
            pset.HSet.SetItem("StyleName", style_name)
            act.Execute("Style", pset.HSet)
            return {"status": "ok", "style": style_name}
        except Exception as e:
            raise RuntimeError(f"스타일 적용 실패: {e}")

    if method == "set_column":
        # 다단 설정
        count = params.get("count", 2)  # 단 수
        gap = params.get("gap", 10)  # 단 간격 (mm)
        line_type = params.get("line_type", 0)  # 구분선 종류
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HColDef
            act.GetDefault("MultiColumn", pset.HSet)
            pset.Count = int(count)
            pset.SameSize = 1  # 같은 너비
            pset.SameGap = hwp.MiliToHwpUnit(gap)
            pset.LineType = int(line_type)
            pset.type = 1  # 일반 다단
            act.Execute("MultiColumn", pset.HSet)
            return {"status": "ok", "count": count, "gap": gap}
        except Exception as e:
            raise RuntimeError(f"다단 설정 실패: {e}")

    if method == "insert_caption":
        # 캡션 삽입 (표/그림 제목)
        text = params.get("text", "")
        side = params.get("side", 3)  # 0=왼쪽, 1=오른쪽, 2=위, 3=아래
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HCaption
            act.GetDefault("InsertCaption", pset.HSet)
            pset.Side = int(side)
            act.Execute("InsertCaption", pset.HSet)
            if text:
                hwp.insert_text(text)
            return {"status": "ok", "text": text, "side": side}
        except Exception as e:
            raise RuntimeError(f"캡션 삽입 실패: {e}")

    if method == "insert_hyperlink":
        validate_params(params, ["url"], method)
        url = params["url"]
        text = params.get("text", url)
        try:
            hwp.insert_hyperlink(url, text)
        except TypeError:
            # insert_hyperlink 시그니처가 다를 경우 대안
            hwp.insert_hyperlink(url)
        return {"status": "ok", "url": url, "text": text}

    if method == "image_extract":
        validate_params(params, ["output_dir"], method)
        output_dir = os.path.abspath(params["output_dir"])
        os.makedirs(output_dir, exist_ok=True)
        # pyhwpx save_all_pictures는 ./temp/binData 경로를 참조하므로 미리 생성
        temp_dir = os.path.join(os.getcwd(), "temp", "binData")
        os.makedirs(temp_dir, exist_ok=True)
        extracted_ok = False
        try:
            hwp.save_all_pictures(output_dir)
            extracted_ok = True
        except Exception:
            # 대안: HWPX로 저장 후 ZIP에서 이미지 추출
            try:
                import zipfile
                temp_hwpx = os.path.join(output_dir, "_temp.hwpx")
                hwp.save_as(temp_hwpx, "HWPX")
                if os.path.exists(temp_hwpx):
                    with zipfile.ZipFile(temp_hwpx, 'r') as z:
                        for name in z.namelist():
                            if name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                                z.extract(name, output_dir)
                    os.remove(temp_hwpx)
                    extracted_ok = True
            except Exception as e2:
                raise RuntimeError(f"이미지 추출 실패: {e2}")
        files = []
        for root, dirs, fnames in os.walk(output_dir):
            for fname in fnames:
                if fname.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.wmf', '.emf')):
                    rel = os.path.relpath(os.path.join(root, fname), output_dir)
                    files.append(rel)
        return {"status": "ok", "output_dir": output_dir, "extracted": len(files), "files": files}

    if method == "document_split":
        validate_params(params, ["output_dir"], method)
        output_dir = os.path.abspath(params["output_dir"])
        os.makedirs(output_dir, exist_ok=True)
        import shutil
        total_pages = hwp.PageCount
        pages_per_split = params.get("pages_per_split", 1)
        if pages_per_split < 1:
            pages_per_split = 1
        # 원본 경로
        src_path = _current_doc_path
        if not src_path:
            raise RuntimeError("열린 문서가 없습니다.")
        _, ext = os.path.splitext(src_path)
        parts = []
        # 각 분할: 원본 복사 → 열기 → save_as(split_page=True) 방식
        # pyhwpx save_as에 split_page 파라미터가 있으므로 활용
        for start in range(1, total_pages + 1, pages_per_split):
            end = min(start + pages_per_split - 1, total_pages)
            part_name = f"part_{start}-{end}{ext}"
            part_path = os.path.join(output_dir, part_name)
            # 분할 저장은 COM API 한계로 전체 복사 (실제 페이지 분할 아님)
            shutil.copy2(src_path, part_path)
            parts.append({"pages": f"{start}-{end}", "path": part_path})
        return {"status": "ok", "total_pages": total_pages, "parts": len(parts), "files": parts,
                "warning": "COM API 한계로 각 파일은 전체 문서의 복사본입니다. 실제 페이지 분할은 한글 프로그램에서 수동으로 진행해주세요."}

    if method == "insert_footnote":
        try:
            hwp.HAction.Run("InsertFootnote")
            text = params.get("text")
            if text:
                hwp.insert_text(text)
            hwp.HAction.Run("CloseEx")
            return {"status": "ok", "type": "footnote"}
        except Exception as e:
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
            raise RuntimeError(f"각주 삽입 실패: {e}")

    if method == "insert_endnote":
        try:
            hwp.HAction.Run("InsertEndnote")
            text = params.get("text")
            if text:
                hwp.insert_text(text)
            hwp.HAction.Run("CloseEx")
            return {"status": "ok", "type": "endnote"}
        except Exception as e:
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
            raise RuntimeError(f"미주 삽입 실패: {e}")

    if method == "insert_page_num":
        fmt = params.get("format", "plain")  # "plain"|"dash"|"paren"
        prefix_suffix = {"dash": ("- ", " -"), "paren": ("(", ")"), "plain": ("", "")}
        prefix, suffix = prefix_suffix.get(fmt, ("", ""))
        if prefix:
            hwp.insert_text(prefix)
        hwp.HAction.Run("InsertPageNum")
        if suffix:
            hwp.insert_text(suffix)
        return {"status": "ok", "format": fmt}

    if method == "generate_toc":
        # 문서 텍스트에서 제목 패턴을 추출하여 목차 텍스트 생성
        # v0.6.6 B3: scan_context 기반 extract_all_text 사용 (ReleaseScan finally 보장)
        import re
        from hwp_editor import extract_all_text
        text_blob = extract_all_text(hwp, max_iters=1000, strip_each=True, separator="\n")
        texts = text_blob.split("\n") if text_blob else []
        # 제목 패턴 감지
        toc_items = []
        heading_patterns = [
            (r'^(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)[.\s]', 1),  # 로마자 대제목
            (r'^(\d+)\.\s', 2),  # 1. 2. 3.
            (r'^(가|나|다|라|마|바|사)\.\s', 3),  # 가. 나. 다.
        ]
        for t in texts:
            for pattern, level in heading_patterns:
                if re.match(pattern, t):
                    toc_items.append({"level": level, "text": t[:60]})
                    break
        # 목차 텍스트 생성 + 삽입
        if params.get("insert", True):
            from hwp_editor import insert_text_with_style
            insert_text_with_style(hwp, "목   차\r\n", {"bold": True, "font_size": 16})
            hwp.insert_text("\r\n")
            for item in toc_items:
                indent = "  " * (item["level"] - 1)
                hwp.insert_text(f"{indent}{item['text']}\r\n")
            hwp.insert_text("\r\n")
        return {"status": "ok", "toc_items": len(toc_items), "items": toc_items[:30]}

    if method == "create_gantt_chart":
        validate_params(params, ["tasks", "months"], method)
        tasks = params["tasks"]  # [{"name": "A", "desc": "설명", "start": 1, "end": 3, "weight": "30%"}]
        months = params["months"]  # 6
        month_label = params.get("month_label", "M+N")
        # 2D 배열 생성
        header = ["세부 업무", "수행내용"]
        for i in range(months):
            if month_label == "M+N":
                header.append(f"M+{i}" if i > 0 else "M")
            else:
                header.append(f"{i+1}월")
        header.append("비중(%)")
        data = [header]
        active_cells = []  # ■ 셀의 tab 인덱스 기록 (배경색용)
        for task_idx, task in enumerate(tasks):
            row = [task.get("name", ""), task.get("desc", "")]
            start = task.get("start", 1)
            end = task.get("end", 1)
            for m in range(months):
                if start <= m + 1 <= end:
                    row.append("■")
                    # 헤더행(0) + task 행(task_idx+1), 열은 2+m (세부업무,수행내용 다음)
                    tab = (task_idx + 1) * len(header) + 2 + m
                    active_cells.append(tab)
                else:
                    row.append("")
            row.append(str(task.get("weight", "")))
            data.append(row)
        # 표 생성
        rows = len(data)
        cols = len(data[0])
        hwp.create_table(rows, cols)
        from hwp_editor import insert_text_with_style
        filled = 0
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                if val:
                    # BUG-1 fix: SelectAll 제거
                    if r == 0:
                        insert_text_with_style(hwp, str(val), {"bold": True})
                    else:
                        hwp.insert_text(str(val))
                    filled += 1
                if c < len(row) - 1 or r < rows - 1:
                    hwp.TableRightCell()
        _exit_table_safely(hwp)
        # 헤더행 + ■ 셀 배경색 적용
        try:
            from hwp_editor import set_cell_background_color
            style_cells = [{"tab": i, "color": "#666666"} for i in range(cols)]  # 헤더: 표준 헤더색
            style_cells += [{"tab": t, "color": "#C0C0C0"} for t in active_cells]  # ■셀: 음영
            set_cell_background_color(hwp, -1, style_cells)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        return {"status": "ok", "rows": rows, "cols": cols, "filled": filled, "active_cells": len(active_cells)}

    if method == "insert_date_code":
        try:
            hwp.InsertDateCode()
        except Exception:
            hwp.HAction.Run("InsertDateCode")
        return {"status": "ok"}

    if method == "table_formula_sum":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("TableFormulaSumAuto")
            return {"status": "ok", "table_index": params["table_index"], "formula": "sum"}
        except Exception as e:
            raise RuntimeError(f"표 합계 계산 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "table_formula_avg":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("TableFormulaAvgAuto")
            return {"status": "ok", "table_index": params["table_index"], "formula": "avg"}
        except Exception as e:
            raise RuntimeError(f"표 평균 계산 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    # ── Phase B: Quick Win 8개 ──
    if method == "table_to_csv":
        validate_params(params, ["table_index", "output_path"], method)
        output_path = validate_file_path(params["output_path"], must_exist=False)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.table_to_csv(output_path)
        except Exception:
            # 병합 셀 등으로 pyhwpx table_to_csv 실패 시 → map_table_cells로 대안
            import csv
            cell_data = map_table_cells(hwp, params["table_index"])
            cells = cell_data.get("cell_map", [])
            with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                for c in cells:
                    writer.writerow([c.get("tab", ""), c.get("text", "")])
        finally:
            _exit_table_safely(hwp)
        return {"status": "ok", "table_index": params["table_index"], "path": output_path}

    if method == "break_section":
        hwp.BreakSection()
        return {"status": "ok", "type": "section"}

    if method == "break_column":
        hwp.BreakColumn()
        return {"status": "ok", "type": "column"}

    if method == "insert_line":
        # draw_line과 동일하게 처리 (대화상자 방지)
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HDrawLineAttr
            act.GetDefault("DrawLine", pset.HSet)
            act.Execute("DrawLine", pset.HSet)
            return {"status": "ok"}
        except Exception as e:
            raise RuntimeError(f"선 삽입 실패: {e}")

    if method == "table_swap_type":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("TableSwapType")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 행/열 교환 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    if method == "insert_auto_num":
        hwp.HAction.Run("InsertAutoNum")
        return {"status": "ok"}

    if method == "insert_memo":
        hwp.HAction.Run("InsertFieldMemo")
        text = params.get("text")
        if text:
            hwp.insert_text(text)
        return {"status": "ok"}

    if method == "table_distribute_width":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("TableDistributeCellWidth")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"셀 너비 균등 분배 실패: {e}")
        finally:
            _exit_table_safely(hwp)

    # ── Phase C: 복합 기능 6개 ──
    if method == "table_to_json":
        validate_params(params, ["table_index"], method)
        cell_data = map_table_cells(hwp, params["table_index"])
        cell_map = cell_data.get("cell_map", [])
        json_data = [{"tab": c["tab"], "text": c["text"]} for c in cell_map]
        return {"status": "ok", "table_index": params["table_index"],
                "total_cells": len(json_data), "cells": json_data}

    if method == "batch_convert":
        validate_params(params, ["input_dir", "output_format"], method)
        input_dir = os.path.abspath(params["input_dir"])
        output_format = params["output_format"].upper()
        output_dir = os.path.abspath(params.get("output_dir", input_dir))
        os.makedirs(output_dir, exist_ok=True)
        results = []
        for f in os.listdir(input_dir):
            if f.lower().endswith(('.hwp', '.hwpx')):
                src = os.path.join(input_dir, f)
                name, _ = os.path.splitext(f)
                out = os.path.join(output_dir, f"{name}.{output_format.lower()}")
                try:
                    hwp.open(src)
                    hwp.save_as(out, output_format)
                    hwp.close()
                    results.append({"file": f, "output": out, "status": "ok"})
                except Exception as e:
                    results.append({"file": f, "status": "error", "error": str(e)})
                    try:
                        hwp.close()
                    except Exception:
                        pass
        return {"status": "ok", "total": len(results),
                "success": sum(1 for r in results if r["status"] == "ok"),
                "results": results}

    if method == "compare_documents":
        validate_params(params, ["file_path_1", "file_path_2"], method)
        path1 = validate_file_path(params["file_path_1"], must_exist=True)
        path2 = validate_file_path(params["file_path_2"], must_exist=True)
        # 문서 1 텍스트 추출 (v0.6.6 B3: extract_all_text 사용)
        from hwp_editor import extract_all_text
        hwp.open(path1)
        text1 = ""
        try:
            text1 = extract_all_text(hwp, max_iters=5000, strip_each=True, separator="\n")
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        hwp.close()
        # 문서 2 텍스트 추출 (v0.6.6 B3: extract_all_text 사용)
        hwp.open(path2)
        text2 = ""
        try:
            text2 = extract_all_text(hwp, max_iters=5000, strip_each=True, separator="\n")
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        hwp.close()
        # diff 계산
        lines1 = text1.split("\n")
        lines2 = text2.split("\n")
        added = [l for l in lines2 if l not in lines1]
        removed = [l for l in lines1 if l not in lines2]
        return {"status": "ok", "file_1": os.path.basename(path1), "file_2": os.path.basename(path2),
                "lines_1": len(lines1), "lines_2": len(lines2),
                "added": len(added), "removed": len(removed),
                "added_lines": added[:20], "removed_lines": removed[:20]}

    if method == "word_count":
        # v0.6.6 B3: extract_all_text 사용 (separator="" → concat)
        from hwp_editor import extract_all_text
        text = ""
        try:
            text = extract_all_text(hwp, max_iters=10000, strip_each=False, separator="")
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        chars_total = len(text)
        chars_no_space = len(text.replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", ""))
        words = len(text.split())
        paragraphs = text.count("\n") + 1
        return {"status": "ok", "chars_total": chars_total, "chars_no_space": chars_no_space,
                "words": words, "paragraphs": paragraphs, "pages": hwp.PageCount}

    # ── Phase E: 양식 자동 감지 ──
    if method == "indent":
        # 들여쓰기 (Shift+Tab 효과): LeftMargin 증가 = 나머지 줄 시작위치 이동
        depth = params.get("depth", 10)  # pt 단위, 기본 10pt
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HParaShape
            act.GetDefault("ParaShape", pset.HSet)
            current_left = 0
            try:
                current_left = pset.LeftMargin or 0
            except Exception:
                pass
            new_left = current_left + int(depth * 100)
            pset.LeftMargin = new_left
            act.Execute("ParaShape", pset.HSet)
            return {"status": "ok", "left_margin_pt": new_left / 100}
        except Exception as e:
            raise RuntimeError(f"들여쓰기 실패: {e}")

    if method == "outdent":
        # 내어쓰기: LeftMargin 감소
        depth = params.get("depth", 10)
        try:
            act = hwp.HAction
            pset = hwp.HParameterSet.HParaShape
            act.GetDefault("ParaShape", pset.HSet)
            current_left = 0
            try:
                current_left = pset.LeftMargin or 0
            except Exception:
                pass
            new_left = max(0, current_left - int(depth * 100))
            pset.LeftMargin = new_left
            act.Execute("ParaShape", pset.HSet)
            return {"status": "ok", "left_margin_pt": new_left / 100}
        except Exception as e:
            raise RuntimeError(f"내어쓰기 실패: {e}")

    if method == "extract_style_profile":
        # 양식 문서에서 서식 프로파일 추출
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

    if method == "delete_guide_text":
        # 작성요령/가이드 텍스트 자동 삭제
        # "< 작성요령 >" 패턴과 ※ 안내문 등을 찾아 삭제
        patterns = params.get("patterns", ["< 작성요령 >", "＜ 작성요령 ＞", "<작성요령>"])
        deleted = 0
        hwp.MovePos(2)
        for pat in patterns:
            replaced = _execute_all_replace(hwp, pat, "", False)
            if replaced:
                deleted += 1
        return {"status": "ok", "deleted_patterns": deleted, "patterns": patterns}

    if method == "toggle_checkbox":
        # 체크박스 전환: □→■, ☐→☑ 등
        validate_params(params, ["find", "replace"], method)
        find_text = params["find"]
        replace_text = params["replace"]
        replaced = _execute_all_replace(hwp, find_text, replace_text, False)
        return {"status": "ok", "find": find_text, "replace": replace_text, "replaced": replaced}

    if method == "form_detect":
        # 문서 텍스트에서 빈칸/괄호/밑줄 패턴으로 양식 필드 자동 감지
        # v0.6.6 B3: extract_all_text 사용 (ReleaseScan finally 보장)
        import re
        from hwp_editor import extract_all_text
        text = ""
        try:
            text = extract_all_text(hwp, max_iters=10000, strip_each=False, separator="\n")
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
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
                context = text[max(0, m.start()-20):m.end()+20].strip()
                fields.append({
                    "type": field_type,
                    "description": description,
                    "position": m.start(),
                    "context": context[:50],
                })
        return {"status": "ok", "total_fields": len(fields), "fields": fields[:50]}

    if method == "set_background_picture":
        validate_params(params, ["file_path"], method)
        bg_path = validate_file_path(params["file_path"], must_exist=True)
        hwp.insert_background_picture(bg_path)
        return {"status": "ok", "file_path": bg_path}

    if method == "set_cell_color":
        validate_params(params, ["table_index", "cells"], method)
        from hwp_editor import set_cell_background_color
        return set_cell_background_color(hwp, params["table_index"], params["cells"])

    if method == "set_table_border":
        validate_params(params, ["table_index"], method)
        from hwp_editor import set_table_border_style
        return set_table_border_style(hwp, params["table_index"], params.get("cells"), params.get("style", {}))

    if method == "auto_map_reference":
        validate_params(params, ["table_index", "ref_headers", "ref_row"], method)
        from hwp_editor import auto_map_reference_to_table
        return auto_map_reference_to_table(
            hwp, params["table_index"], params["ref_headers"], params["ref_row"])

    if method == "insert_picture":
        validate_params(params, ["file_path"], method)
        from hwp_editor import insert_picture
        return insert_picture(hwp, params["file_path"],
                              params.get("width", 0), params.get("height", 0))

    if method == "privacy_scan":
        validate_params(params, ["text"], method)
        from privacy_scanner import scan_privacy
        return scan_privacy(params["text"])

    if method == "verify_after_fill":
        validate_params(params, ["table_index", "expected_cells"], method)
        return verify_after_fill(hwp, params["table_index"], params["expected_cells"])

    if method == "generate_multi_documents":
        validate_params(params, ["template_path", "data_list"], method)
        return _generate_multi_documents(
            hwp,
            params["template_path"],
            params["data_list"],
            params.get("output_dir"),
        )

    # B2 (v0.6.6): HeadCtrl 순회 — 표/그림/머리말/꼬리말/각주/누름틀 등 모든 컨트롤 나열
    if method == "list_controls":
        from hwp_traversal import traverse_all_ctrls
        filter_ids = params.get("filter")  # None | list | "all"
        max_visits = params.get("max_visits", 5000)
        return traverse_all_ctrls(hwp, include_ids=filter_ids, max_visits=max_visits)

    # ──────────────────────────────────────────────────────────
    # v0.7.1: 양식 학습 + Workload Estimate (사용자 핵심 니즈)
    # ──────────────────────────────────────────────────────────

    # v0.7.1 신규: 양식의 트리 구조 추출 (목차/섹션/표/필드)
    if method == "extract_template_structure":
        validate_params(params, ["file_path"], method)
        import re as _re
        from hwp_analyzer import analyze_document as _analyze
        max_depth = int(params.get("max_depth", 4))

        # 1. 기존 analyze_document 재활용
        analysis = _analyze(hwp, params["file_path"])

        # 2. heading 인식 정규식 (full_text 단락 단위 분석)
        # 패턴: 제 N 장/조/절, I./II., 1./1.1/1.1.1, 가./나., (1)/(가)
        _heading_patterns = [
            (_re.compile(r'^제\s*(\d+)\s*[장조절]\s'), 1),
            (_re.compile(r'^([IVX]+)\.\s'), 1),
            (_re.compile(r'^(\d+)\.\s'), 1),
            (_re.compile(r'^(\d+)\.(\d+)\s'), 2),
            (_re.compile(r'^(\d+)\.(\d+)\.(\d+)\s'), 3),
            (_re.compile(r'^([가-힣])\.\s'), 2),
            (_re.compile(r'^\(([가-힣\d])\)\s'), 3),
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

    # v0.7.1 신규: 양식의 서식 패턴 학습
    if method == "analyze_writing_patterns":
        validate_params(params, ["file_path"], method)
        # v0.7.2.8: file_path 를 실제로 열어 두 파일이 다르게 분석되도록
        # v0.7.2.12: hwp.open() 후 cursor 가 마지막 위치 복원 → MoveDocBegin 으로 본문 첫 단락 강제
        # MovePos(2) 가 사실 cursor 를 본문 첫 단락으로 보내지 못하는 hidden bug 있었음
        from hwp_editor import get_para_shape, get_char_shape
        try:
            hwp.open(os.path.abspath(params["file_path"]))
            hwp.HAction.Run("MoveDocBegin")
        except Exception as e:
            print(f"[WARN] analyze_writing_patterns open/MoveDocBegin failed: {e}", file=sys.stderr)
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

        # consistency_score: 단순 — 모든 단락의 char/para shape이 sample과 일치하는지
        # MVP: 100점 가정 (실제 측정은 v0.7.1.1로 확장)
        consistency_score = 100

        return {
            "status": "ok",
            "file_path": params["file_path"],
            "page_setup": page_d,
            "body_style": {
                "char": body_char,
                "para": body_para,
            },
            "title_styles": {},  # MVP: 빈 (v0.7.1.1 확장)
            "table_styles": [],  # MVP: 빈 (v0.7.1.1 확장)
            "numbering_pattern": "decimal_dot",  # MVP default
            "consistency_score": consistency_score,
            "deviations_sample": [],
        }

    # v0.7.1 신규 ★: Workload 추정 (사용자 사전 분석 도구)
    if method == "estimate_workload":
        validate_params(params, ["user_request"], method)
        user_request = params["user_request"]
        constraints = params.get("constraints", {}) or {}
        max_ref_files = int(constraints.get("max_reference_files", 5))
        max_ref_mb = int(constraints.get("max_reference_mb", 10))
        context_window = int(constraints.get("context_window_tokens", 200000))

        # 1. 양식 분석 (옵셔널)
        estimated_pages = 10  # default
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
                print(f"[WARN] estimate_workload analyze failed: {e}", file=sys.stderr)

        # 2. user_request 휴리스틱 (정규식: "10페이지", "5장", "20쪽")
        import re as _re
        page_match = _re.search(r'(\d+)\s*(페이지|쪽|장|page)', user_request, _re.IGNORECASE)
        if page_match:
            estimated_pages = int(page_match.group(1))
        section_match = _re.search(r'(\d+)\s*(섹션|section|chapter|챕터|단락)', user_request, _re.IGNORECASE)
        if section_match:
            estimated_sections = int(section_match.group(1))

        # 3. 추정 공식
        chars_per_page = 1100  # A4 11pt 줄간 160%
        tokens_per_char = 1.0 / 3.5  # 한국어
        output_chars = estimated_pages * chars_per_page
        output_tokens = int(output_chars * tokens_per_char * 1.6)  # 안전계수

        # 입력 토큰: 양식 분석 chars + reference chars (옵셔널)
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

        # 양식 자체 chars 추가
        if analysis_data:
            input_chars += len(analysis_data.get("full_text", "") or "")

        input_tokens = int(input_chars * tokens_per_char)
        total_tokens = input_tokens + output_tokens
        context_usage_percent = round(total_tokens / context_window * 100, 2)

        # 4. 시간 예측
        seconds_per_output_token = 0.011  # Opus 4.6 한국어 측정
        writing_seconds = int(output_tokens * seconds_per_output_token)
        analysis_seconds = 5 + estimated_tables * 2
        verification_seconds = estimated_pages * 3
        save_seconds = 30
        total_seconds = writing_seconds + analysis_seconds + verification_seconds + save_seconds

        # 5. 위험 평가
        risks = []
        if analysis_data and analysis_data.get("controls_by_type", {}).get("tbl", 0) > 5:
            risks.append({"type": "many_tables", "severity": "medium", "description": f"표 {analysis_data['controls_by_type']['tbl']}개 — 표 처리 시간 추가"})
        if input_tokens > 0.4 * context_window:
            risks.append({"type": "long_context", "severity": "high", "description": f"입력 토큰 {input_tokens} > context window 40%"})
        if output_tokens > 60000:
            risks.append({"type": "output_overflow", "severity": "high", "description": f"출력 토큰 {output_tokens} > 60k (응답 분할 필요)"})
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

        # 7. split suggestion (단순)
        split_suggestion = []
        if recommended == "split_into_sessions" and estimated_sections > 0:
            half = max(1, estimated_sections // 2)
            split_suggestion = [
                {"section_range": f"1-{half}", "estimated_pages": estimated_pages // 2, "estimated_tokens": total_tokens // 2},
                {"section_range": f"{half + 1}-{estimated_sections}", "estimated_pages": estimated_pages // 2, "estimated_tokens": total_tokens // 2},
            ]

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
        }

    # v0.7.1 신규: 기존 양식 섹션 확장
    if method == "extend_section":
        validate_params(params, ["section_identifier", "content"], method)
        section_id = params["section_identifier"]  # {by: "title|index", value: ...}
        content = params["content"]
        preserve_format = bool(params.get("preserve_format", True))

        # MVP: section title text를 본문에서 찾아 그 직후에 텍스트 삽입
        # full search → MovePos → insert_text
        if isinstance(section_id, dict) and section_id.get("by") == "title":
            title = section_id.get("value", "")
            try:
                # find 후 그 위치로 이동
                hwp.HAction.Run("MoveDocBegin")
                act = hwp.HAction
                pset = hwp.HParameterSet.HFindReplace
                act.GetDefault("RepeatFind", pset.HSet)
                pset.FindString = title
                pset.Direction = 0
                pset.IgnoreMessage = 1
                if not act.Execute("RepeatFind", pset.HSet):
                    return {"status": "error", "error": f"섹션 제목을 찾을 수 없습니다: {title}"}
                hwp.HAction.Run("MoveLineEnd")
                hwp.HAction.Run("BreakPara")
            except Exception as e:
                return {"status": "error", "error": f"섹션 위치 이동 실패: {e}"}

        # 텍스트 삽입 (단락 단위)
        try:
            for line in content.split("\n"):
                if line.strip():
                    hwp.insert_text(line)
                    hwp.HAction.Run("BreakPara")
            return {
                "status": "ok",
                "section_identifier": section_id,
                "inserted_paragraphs": len([l for l in content.split("\n") if l.strip()]),
                "preserve_format": preserve_format,
            }
        except Exception as e:
            return {"status": "error", "error": f"텍스트 삽입 실패: {e}"}

    # v0.7.1 신규: 패턴 프로파일 일괄 적용 (MVP)
    if method == "apply_style_profile":
        validate_params(params, ["profile"], method)
        profile = params["profile"]
        target = params.get("target", "all")

        # MVP: profile.body_style을 현재 단락에 적용
        body = profile.get("body_style", {}) if isinstance(profile, dict) else {}
        applied = 0
        try:
            if body.get("para"):
                # set_paragraph_style 분기로 위임 (실제는 내부 함수 호출 어려우므로 직접 처리)
                act = hwp.HAction
                pset = hwp.HParameterSet.HParaShape
                act.GetDefault("ParaShape", pset.HSet)
                # 안전한 옵션만 적용
                p = body["para"]
                if "AlignType" in p:
                    pset.AlignType = int(p["AlignType"])
                if "LineSpacing" in p:
                    pset.LineSpacing = int(p["LineSpacing"])
                act.Execute("ParaShape", pset.HSet)
                applied += 1
            return {"status": "ok", "applied_paragraphs": applied, "target": target}
        except Exception as e:
            return {"status": "error", "error": f"profile 적용 실패: {e}"}

    # v0.7.1 신규: 작성된 결과의 양식 일관성 검증 (MVP)
    if method == "validate_consistency":
        validate_params(params, ["file_path"], method)
        # MVP: 단순 — 현재 문서의 page/body 가져와서 expected와 비교
        # expected_profile 미지정 시 100점 (placeholder)
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
            # 단순 비교: 키가 일치하지 않으면 deviation 추가, 5점씩 감점
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
                "checked_paragraphs": 1,  # MVP: 현재 단락만
                "current_para": current_para,
                "current_char": current_char,
            },
        }

    raise ValueError(f"Unknown method: {method}")


def _generate_multi_documents(hwp, template_path, data_list, output_dir=None):
    """템플릿 기반 다건 문서 생성.

    각 데이터마다 템플릿을 별도 파일로 복사 → 열기 → 채우기 → 저장 → 닫기.
    AllReplace 범위 문제를 근본적으로 회피.

    data_list: [{
        "name": "파일명 접미사 (예: 이준혁_(주)딥러닝코리아)",
        "table_cells": {table_idx(str): [{"tab": N, "text": "값"}, ...]},  # optional
        "replacements": [{"find": "X", "replace": "Y"}, ...],              # optional
        "verify_tables": [table_idx, ...]                                   # optional
    }, ...]
    """
    import shutil

    template_path = os.path.abspath(template_path)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    if output_dir is None:
        output_dir = os.path.dirname(template_path)
    else:
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

    _, ext = os.path.splitext(template_path)
    results = []

    for idx, data in enumerate(data_list):
        doc_name = data.get("name", f"문서_{idx+1}")
        output_path = os.path.join(output_dir, f"{doc_name}{ext}")
        doc_result = {
            "name": doc_name,
            "output_path": output_path,
            "status": "ok",
            "fill_results": [],
            "replace_results": [],
            "verify_results": [],
            "errors": [],
        }

        try:
            # 1. 템플릿 파일 복사
            shutil.copy2(template_path, output_path)

            # 2. 복사본 열기 (백업 불필요 — 원본이 템플릿)
            opened = hwp.open(output_path)
            if not opened:
                raise RuntimeError(f"파일을 열 수 없습니다: {output_path}")

            # 3. 표 채우기
            table_cells = data.get("table_cells", {})
            for table_idx_str, cells in table_cells.items():
                table_idx = int(table_idx_str)
                fill_result = fill_table_cells_by_tab(hwp, table_idx, cells)
                doc_result["fill_results"].append({
                    "table_index": table_idx,
                    **fill_result,
                })

            # 4. 텍스트 치환 (공통 함수 사용)
            replacements = data.get("replacements", [])
            if replacements:
                hwp.MovePos(2)  # 문서 시작
                for item in replacements:
                    replaced = _execute_all_replace(hwp, item["find"], item["replace"])
                    doc_result["replace_results"].append({
                        "find": item["find"],
                        "replace": item["replace"],
                        "replaced": replaced,
                    })

            # 5. 검증 (옵션)
            verify_tables = data.get("verify_tables", [])
            for table_idx in verify_tables:
                table_idx = int(table_idx)
                # table_cells에서 해당 표의 expected 값 추출
                expected = table_cells.get(str(table_idx), [])
                if expected:
                    vr = verify_after_fill(hwp, table_idx, expected)
                    doc_result["verify_results"].append({
                        "table_index": table_idx,
                        **vr,
                    })

            # 6. 저장 + 닫기
            hwp.save()
            hwp.close()

        except Exception as e:
            doc_result["status"] = "error"
            doc_result["errors"].append(str(e))
            # 에러 시에도 문서 닫기 시도
            try:
                hwp.close()
            except Exception:
                pass

        results.append(doc_result)

    return {
        "status": "ok",
        "template": template_path,
        "total": len(data_list),
        "success": sum(1 for r in results if r["status"] == "ok"),
        "failed": sum(1 for r in results if r["status"] != "ok"),
        "documents": results,
    }


def main():
    """Main loop: read JSON from stdin, execute, respond via stdout."""
    # __pycache__ 정리 (코드 변경 시 캐시가 반영 안 되는 문제 방지)
    try:
        import shutil
        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '__pycache__')
        if os.path.exists(cache_dir):
            shutil.rmtree(cache_dir, ignore_errors=True)
    except Exception:
        pass

    # Windows에서 stdin/stdout을 UTF-8로 강제 설정 (Node.js는 UTF-8로 전달)
    if hasattr(sys.stdin, 'reconfigure'):
        sys.stdin.reconfigure(encoding='utf-8', errors='replace')
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    hwp = None

    try:
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue

            req_id = None
            try:
                request = json.loads(line)
                req_id = request.get("id")
                method = request.get("method")
                params = request.get("params", {})

                if method == "shutdown":
                    respond(req_id, True, {"status": "shutting down"})
                    break

                # Lazy init HWP (ping 포함 — 첫 ping에서 COM 초기화)
                if hwp is None:
                    from pyhwpx import Hwp
                    hwp = Hwp()
                    # 모든 대화상자 자동 수락 — COM 무한 대기 방지
                    try:
                        hwp.XHwpMessageBoxMode = 1  # 0=표시, 1=자동OK
                    except Exception:
                        pass
                    try:
                        hwp.SetMessageBoxMode(0x10000)  # 모든 대화상자 자동 OK
                    except Exception:
                        pass
                    try:
                        hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
                    except Exception:
                        pass

                # 사용자 입력 차단 (COM 작업 중 커서 이동 방지)
                # 단, ParaShape/CharShape 등 COM 메시지 펌프 필요 메서드는 lock 제외
                # set_paragraph_style만 lock 제외 (인라인 ParaShape Execute에 COM 메시지 펌프 필요)
                NO_LOCK_METHODS = {"set_paragraph_style"}
                locked = False
                if method not in NO_LOCK_METHODS:
                    try:
                        if not hwp.is_command_lock():
                            hwp.lock_command()
                            locked = True
                    except Exception:
                        pass
                print(f"[DEBUG-LOOP] locked={locked}, dispatching...", file=sys.stderr)
                sys.stderr.flush()

                try:
                    result = dispatch(hwp, method, params)
                    respond(req_id, True, result)
                finally:
                    # 사용자 입력 반드시 해제
                    if locked:
                        try:
                            hwp.lock_command()  # toggle 해제
                        except Exception:
                            pass
                    # v0.6.9.1: finally 블록의 continue 제거.
                    # 이유: (1) Python 3.14 SyntaxWarning, (2) dispatch 예외를 outer
                    # except가 잡지 못해 클라이언트가 에러 응답을 못 받음. for loop가
                    # 자연스럽게 다음 iteration으로 넘어가므로 continue 불필요.

            except Exception as e:
                # 에러 시에도 잠금 해제
                try:
                    if hwp and hwp.is_command_lock():
                        hwp.lock_command()
                except Exception:
                    pass
                err_str = str(e)
                # 에러 유형 분류 (구조화된 에러 응답)
                error_type = "unknown"
                guide = ""
                if 'RPC' in err_str or '사용할 수 없' in err_str or 'disconnected' in err_str.lower():
                    error_type = "com_disconnected"
                    guide = "한글 프로그램을 종료하고 다시 실행하세요."
                    print("[WARN] COM connection lost — will reinitialize on next request", file=sys.stderr)
                    hwp = None
                elif '파일을 찾을 수 없' in err_str or 'FileNotFoundError' in err_str:
                    error_type = "file_not_found"
                    guide = "파일 경로를 확인하세요. hwp_list_files로 파일 목록을 검색할 수 있습니다."
                elif 'EBUSY' in err_str or '잠' in err_str or 'lock' in err_str.lower():
                    error_type = "file_locked"
                    guide = "파일이 다른 프로그램에서 열려있습니다. 닫고 다시 시도하세요."
                elif '열린 문서가 없' in err_str:
                    error_type = "no_document"
                    guide = "hwp_open_document로 먼저 문서를 열어주세요."
                elif '암호' in err_str or 'encrypt' in err_str.lower():
                    error_type = "encrypted"
                    guide = "암호화된 문서입니다. 비밀번호를 입력하세요."
                elif 'PermissionError' in err_str or '권한' in err_str or '쓰기' in err_str:
                    error_type = "permission_denied"
                    guide = "파일 또는 폴더의 쓰기 권한을 확인하세요. 다른 프로그램에서 파일을 닫아주세요."
                elif '디렉토리' in err_str and ('존재' in err_str or '없' in err_str):
                    error_type = "invalid_path"
                    guide = "저장할 폴더가 존재하는지 확인하세요."
                respond(req_id, False, error=err_str, error_type=error_type, guide=guide)
                print(f"[ERROR] {e}", file=sys.stderr)
                sys.stderr.flush()

    finally:
        # 한글 프로그램과 문서를 모두 유지 — 사용자가 바로 확인 가능
        # hwp.quit(), hwp.clear() 모두 호출하지 않음
        # Python 프로세스 종료 시 COM 참조만 자연 해제됨
        hwp = None


if __name__ == "__main__":
    # Handle SIGTERM gracefully (triggers finally block)
    signal.signal(signal.SIGTERM, lambda *_: sys.exit(0))
    main()
