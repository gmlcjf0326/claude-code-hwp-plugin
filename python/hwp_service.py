"""HWP Studio AI - Python HWP Service Bridge
stdin/stdout JSON protocol for Electron <-> Python communication.

v0.7.6.0 모듈화: helper 함수는 hwp_core/_helpers.py 로 이동.
  - _exit_table_safely, validate_file_path, _execute_all_replace, validate_params
  - _current_doc_path 는 hwp_core/_state.py 로 이동
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

# v0.7.6.0 P1-1: hwp_core 모듈화 import
from hwp_core import REGISTRY
from hwp_core._helpers import (
    _exit_table_safely,
    validate_file_path,
    _execute_all_replace,
    validate_params,
)
# v0.7.6.0 P1-3: state sync — hwp_service 의 _current_doc_path 와 hwp_core._state 동기화
from hwp_core._state import set_current_doc_path as _sync_doc_path


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


_current_doc_path = None

def dispatch(hwp, method, params):
    """Route method calls to appropriate handlers.

    v0.7.6.0 P1-2: REGISTRY-first lookup + if-elif fallback.
    1. hwp_core/REGISTRY 에서 method 찾으면 handler 호출 (O(1))
    2. miss 시 기존 if-elif 사슬 fallback (이관 중인 메서드 지원)
    """
    global _current_doc_path

    # v0.7.6.0 P1-2: REGISTRY-first lookup
    handler = REGISTRY.get(method)
    if handler is not None:
        try:
            return handler(hwp, params)
        except Exception as e:
            # handler 에러는 상위 main 에서 catch 되므로 재전달
            raise

    # v0.7.6.0 P1-3: ping 은 hwp_core/utility.py 로 이관됨 (REGISTRY)

    # v0.7.5.4 P4-1: 5단계 검증 helper (TEST_CHECKLIST.md Phase 19 표준)
    # v0.7.6.0 P1-B4: verify_5stage 는 hwp_core/analysis.py 로 이관됨

    # v0.7.6.0 P1-B4: snapshot_template_style 는 hwp_core/analysis.py 로 이관됨

    # v0.7.6.0 P1-B4: detect_document_type 는 hwp_core/analysis.py 로 이관됨

    if method == "inspect_com_object":
        obj_name = params.get("object", "HCharShape")
        if obj_name == "HCharShape":
            pset = hwp.HParameterSet.HCharShape
            hwp.HAction.GetDefault("CharShape", pset.HSet)
        elif obj_name == "HParaShape":
            pset = hwp.HParameterSet.HParaShape
            hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
        elif obj_name == "HFindReplace":
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("AllReplace", pset.HSet)
        elif obj_name == "HSecDef":
            pset = hwp.HParameterSet.HSecDef
            hwp.HAction.GetDefault("PageSetup", pset.HSet)
        elif obj_name == "HPageDef":
            hsec = hwp.HParameterSet.HSecDef
            hwp.HAction.GetDefault("PageSetup", hsec.HSet)
            pset = hsec.PageDef
        else:
            return {"error": f"Unknown object: {obj_name}"}
        attrs = [a for a in dir(pset) if not a.startswith('_')]
        # 값도 같이 dump
        values = {}
        for a in attrs:
            try:
                v = getattr(pset, a)
                if isinstance(v, (int, float, str, bool)):
                    values[a] = v
            except Exception:
                pass
        return {"object": obj_name, "attributes": attrs, "count": len(attrs), "values": values}

    # v0.7.2.5: 빈 문서 생성 (autopilot blank 분기에서 호출)
    # v0.7.2.9: cursor 초기화 + 본문 시작점 보장 (이전 문서가 표 셀 cursor였을 수 있음)
    # v0.7.6.0 P1-3: document_new 는 hwp_core/document.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: open_document 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: save_document 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: save_as 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: close_document 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: document_merge 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: export_format 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: insert_textbox 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: draw_line 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: image_extract 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: document_split 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: insert_page_num 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: generate_toc 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: create_gantt_chart 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: insert_auto_num 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: insert_memo 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: batch_convert 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: compare_documents 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: delete_guide_text 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: verify_after_fill 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: generate_multi_documents 는 hwp_core/*.py 로 이관됨

    # v0.7.6.0 P1-B5/B6/B7: clone_pdf_to_hwp 는 hwp_core/*.py 로 이관됨

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
                # v0.7.4.5 Fix C: 장시간 실행 메서드는 lock 제외 (다른 RPC 호출 stall 방지)
                NO_LOCK_METHODS = {
                    "set_paragraph_style",        # 인라인 ParaShape Execute COM 메시지 펌프 필요
                    "clone_pdf_to_hwp",           # v0.7.4.5: PDF OCR 최대 10분 장시간 실행
                    "analyze_writing_patterns",   # v0.7.4.5: .hwpx XML 직접 파싱 시간
                    "extract_template_structure", # v0.7.4.5: 양식 구조 분석 시간
                    "estimate_workload",          # v0.7.4.5: 파일 분석 시간
                }
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
