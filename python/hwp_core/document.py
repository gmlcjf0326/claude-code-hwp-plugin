"""HWP Core — Document management handlers.

문서 메타데이터 / 정보 조회 + 표 cell 매핑 / fill handlers.
v0.7.6.0 P1-3 두 번째~네 번째 배치 이관.

이 모듈은:
- get_document_info / get_selected_text — read-only
- document_new — _current_doc_path clear
- analyze_document, map_table_cells — hwp_analyzer 위임
- fill_document / fill_by_tab / fill_by_label — hwp_editor 위임
"""
import os
import sys
from . import register
from ._state import get_current_doc_path, set_current_doc_path
from ._helpers import validate_params, validate_file_path


@register("get_document_info")
def get_document_info(hwp, params):
    """경량 메타데이터 반환 (analyze_document 보다 빠름).

    Returns:
        status: "ok"
        pages: int — 총 페이지 수
        current_path: str — 현재 열린 문서 경로 (없으면 "")
    """
    result = {"status": "ok"}
    try:
        result["pages"] = hwp.PageCount
    except Exception:
        result["pages"] = 0
    try:
        result["current_path"] = get_current_doc_path() or ""
    except Exception as e:
        print(f"[WARN] get_document_info current_path: {e}", file=sys.stderr)
        result["current_path"] = ""
    return result


@register("get_selected_text")
def get_selected_text(hwp, params):
    """현재 선택된 텍스트 반환."""
    try:
        text = hwp.get_selected_text()
        return {"text": text}
    except Exception as e:
        return {"text": "", "error": str(e)}


@register("document_new")
def document_new(hwp, params):
    """새 빈 문서 생성. _current_doc_path 는 None 으로 clear."""
    try:
        hwp.HAction.Run("FileNew")
    except Exception as e:
        return {"status": "error", "error": f"FileNew failed: {e}"}
    # 이전 문서 잔여 상태 정리
    try:
        if hwp.is_cell():
            hwp.MovePos(3)  # 표 탈출 (Cancel() 안 됨, MovePos(3) 필수)
    except Exception:
        pass
    try:
        hwp.MovePos(2)  # movePOS_START: 본문 첫 단락
    except Exception as e:
        print(f"[WARN] document_new MovePos failed: {e}", file=sys.stderr)
    # v0.7.5.4 P0-3: 새 빈 문서는 경로 없음
    set_current_doc_path(None)
    return {"status": "ok"}


@register("analyze_document")
def analyze_document_handler(hwp, params):
    """문서 구조 분석 (페이지/표/필드/본문)."""
    from hwp_analyzer import analyze_document as _analyze
    validate_params(params, ["file_path"], "analyze_document")
    file_path = os.path.abspath(params["file_path"])
    return _analyze(hwp, file_path, already_open=(file_path == get_current_doc_path()))


@register("map_table_cells")
def map_table_cells_handler(hwp, params):
    """표 셀의 Tab 인덱스 매핑."""
    from hwp_analyzer import map_table_cells as _map
    validate_params(params, ["table_index"], "map_table_cells")
    return _map(hwp, params["table_index"])


@register("fill_document")
def fill_document_handler(hwp, params):
    """문서 일괄 채우기 (라벨 매칭 + 필드 + 표)."""
    from hwp_editor import fill_document as _fill
    return _fill(hwp, params)


@register("fill_by_tab")
def fill_by_tab_handler(hwp, params):
    """Tab 인덱스 기반 표 셀 채우기 (병합 셀 안전)."""
    from hwp_editor import fill_table_cells_by_tab as _fill
    validate_params(params, ["table_index", "cells"], "fill_by_tab")
    return _fill(hwp, params["table_index"], params["cells"])


@register("fill_by_label")
def fill_by_label_handler(hwp, params):
    """라벨 텍스트 기반 표 셀 채우기."""
    from hwp_editor import fill_table_cells_by_label as _fill
    validate_params(params, ["table_index", "cells"], "fill_by_label")
    return _fill(hwp, params["table_index"], params["cells"])


@register("smart_fill_table_auto")
def smart_fill_table_auto_handler(hwp, params):
    """v0.7.12 Phase 5E: type 기반 auto-fit 표 채우기.

    apply_auto_style (font_size/align dynamic) + fill_table_cells_by_tab + width auto.
    Hardcode 없이 cell text 기반 적응.
    """
    from hwp_editor import smart_fill_table_auto as _fill
    validate_params(params, ["table_index", "cells"], "smart_fill_table_auto")
    table_type = params.get("table_type")
    return _fill(hwp, params["table_index"], params["cells"], table_type=table_type)


# ============================================================================
# v0.7.6.0 P1-B5: document management (open/save/close) 이관 (~340 lines)
# ============================================================================


@register("open_document")
def open_document(hwp, params):
    """문서 열기 + 백업 + cursor reset + state sync.

    v0.7.4.5 Fix D: 세대별 백업
    v0.7.4.9 S2-NEW-2: cursor textbox/cell exit
    v0.7.6.0 P1-3: hwp_core._state 동기화
    """
    import os
    import sys
    import shutil
    from datetime import datetime
    validate_params(params, ["file_path"], "open_document")
    file_path = validate_file_path(params["file_path"], must_exist=True)

    # HWP 자동저장 디렉토리 확인 (.asv 방지)
    try:
        import tempfile
        asv_dir = os.path.join(tempfile.gettempdir(), "Hwp90")
        if not os.path.exists(asv_dir):
            os.makedirs(asv_dir, exist_ok=True)
    except Exception:
        pass

    # COM 상태 초기화
    try:
        hwp.MovePos(2)
    except Exception:
        pass

    # v0.7.4.9 S2-NEW-2: cursor textbox/cell exit
    try:
        hwp.XHwpMessageBoxMode = 1
    except Exception:
        pass
    for _ in range(3):
        try:
            if hwp.is_cell():
                hwp.MovePos(3)
        except Exception:
            break

    # 원본 백업 (기본 활성)
    if params.get("backup", True):
        root, ext = os.path.splitext(file_path)
        backup_path = f"{root}_backup{ext}"
        try:
            if os.path.exists(backup_path):
                ts = datetime.fromtimestamp(os.path.getmtime(backup_path)).strftime("%Y%m%d_%H%M%S")
                archive_path = f"{root}_backup_{ts}{ext}"
                if not os.path.exists(archive_path):
                    shutil.move(backup_path, archive_path)
            shutil.copy2(file_path, backup_path)
        except Exception as e:
            print(f"[WARN] open_document backup failed: {e}", file=sys.stderr)

    try:
        hwp.XHwpMessageBoxMode = 1
    except Exception as e:
        print(f"[WARN] {e}", file=sys.stderr)

    result = hwp.open(file_path)
    if not result:
        raise RuntimeError(f"한글 프로그램에서 파일을 열 수 없습니다: {file_path}")

    # state 동기화
    set_current_doc_path(file_path)

    # v0.7.2.12: MoveDocBegin
    try:
        hwp.HAction.Run("MoveDocBegin")
    except Exception as e:
        print(f"[WARN] open_document MoveDocBegin: {e}", file=sys.stderr)
    return {"status": "ok", "file_path": file_path, "pages": hwp.PageCount}


@register("save_document")
def save_document(hwp, params):
    """문서 저장 — path 지정 시 save_as 위임, 미지정 시 confirm_overwrite 필요.

    v0.7.5.4 P0-1: 원본 양식 보호 — confirm_overwrite 또는 path 필수.
    """
    import os
    import sys
    explicit_path = params.get("path")
    if explicit_path:
        save_path = validate_file_path(explicit_path, must_exist=False)
        fmt = params.get("format", "HWPX" if save_path.lower().endswith(".hwpx") else "HWP").upper()
        try:
            hwp.save_as(save_path, fmt)
            if not os.path.exists(save_path):
                return {
                    "status": "error",
                    "saved": False,
                    "error_type": "save_failed",
                    "error": f"파일이 생성되지 않았습니다: {save_path}",
                }
            set_current_doc_path(save_path)
            return {"status": "ok", "saved": True, "path": save_path, "file_size": os.path.getsize(save_path)}
        except Exception as e:
            return {
                "status": "error",
                "saved": False,
                "error_type": "save_failed",
                "error": str(e),
            }

    # path 미지정
    current = get_current_doc_path()
    if current:
        confirm_overwrite = bool(params.get("confirm_overwrite", False))
        if not confirm_overwrite:
            return {
                "status": "error",
                "saved": False,
                "error_type": "overwrite_confirmation_required",
                "error": f"원본 덮어쓰기 방지: confirm_overwrite=true 또는 path 필요. 현재 경로: {current}",
                "current_path": current,
                "hint": "새 파일로 저장하려면 path 파라미터, 원본 덮어쓰기는 confirm_overwrite=true",
            }
        try:
            hwp.save()
            return {"status": "ok", "saved": True, "path": current, "overwritten": True}
        except Exception as e:
            print(f"[WARN] save_document failed: {e}", file=sys.stderr)
            return {
                "status": "error",
                "saved": False,
                "error_type": "save_failed",
                "error": str(e),
                "path": current,
            }
    return {
        "status": "error",
        "saved": False,
        "error_type": "no_document",
        "error": "열린 문서가 없습니다",
    }


@register("save_as")
def save_as(hwp, params):
    """다른 이름으로 저장 — PDF/DOCX/HTML export 지원.

    v0.7.5.4 P0-3: 선행 hwp.save() 제거 (원본 덮어쓰기 방지).
    """
    import os
    import tempfile
    validate_params(params, ["path"], "save_as")
    save_path = validate_file_path(params["path"], must_exist=False)
    fmt = params.get("format", "HWP").upper()

    hwp.save_as(save_path, fmt)

    # 파일 실제 생성 확인 — 실패 시 임시 dir fallback
    if not os.path.exists(save_path):
        temp_path = os.path.join(tempfile.gettempdir(), os.path.basename(save_path))
        hwp.save_as(temp_path, fmt)
        if os.path.exists(temp_path):
            os.replace(temp_path, save_path)

    exists = os.path.exists(save_path)
    file_size = os.path.getsize(save_path) if exists else 0
    if not exists:
        raise RuntimeError(f"저장 실패: 파일이 생성되지 않았습니다. 경로: {save_path}")
    if file_size == 0:
        raise RuntimeError(f"저장 실패: 0바이트 빈 파일. 경로: {save_path}")

    # HWP/HWPX 일 때만 current_doc_path 갱신
    if fmt in ("HWP", "HWPX"):
        set_current_doc_path(save_path)
    return {"status": "ok", "path": save_path, "file_size": file_size}


@register("close_document")
def close_document(hwp, params):
    """문서 닫기 + state clear."""
    import sys
    try:
        hwp.XHwpMessageBoxMode = 0
    except Exception as e:
        print(f"[WARN] {e}", file=sys.stderr)
    hwp.close()
    set_current_doc_path(None)
    return {"status": "ok"}


@register("document_merge")
def document_merge(hwp, params):
    """다른 문서 병합 (문서 끝에 insert_file)."""
    import sys
    validate_params(params, ["file_path"], "document_merge")
    merge_path = validate_file_path(params["file_path"], must_exist=True)
    hwp.MovePos(3)
    try:
        hwp.HAction.Run("BreakSection")
    except Exception as e:
        print(f"[WARN] {e}", file=sys.stderr)
    hwp.insert_file(merge_path)
    return {"status": "ok", "merged_file": merge_path, "pages": hwp.PageCount}


@register("document_split")
def document_split(hwp, params):
    """문서 분할 — COM API 한계로 전체 복사 (실제 페이지 분할 아님)."""
    import os
    import shutil
    validate_params(params, ["output_dir"], "document_split")
    output_dir = os.path.abspath(params["output_dir"])
    os.makedirs(output_dir, exist_ok=True)
    total_pages = hwp.PageCount
    pages_per_split = params.get("pages_per_split", 1)
    if pages_per_split < 1:
        pages_per_split = 1
    src_path = get_current_doc_path()
    if not src_path:
        raise RuntimeError("열린 문서가 없습니다.")
    _, ext = os.path.splitext(src_path)
    parts = []
    for start in range(1, total_pages + 1, pages_per_split):
        end = min(start + pages_per_split - 1, total_pages)
        part_name = f"part_{start}-{end}{ext}"
        part_path = os.path.join(output_dir, part_name)
        shutil.copy2(src_path, part_path)
        parts.append({"pages": f"{start}-{end}", "path": part_path})
    return {
        "status": "ok",
        "total_pages": total_pages,
        "parts": len(parts),
        "files": parts,
        "warning": "COM API 한계로 각 파일은 전체 문서 복사본입니다.",
    }


@register("export_format")
def export_format(hwp, params):
    """export_format — PDF 만 공식 지원, DOCX/HTML 은 대체 제안."""
    import os
    validate_params(params, ["path", "format"], "export_format")
    save_path = validate_file_path(params["path"], must_exist=False)
    fmt = params["format"].upper()

    if fmt in ("DOCX", "DOC"):
        return {
            "status": "not_supported",
            "message": "DOCX 직접 내보내기는 한/글 COM에서 지원되지 않습니다. PDF로 내보내기 권장.",
            "alternative": "hwp_export_pdf",
        }
    if fmt == "HTML":
        return {
            "status": "not_supported",
            "message": "HTML 직접 내보내기는 미지원. hwp_get_as_markdown 사용 권장.",
            "alternative": "hwp_get_as_markdown",
        }

    # v0.7.5.4 P0-3: 원본 save() 선행 호출 제거
    # (기존: current_doc_path save 하면 원본 덮어쓰기)
    # save_as 가 직접 새 경로에 저장하므로 선행 save 불필요
    result = hwp.save_as(save_path, fmt)
    file_exists = os.path.exists(save_path)
    file_size = os.path.getsize(save_path) if file_exists else 0
    ok = file_exists and file_size > 0
    return {
        "status": "ok" if ok else "warning",
        "path": save_path,
        "format": fmt,
        "success": bool(result) and ok,
        "file_exists": file_exists,
        "file_size": file_size,
    }


@register("batch_convert")
def batch_convert(hwp, params):
    """일괄 변환 — 디렉토리 내 .hwp/.hwpx 전부 변환."""
    import os
    validate_params(params, ["input_dir", "output_format"], "batch_convert")
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
    return {
        "status": "ok",
        "total": len(results),
        "success": sum(1 for r in results if r["status"] == "ok"),
        "results": results,
    }


@register("compare_documents")
def compare_documents(hwp, params):
    """두 문서 비교 — 텍스트 diff (added/removed 라인)."""
    import os
    import sys
    validate_params(params, ["file_path_1", "file_path_2"], "compare_documents")
    path1 = validate_file_path(params["file_path_1"], must_exist=True)
    path2 = validate_file_path(params["file_path_2"], must_exist=True)
    from hwp_editor import extract_all_text

    hwp.open(path1)
    text1 = ""
    try:
        text1 = extract_all_text(hwp, max_iters=5000, strip_each=True, separator="\n")
    except Exception as e:
        print(f"[WARN] {e}", file=sys.stderr)
    hwp.close()

    hwp.open(path2)
    text2 = ""
    try:
        text2 = extract_all_text(hwp, max_iters=5000, strip_each=True, separator="\n")
    except Exception as e:
        print(f"[WARN] {e}", file=sys.stderr)
    hwp.close()

    lines1 = text1.split("\n")
    lines2 = text2.split("\n")
    added = [l for l in lines2 if l not in lines1]
    removed = [l for l in lines1 if l not in lines2]
    return {
        "status": "ok",
        "file_1": os.path.basename(path1),
        "file_2": os.path.basename(path2),
        "lines_1": len(lines1),
        "lines_2": len(lines2),
        "added": len(added),
        "removed": len(removed),
        "added_lines": added[:20],
        "removed_lines": removed[:20],
    }
