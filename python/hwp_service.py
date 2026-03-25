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


def validate_file_path(file_path, must_exist=True):
    """경로 보안 검증. 심볼릭 링크 거부, 존재 여부 확인."""
    real = os.path.abspath(file_path)
    if must_exist and not os.path.exists(real):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {real}")
    if os.path.islink(file_path):
        raise ValueError(f"심볼릭 링크는 허용되지 않습니다: {file_path}")
    return real


def _execute_all_replace(hwp, find_str, replace_str, use_regex=False):
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


def respond(req_id, success, data=None, error=None):
    """Send JSON response to stdout."""
    response = {"id": req_id, "success": success}
    if data is not None:
        response["data"] = data
    if error is not None:
        response["error"] = error
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

    if method == "open_document":
        validate_params(params, ["file_path"], method)
        file_path = validate_file_path(params["file_path"], must_exist=True)

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

    if method == "get_cursor_context":
        return {"context": "cursor context placeholder"}

    if method == "save_as":
        validate_params(params, ["path"], method)
        save_path = validate_file_path(params["path"], must_exist=False)
        fmt = params.get("format", "HWP").upper()  # pyhwpx는 대문자 포맷 필요 (HWP, HWPX, PDF 등)
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
        hwp.MovePos(2)  # 문서 시작
        results = []
        for i in range(max_results):
            act = hwp.HAction
            pset = hwp.HParameterSet.HFindReplace
            act.GetDefault("FindReplace", pset.HSet)
            pset.FindString = search_text
            pset.Direction = 0
            pset.IgnoreMessage = 1
            act.Execute("FindReplace", pset.HSet)
            # BUG-3 fix: 반환값 대신 선택 영역 존재 여부로 판단
            context = ""
            try:
                context = hwp.GetTextFile("TEXT", "saveblock").strip()[:200]
            except Exception:
                pass
            if not context:
                break  # 선택 영역이 없으면 더 이상 찾을 수 없음
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
        replaced = _execute_all_replace(hwp, params["find"], params["replace"], use_regex)
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
        act = hwp.HAction
        pset = hwp.HParameterSet.HFindReplace
        act.GetDefault("FindReplace", pset.HSet)
        pset.FindString = params["find"]
        pset.Direction = 0
        pset.IgnoreMessage = 1
        act.Execute("FindReplace", pset.HSet)

        # BUG-4 fix: 반환값 대신 선택 영역으로 찾기 성공 판단
        try:
            selected = hwp.GetTextFile("TEXT", "saveblock").strip()
        except Exception:
            selected = ""
        if not selected:
            return {"status": "not_found", "find": params["find"]}

        # 2C2 fix: 찾은 텍스트 끝으로 커서 이동
        # FindReplace가 텍스트를 선택한 상태에서 MoveRight → 선택 해제 + 선택 끝으로 이동
        hwp.HAction.Run("MoveRight")

        # 색상 설정 (옵션)
        color = params.get("color")  # [r, g, b]
        if color:
            from hwp_editor import insert_text_with_color
            insert_text_with_color(hwp, params["append_text"], tuple(color))
        else:
            hwp.insert_text(params["append_text"])

        return {"status": "ok", "find": params["find"], "appended": True}

    if method == "insert_text":
        validate_params(params, ["text"], method)
        style = params.get("style")
        color = params.get("color")  # [r, g, b] 하위 호환
        if style:
            from hwp_editor import insert_text_with_style
            insert_text_with_style(hwp, params["text"], style)
        elif color:
            from hwp_editor import insert_text_with_color
            insert_text_with_color(hwp, params["text"], tuple(color))
        else:
            hwp.insert_text(params["text"])
        return {"status": "ok"}

    if method == "set_paragraph_style":
        validate_params(params, ["style"], method)
        set_paragraph_style(hwp, params["style"])
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
        nth = params["nth"]  # 1-based
        if nth < 1 or nth > 10000:
            raise ValueError("nth must be between 1 and 10000")
        hwp.MovePos(2)  # 문서 시작
        for i in range(nth):
            act = hwp.HAction
            pset = hwp.HParameterSet.HFindReplace
            act.GetDefault("FindReplace", pset.HSet)
            pset.FindString = params["find"]
            pset.Direction = 0
            pset.IgnoreMessage = 1
            found = act.Execute("FindReplace", pset.HSet)
            if not found:
                return {"status": "not_found", "find": params["find"], "searched": i, "nth": nth}
        # N번째 매칭이 선택된 상태 → 텍스트 교체
        hwp.insert_text(params["replace"])
        return {"status": "ok", "find": params["find"], "replace": params["replace"], "nth": nth}

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
            try:
                hwp.Cancel()
            except Exception:
                pass

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
            try:
                hwp.Cancel()
            except Exception:
                pass

    if method == "table_add_column":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("InsertColumnRight")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 열 추가 실패: {e}")
        finally:
            try:
                hwp.Cancel()
            except Exception:
                pass

    if method == "table_delete_column":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("DeleteColumn")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 열 삭제 실패: {e}")
        finally:
            try:
                hwp.Cancel()
            except Exception:
                pass

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
                # H2: 범위 지정 병합 — TableCellBlock으로 셀 선택 후 병합
                hwp.HAction.Run("TableColBegin")
                hwp.HAction.Run("TableRowBegin")
                for _ in range(start_row):
                    hwp.HAction.Run("TableLowerCell")
                for _ in range(start_col):
                    hwp.HAction.Run("TableRightCell")
                hwp.HAction.Run("TableCellBlock")
                for _ in range(end_col - start_col):
                    hwp.HAction.Run("TableRightCell")
                for _ in range(end_row - start_row):
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
            try:
                hwp.Cancel()
            except Exception:
                pass

    if method == "table_split_cell":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.TableSplitCell()
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"셀 분할 실패: {e}")
        finally:
            try:
                hwp.Cancel()
            except Exception:
                pass

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
        # 셀 채우기
        filled = 0
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                if val:
                    # BUG-1 fix: SelectAll 제거 — 새 표의 빈 셀에 직접 삽입
                    if header_style and r == 0:
                        from hwp_editor import insert_text_with_style
                        insert_text_with_style(hwp, str(val), {"bold": True})
                    else:
                        hwp.insert_text(str(val))
                    filled += 1
                if c < len(row) - 1 or r < rows - 1:
                    hwp.TableRightCell()
        # 표 밖으로 커서 이동 (표 생성 후 커서가 표 안에 남아있음)
        try:
            hwp.Cancel()  # 셀 선택 해제
            # 표 아래로 커서 이동: Ctrl+End 방향으로 표 탈출
            hwp.HAction.Run("MoveDocEnd")  # 문서 끝으로 이동
            hwp.HAction.Run("BreakPara")   # 새 문단 생성 (표 아래)
        except Exception as e:
            print(f"[WARN] Table exit: {e}", file=sys.stderr)
        # 헤더행 배경색 적용 (옵션)
        if header_style and rows > 0:
            try:
                from hwp_editor import set_cell_background_color
                header_cells = [{"tab": i, "color": "#E8E8E8"} for i in range(cols)]
                set_cell_background_color(hwp, -1, header_cells)  # -1 = 현재 표
            except Exception:
                pass  # 배경색 실패해도 표 자체는 유지
        return {"status": "ok", "rows": rows, "cols": cols, "filled": filled, "header_styled": bool(header_style)}

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
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        return {"status": "ok", "file": os.path.basename(csv_path), "rows": rows, "cols": cols, "filled": filled}

    if method == "insert_heading":
        validate_params(params, ["text", "level"], method)
        from hwp_editor import insert_text_with_style
        level = min(max(params["level"], 1), 6)
        sizes = {1: 22, 2: 18, 3: 15, 4: 13, 5: 11, 6: 10}
        text = params["text"]
        # 순번 자동 생성
        numbering = params.get("numbering")
        number = params.get("number", 1)
        if numbering:
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
        return {"status": "ok", "level": level, "text": text}

    if method == "export_format":
        validate_params(params, ["path", "format"], method)
        save_path = validate_file_path(params["path"], must_exist=False)
        fmt = params["format"].upper()  # HWP, HWPX, PDF, HTML, TXT 등
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
            # 분할 저장은 COM API 한계로 전체 복사 후 저장
            shutil.copy2(src_path, part_path)
            parts.append({"pages": f"{start}-{end}", "path": part_path})
        return {"status": "ok", "total_pages": total_pages, "parts": len(parts), "files": parts}

    if method == "insert_footnote":
        hwp.HAction.Run("InsertFootnote")
        text = params.get("text")
        if text:
            hwp.insert_text(text)
        return {"status": "ok", "type": "footnote"}

    if method == "insert_endnote":
        hwp.HAction.Run("InsertEndnote")
        text = params.get("text")
        if text:
            hwp.insert_text(text)
        return {"status": "ok", "type": "endnote"}

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
        import re
        hwp.InitScan(0x0077)
        texts = []
        count = 0
        while count < 1000:
            state, t = hwp.GetText()
            if state <= 0:
                break
            if t and t.strip():
                texts.append(t.strip())
            count += 1
        hwp.ReleaseScan()
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
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        # 헤더행 + ■ 셀 배경색 적용
        try:
            from hwp_editor import set_cell_background_color
            style_cells = [{"tab": i, "color": "#D9D9D9"} for i in range(cols)]  # 헤더: 연회색
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
            try:
                hwp.Cancel()
            except Exception:
                pass

    if method == "table_formula_avg":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("TableFormulaAvgAuto")
            return {"status": "ok", "table_index": params["table_index"], "formula": "avg"}
        except Exception as e:
            raise RuntimeError(f"표 평균 계산 실패: {e}")
        finally:
            try:
                hwp.Cancel()
            except Exception:
                pass

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
            try:
                hwp.Cancel()
            except Exception:
                pass
        return {"status": "ok", "table_index": params["table_index"], "path": output_path}

    if method == "break_section":
        hwp.BreakSection()
        return {"status": "ok", "type": "section"}

    if method == "break_column":
        hwp.BreakColumn()
        return {"status": "ok", "type": "column"}

    if method == "insert_line":
        hwp.HAction.Run("InsertLine")
        return {"status": "ok"}

    if method == "table_swap_type":
        validate_params(params, ["table_index"], method)
        try:
            hwp.get_into_nth_table(params["table_index"])
            hwp.HAction.Run("TableSwapType")
            return {"status": "ok", "table_index": params["table_index"]}
        except Exception as e:
            raise RuntimeError(f"표 행/열 교환 실패: {e}")
        finally:
            try:
                hwp.Cancel()
            except Exception:
                pass

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
            try:
                hwp.Cancel()
            except Exception:
                pass

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
        # 문서 1 텍스트 추출
        hwp.open(path1)
        text1 = ""
        try:
            hwp.InitScan(0x0077)
            parts = []
            count = 0
            while count < 5000:
                state, t = hwp.GetText()
                if state <= 0:
                    break
                if t and t.strip():
                    parts.append(t.strip())
                count += 1
            hwp.ReleaseScan()
            text1 = "\n".join(parts)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        hwp.close()
        # 문서 2 텍스트 추출
        hwp.open(path2)
        text2 = ""
        try:
            hwp.InitScan(0x0077)
            parts = []
            count = 0
            while count < 5000:
                state, t = hwp.GetText()
                if state <= 0:
                    break
                if t and t.strip():
                    parts.append(t.strip())
                count += 1
            hwp.ReleaseScan()
            text2 = "\n".join(parts)
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
        text = ""
        try:
            hwp.InitScan(0x0077)
            parts = []
            count = 0
            while count < 10000:
                state, t = hwp.GetText()
                if state <= 0:
                    break
                if t:
                    parts.append(t)
                count += 1
            hwp.ReleaseScan()
            text = "".join(parts)
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
            hwp.Cancel()
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
        import re
        text = ""
        try:
            hwp.InitScan(0x0077)
            parts = []
            count = 0
            while count < 10000:
                state, t = hwp.GetText()
                if state <= 0:
                    break
                if t:
                    parts.append(t)
                count += 1
            hwp.ReleaseScan()
            text = "\n".join(parts)
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
                    # 메시지박스(얼럿/다이얼로그) 자동 확인 — COM 무한 대기 방지
                    try:
                        hwp.XHwpMessageBoxMode = 1  # 0=표시, 1=자동OK
                    except Exception:
                        pass

                result = dispatch(hwp, method, params)
                respond(req_id, True, result)

            except Exception as e:
                err_str = str(e)
                # RPC/COM 연결 끊김 시 다음 요청에서 자동 재초기화
                if 'RPC' in err_str or '사용할 수 없' in err_str or 'disconnected' in err_str.lower():
                    print("[WARN] COM connection lost — will reinitialize on next request", file=sys.stderr)
                    hwp = None
                respond(req_id, False, error=err_str)
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
