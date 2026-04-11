"""hwp_core.table_editing.navigation — 표 진입/탈출/네비게이션.

Handlers:
- enter_table              : 단일 index 또는 nested path 진입
- exit_table               : 표 안전 탈출 (MovePos(3))
- navigate_cell            : 방향별 셀 이동 (left/right/upper/lower)
- insert_row_at_cursor     : 커서 위치 기준 행 추가 (above/below/append)
- merge_current_selection  : 이미 선택된 블록 병합
"""
from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely  # 두 점!


@register("enter_table")
def enter_table(hwp, params):
    """표 진입 (단일 index 또는 nested path)."""
    validate_params(params, ["table_index"], "enter_table")
    table_index = params["table_index"]
    select_cell = params.get("select_cell", False)
    try:
        if hwp.is_cell():
            _exit_table_safely(hwp)
    except Exception:
        pass

    try:
        if isinstance(table_index, list):
            # path-based 진입 (nested)
            if not table_index:
                return {"status": "error", "error": "table_index list 비어있음"}
            path = table_index
            hwp.get_into_nth_table(path[0])
            i = 1
            while i + 2 < len(path):
                row = int(path[i])
                col = int(path[i + 1])
                inner_idx = int(path[i + 2])
                # 현재 표 첫 셀로
                for _ in range(50):
                    try:
                        hwp.HAction.Run("TableLeftCell")
                    except Exception:
                        break
                for _ in range(50):
                    try:
                        hwp.HAction.Run("TableUpperCell")
                    except Exception:
                        break
                for _ in range(row):
                    hwp.HAction.Run("TableLowerCell")
                for _ in range(col):
                    hwp.HAction.Run("TableRightCell")
                try:
                    hwp.get_into_nth_table(inner_idx)
                except Exception as e:
                    return {"status": "error", "error": f"path 단계 {i//3 + 1} inner_idx={inner_idx} 진입 실패: {e}"}
                i += 3
            in_cell = False
            try:
                in_cell = hwp.is_cell()
            except Exception:
                pass
            if select_cell:
                try:
                    hwp.HAction.Run("TableCellBlock")
                except Exception:
                    pass
            return {"status": "ok", "table_index": path, "in_cell": in_cell, "path_depth": (len(path) - 1) // 3 + 1}
        else:
            hwp.get_into_nth_table(int(table_index))
            in_cell = False
            try:
                in_cell = hwp.is_cell()
            except Exception:
                pass
            if not in_cell:
                return {"status": "error", "error": f"표 {table_index} 진입 실패"}
            if select_cell:
                try:
                    hwp.HAction.Run("TableCellBlock")
                except Exception:
                    pass
            return {"status": "ok", "table_index": table_index, "in_cell": in_cell}
    except Exception as e:
        return {"status": "error", "error": f"enter_table 실패: {e}"}


@register("exit_table")
def exit_table(hwp, params):
    """표 안전 탈출."""
    was_in_cell = False
    try:
        was_in_cell = hwp.is_cell()
    except Exception:
        pass
    try:
        _exit_table_safely(hwp)
    except Exception as e:
        return {"status": "error", "error": str(e)}
    return {"status": "ok", "was_in_cell": was_in_cell}


@register("navigate_cell")
def navigate_cell(hwp, params):
    """표 셀 이동 (cursor 유지)."""
    import sys
    validate_params(params, ["direction"], "navigate_cell")
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
        if hasattr(hwp, action):
            try:
                result = getattr(hwp, action)()
                moved = bool(result) if result is not None else True
            except Exception as e:
                print(f"[WARN] pyhwpx {action} failed: {e}", file=sys.stderr)
                hwp.HAction.Run(action)
                moved = True
        else:
            hwp.HAction.Run(action)
            moved = True
        return {"status": "ok", "direction": direction, "moved": moved}
    except Exception as e:
        raise RuntimeError(f"셀 이동 실패 ({direction}): {e}")


@register("insert_row_at_cursor")
def insert_row_at_cursor(hwp, params):
    """커서 위치 기준 행 추가 (above/below/append)."""
    validate_params(params, ["position"], "insert_row_at_cursor")
    position = params["position"]
    if not hwp.is_cell():
        return {"status": "error", "error": "현재 커서가 표 안에 없습니다."}
    action_map = {
        "above": "TableInsertUpperRow",
        "below": "TableInsertLowerRow",
        "append": "TableAppendRow",
    }
    action = action_map.get(position)
    if action is None:
        raise ValueError(f"invalid position: {position}")
    try:
        hwp.HAction.Run(action)
        return {"status": "ok", "position": position}
    except Exception as e:
        raise RuntimeError(f"행 추가 실패 ({position}): {e}")
    finally:
        _exit_table_safely(hwp)


@register("merge_current_selection")
def merge_current_selection(hwp, params):
    """이미 선택된 블록 병합."""
    try:
        hwp.TableMergeCell()
        return {"status": "ok"}
    except Exception as e:
        raise RuntimeError(f"선택 블록 병합 실패: {e}")
    finally:
        _exit_table_safely(hwp)
