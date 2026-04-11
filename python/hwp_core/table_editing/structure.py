"""hwp_core.table_editing.structure — 표 구조 변경 (행/열/셀 추가/삭제/병합/분할).

Handlers:
- table_add_row       : 표 마지막에 행 추가
- table_delete_row    : 현재 행 삭제
- table_add_column    : 오른쪽에 열 추가
- table_delete_column : 현재 열 삭제
- table_merge_cells   : 셀 병합 (좌표 기반 또는 선택)
- table_split_cell    : 셀 분할 (rows × cols)
"""
import sys

from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely  # 두 점!


@register("table_add_row")
def table_add_row(hwp, params):
    """표 마지막에 행 추가."""
    validate_params(params, ["table_index"], "table_add_row")
    hwp.get_into_nth_table(params["table_index"])
    try:
        hwp.HAction.Run("TableAppendRow")
        return {"status": "ok", "table_index": params["table_index"]}
    except Exception:
        try:
            hwp.HAction.Run("InsertRowBelow")
            return {"status": "ok", "table_index": params["table_index"], "method": "InsertRowBelow"}
        except Exception as e:
            raise RuntimeError(f"표 행 추가 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_delete_row")
def table_delete_row(hwp, params):
    """표 현재 행 삭제."""
    validate_params(params, ["table_index"], "table_delete_row")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.TableSubtractRow()
        return {"status": "ok", "table_index": params["table_index"]}
    except Exception as e:
        raise RuntimeError(f"표 행 삭제 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_add_column")
def table_add_column(hwp, params):
    """표 오른쪽에 열 추가."""
    validate_params(params, ["table_index"], "table_add_column")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.HAction.Run("InsertColumnRight")
        return {"status": "ok", "table_index": params["table_index"]}
    except Exception as e:
        raise RuntimeError(f"표 열 추가 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_delete_column")
def table_delete_column(hwp, params):
    """표 현재 열 삭제."""
    validate_params(params, ["table_index"], "table_delete_column")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.HAction.Run("TableDeleteColumn")
        return {"status": "ok", "table_index": params["table_index"]}
    except Exception as e:
        raise RuntimeError(f"표 열 삭제 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_merge_cells")
def table_merge_cells(hwp, params):
    """표 셀 병합 (좌표 기반 또는 현재 선택)."""
    validate_params(params, ["table_index"], "table_merge_cells")
    table_index = params["table_index"]
    start_row = params.get("start_row")
    start_col = params.get("start_col")
    end_row = params.get("end_row")
    end_col = params.get("end_col")
    try:
        hwp.get_into_nth_table(table_index)
        if start_row is not None and end_row is not None and start_col is not None and end_col is not None:
            # v0.7.3 #1: TableCellBlockExtend toggle 한 번만
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
            for _ in range(start_row):
                hwp.HAction.Run("TableLowerCell")
            for _ in range(start_col):
                hwp.HAction.Run("TableRightCell")
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            for _ in range(end_col - start_col):
                hwp.HAction.Run("TableRightCell")
            for _ in range(end_row - start_row):
                hwp.HAction.Run("TableLowerCell")
            hwp.HAction.Run("TableMergeCell")
        else:
            hwp.TableMergeCell()
        return {
            "status": "ok",
            "table_index": table_index,
            "range": {"start_row": start_row, "start_col": start_col, "end_row": end_row, "end_col": end_col} if start_row is not None else None,
        }
    except Exception as e:
        raise RuntimeError(f"셀 병합 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_split_cell")
def table_split_cell(hwp, params):
    """셀 분할."""
    validate_params(params, ["table_index"], "table_split_cell")
    rows = params.get("rows")
    cols = params.get("cols")
    try:
        hwp.get_into_nth_table(params["table_index"])
        if rows is not None or cols is not None:
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
                print(f"[WARN] HTableSplitCell ParameterSet failed: {e}", file=sys.stderr)
                hwp.TableSplitCell()
        else:
            hwp.TableSplitCell()
        return {"status": "ok", "table_index": params["table_index"], "rows": rows, "cols": cols}
    except Exception as e:
        raise RuntimeError(f"셀 분할 실패: {e}")
    finally:
        _exit_table_safely(hwp)
