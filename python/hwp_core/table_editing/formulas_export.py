"""hwp_core.table_editing.formulas_export — 표 수식 + 구조 변환 + export.

Handlers:
- table_formula_sum       : 표 합계 자동 계산
- table_formula_avg       : 표 평균 자동 계산
- table_to_csv            : 표를 CSV 파일로 내보내기 (병합 셀 fallback)
- table_to_json           : 표를 JSON 으로 반환
- table_swap_type         : 표 행/열 교환
- table_distribute_width  : 셀 너비 균등 분배
"""
from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely, validate_file_path  # 두 점!


@register("table_formula_sum")
def table_formula_sum(hwp, params):
    """표 합계 자동 계산."""
    validate_params(params, ["table_index"], "table_formula_sum")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.HAction.Run("TableFormulaSumAuto")
        return {"status": "ok", "table_index": params["table_index"], "formula": "sum"}
    except Exception as e:
        raise RuntimeError(f"표 합계 계산 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_formula_avg")
def table_formula_avg(hwp, params):
    """표 평균 자동 계산."""
    validate_params(params, ["table_index"], "table_formula_avg")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.HAction.Run("TableFormulaAvgAuto")
        return {"status": "ok", "table_index": params["table_index"], "formula": "avg"}
    except Exception as e:
        raise RuntimeError(f"표 평균 계산 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_to_csv")
def table_to_csv(hwp, params):
    """표를 CSV 파일로 내보내기."""
    validate_params(params, ["table_index", "output_path"], "table_to_csv")
    output_path = validate_file_path(params["output_path"], must_exist=False)
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.table_to_csv(output_path)
    except Exception:
        # 병합 셀 등으로 실패 시 map_table_cells 대안
        import csv
        from hwp_analyzer import map_table_cells
        cell_data = map_table_cells(hwp, params["table_index"])
        cells = cell_data.get("cell_map", [])
        with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            for c in cells:
                writer.writerow([c.get("tab", ""), c.get("text", "")])
    finally:
        _exit_table_safely(hwp)
    return {"status": "ok", "table_index": params["table_index"], "path": output_path}


@register("table_to_json")
def table_to_json(hwp, params):
    """표를 JSON 으로 반환."""
    validate_params(params, ["table_index"], "table_to_json")
    from hwp_analyzer import map_table_cells
    cell_data = map_table_cells(hwp, params["table_index"])
    cell_map = cell_data.get("cell_map", [])
    json_data = [{"tab": c["tab"], "text": c["text"]} for c in cell_map]
    return {
        "status": "ok",
        "table_index": params["table_index"],
        "total_cells": len(json_data),
        "cells": json_data,
    }


@register("table_swap_type")
def table_swap_type(hwp, params):
    """표 행/열 교환."""
    validate_params(params, ["table_index"], "table_swap_type")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.HAction.Run("TableSwapType")
        return {"status": "ok", "table_index": params["table_index"]}
    except Exception as e:
        raise RuntimeError(f"표 행/열 교환 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("table_distribute_width")
def table_distribute_width(hwp, params):
    """셀 너비 균등 분배."""
    validate_params(params, ["table_index"], "table_distribute_width")
    try:
        hwp.get_into_nth_table(params["table_index"])
        hwp.HAction.Run("TableDistributeCellWidth")
        return {"status": "ok", "table_index": params["table_index"]}
    except Exception as e:
        raise RuntimeError(f"셀 너비 균등 분배 실패: {e}")
    finally:
        _exit_table_safely(hwp)
