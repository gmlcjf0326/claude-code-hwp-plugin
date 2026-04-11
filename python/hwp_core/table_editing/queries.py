"""hwp_core.table_editing.queries — 표 조회 + 읽기 전용 wrapper.

Handlers:
- get_table_dimensions      : 표 치수 (너비, 여백, 셀 수)
- get_cell_format           : 특정 셀 서식 (hwp_editor 위임)
- get_table_format_summary  : 표 전체 서식 요약 (hwp_editor 위임)
- smart_fill                : 서식 감지 + 채우기 (hwp_editor 위임)
- read_reference            : 참고자료 파일 읽기 (ref_reader 위임)
"""
from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely  # 두 점!


@register("get_table_dimensions")
def get_table_dimensions(hwp, params):
    """표 치수 추출 — 전체 너비, 셀 여백, 행/열 구조."""
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
    try:
        from hwp_analyzer import map_table_cells as _map
        cell_data = _map(hwp, table_index)
        result["total_cells"] = cell_data.get("total_cells", 0)
    except Exception:
        pass
    _exit_table_safely(hwp)
    return result


@register("get_cell_format")
def get_cell_format(hwp, params):
    """표 셀 서식 조회."""
    validate_params(params, ["table_index", "cell_tab"], "get_cell_format")
    from hwp_editor import get_cell_format as _get
    return _get(hwp, params["table_index"], params["cell_tab"])


@register("get_table_format_summary")
def get_table_format_summary(hwp, params):
    """표 전체 서식 요약."""
    validate_params(params, ["table_index"], "get_table_format_summary")
    from hwp_editor import get_table_format_summary as _get
    return _get(hwp, params["table_index"], params.get("sample_tabs"))


@register("smart_fill")
def smart_fill(hwp, params):
    """서식 감지 후 자동 적용 표 채우기."""
    validate_params(params, ["table_index", "cells"], "smart_fill")
    from hwp_editor import smart_fill_table_cells as _fill
    return _fill(hwp, params["table_index"], params["cells"])


@register("read_reference")
def read_reference(hwp, params):
    """참고 자료 파일 읽기 (Excel/CSV/PDF/HWP)."""
    validate_params(params, ["file_path"], "read_reference")
    from ref_reader import read_reference as _read
    return _read(
        params["file_path"],
        params.get("max_chars", 30000),
        hwp=hwp,
    )
