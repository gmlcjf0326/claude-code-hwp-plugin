"""hwp_core.formatting.quick_actions — 간단한 서식/셀 wrapper handlers.

Handlers:
- toggle_checkbox         : 체크박스 전환 (find/replace wrapper)
- set_background_picture  : 문서 배경 이미지
- set_cell_color          : 셀 배경색 (hwp_editor 위임)
- set_table_border        : 표 테두리 (hwp_editor 위임)
- auto_map_reference      : 참고자료 → 표 셀 자동 매핑 (hwp_editor 위임)
"""
from .. import register  # 두 점!
from .._helpers import _execute_all_replace, validate_params, validate_file_path  # 두 점!


@register("toggle_checkbox")
def toggle_checkbox(hwp, params):
    """체크박스 전환: □→■, ☐→☑ 등. 단순 find/replace wrapper."""
    validate_params(params, ["find", "replace"], "toggle_checkbox")
    find_text = params["find"]
    replace_text = params["replace"]
    replaced = _execute_all_replace(hwp, find_text, replace_text, False)
    return {
        "status": "ok",
        "find": find_text,
        "replace": replace_text,
        "replaced": replaced,
    }


@register("set_background_picture")
def set_background_picture(hwp, params):
    """문서 배경에 이미지 삽입."""
    validate_params(params, ["file_path"], "set_background_picture")
    bg_path = validate_file_path(params["file_path"], must_exist=True)
    hwp.insert_background_picture(bg_path)
    return {"status": "ok", "file_path": bg_path}


@register("set_cell_color")
def set_cell_color(hwp, params):
    """표 셀 배경색 변경 (hwp_editor 위임)."""
    validate_params(params, ["table_index", "cells"], "set_cell_color")
    from hwp_editor import set_cell_background_color
    return set_cell_background_color(hwp, params["table_index"], params["cells"])


@register("set_table_border")
def set_table_border(hwp, params):
    """표 테두리 스타일 (hwp_editor 위임)."""
    validate_params(params, ["table_index"], "set_table_border")
    from hwp_editor import set_table_border_style
    return set_table_border_style(
        hwp, params["table_index"], params.get("cells"), params.get("style", {})
    )


@register("auto_map_reference")
def auto_map_reference(hwp, params):
    """참고자료 라벨→표 셀 자동 매핑 (hwp_editor 위임)."""
    validate_params(params, ["table_index", "ref_headers", "ref_row"], "auto_map_reference")
    from hwp_editor import auto_map_reference_to_table
    return auto_map_reference_to_table(
        hwp, params["table_index"], params["ref_headers"], params["ref_row"]
    )
