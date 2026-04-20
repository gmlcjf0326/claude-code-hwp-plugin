"""hwp_core.table_editing.creation — 표 생성 (데이터 기반 / 결재란 / CSV 로드).

Handlers:
- table_create_from_data : data 2D 배열로 표 생성 + col_widths 자동 축소 + header_style
                           v0.7.3.3 FB-1: cell-aware 자동 탈출 (nested 표 지원)
- create_approval_box    : 결재란 4×N 표 자동 생성
- table_insert_from_csv  : CSV/Excel 파일을 표로 삽입
"""
import os
import sys

from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely, validate_file_path  # 두 점!


@register("table_create_from_data")
def table_create_from_data(hwp, params):
    """표 생성 + col_widths 자동 축소 + row_heights 자동 축소 + alignment + header_style.

    v0.7.3.1: nested 표 지원 (cell 안 호출 시 부모 cell 너비 기준).
    v0.7.3.3 #FB-1: cell-aware 자동 탈출.
    """
    validate_params(params, ["data"], "table_create_from_data")
    data = params["data"]
    if not data or not isinstance(data, list):
        raise ValueError("data must be a non-empty 2D array")
    rows = len(data)
    cols = max(len(row) for row in data) if data else 0
    header_style = params.get("header_style", False)
    col_widths = params.get("col_widths")
    row_heights = params.get("row_heights")
    alignment = params.get("alignment")
    treat_as_char = params.get("treat_as_char")

    col_width_warning = None
    in_cell_for_table = False
    try:
        in_cell_for_table = hwp.is_cell()
    except Exception:
        pass

    if in_cell_for_table:
        cell_width_mm = 50
        try:
            cs = hwp.CellShape
            cell_width_hwpu = None
            try:
                cell_width_hwpu = cs.Item("Width")
            except Exception:
                cell_width_hwpu = getattr(cs, "Width", None)
            if cell_width_hwpu:
                cell_width_mm = cell_width_hwpu / 283.465
        except Exception:
            try:
                pset_cs = hwp.HParameterSet.HCellShape
                hwp.HAction.GetDefault("CellShape", pset_cs.HSet)
                cell_width_mm = pset_cs.Width / 283.465
            except Exception as e:
                print(f"[INFO] cell width 측정 실패: {e}", file=sys.stderr)
        target_width = max(cell_width_mm - 2, 10)
        page_d = {}
    else:
        try:
            page_d = hwp.get_pagedef_as_dict()
            usable_width = page_d.get("용지폭", 210) - page_d.get("왼쪽", 30) - page_d.get("오른쪽", 30)
        except Exception:
            usable_width = 160
            page_d = {}
        usable_width = max(usable_width, 50)
        target_width = max(usable_width - 5, 20)

    if col_widths:
        total_width = sum(col_widths)
        if abs(total_width - target_width) > 1:
            ratio = target_width / total_width
            col_widths = [round(w * ratio, 1) for w in col_widths]
            if total_width > target_width + 5:
                ctx = "cell 폭" if in_cell_for_table else "페이지 폭"
                col_width_warning = f"col_widths 합계({total_width}mm)를 {ctx}({target_width}mm)에 맞춰 조정했습니다."
    else:
        if cols > 0:
            col_widths = [round(target_width / cols, 1)] * cols
        else:
            col_widths = []

    row_height_warning = None
    if not in_cell_for_table and row_heights:
        try:
            usable_height = page_d.get("용지길이", 297) - page_d.get("위쪽", 20) - page_d.get("아래쪽", 15)
            target_height = max(usable_height - 5, 30)
            total_height = sum(row_heights)
            if total_height > target_height + 1:
                ratio = target_height / total_height
                row_heights = [round(h * ratio, 1) for h in row_heights]
                row_height_warning = f"row_heights 합계({total_height}mm)를 페이지 높이({target_height}mm)에 맞춰 조정했습니다."
        except Exception as e:
            print(f"[WARN] row_heights ratio: {e}", file=sys.stderr)

    # HTableCreation 으로 정밀 생성
    if col_widths or row_heights:
        try:
            tc = hwp.HParameterSet.HTableCreation
            hwp.HAction.GetDefault("TableCreate", tc.HSet)
            tc.Rows = rows
            tc.Cols = cols
            tc.WidthType = 2
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
            print(f"[WARN] HTableCreation failed: {e}", file=sys.stderr)
            hwp.create_table(rows, cols)
    else:
        hwp.create_table(rows, cols)

    # 셀 채우기
    align_map = {"left": 0, "center": 1, "right": 2}
    wide_table_font_size = 9 if cols >= 6 else None
    # 헤더 기본 서식 (v0.7.6+ defaults 개선):
    #   배경색 #E8E8E8 (밝은 회색, brightness 232 — 검정 글자 가독성 우수)
    #   글자색 #333333 (진한 회색 — 검정보다 부드럽고 인쇄 대비 양호)
    #   정렬: 헤더 행은 항상 가운데 강제 (본문 행은 alignment 파라미터 따름)
    header_bg_color = params.get("header_bg_color") or [232, 232, 232]
    header_text_color = params.get("header_text_color") or [51, 51, 51]
    filled = 0
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            # 헤더 행은 가운데 강제, 본문 행은 지정된 alignment
            row_alignment = "center" if (header_style and r == 0) else alignment
            if row_alignment and row_alignment in align_map:
                try:
                    act_p = hwp.HAction
                    ps = hwp.HParameterSet.HParaShape
                    act_p.GetDefault("ParagraphShape", ps.HSet)
                    ps.AlignType = align_map[row_alignment]
                    act_p.Execute("ParagraphShape", ps.HSet)
                except Exception as e:
                    print(f"[WARN] Cell align: {e}", file=sys.stderr)
            if val:
                if header_style and r == 0:
                    from hwp_editor import insert_text_with_style
                    style = {"bold": True, "color": list(header_text_color)}
                    if wide_table_font_size:
                        style["font_size"] = wide_table_font_size
                    insert_text_with_style(hwp, str(val), style)
                elif wide_table_font_size and r > 0:
                    from hwp_editor import insert_text_with_style
                    insert_text_with_style(hwp, str(val), {"font_size": wide_table_font_size})
                else:
                    hwp.insert_text(str(val))
                filled += 1
            # 헤더 셀이면 배경색 채우기 (텍스트 삽입 후, 다음 셀 이동 전)
            if header_style and r == 0:
                try:
                    bg = header_bg_color
                    hwp.cell_fill((int(bg[0]), int(bg[1]), int(bg[2])))
                except Exception as e:
                    print(f"[WARN] header cell_fill: {e}", file=sys.stderr)
            if c < len(row) - 1 or r < rows - 1:
                hwp.TableRightCell()

    # cell-aware 자동 탈출
    if not in_cell_for_table:
        _exit_table_safely(hwp)

    # treat_as_char 옵션
    treat_as_char_result = None
    if treat_as_char is not None:
        try:
            action_tried = False
            for action_name in ("TableTreatAsChar", "ShapeObjTreatAsChar", "TreatAsChar"):
                try:
                    hwp.HAction.Run(action_name)
                    treat_as_char_result = {"action": action_name, "value": treat_as_char}
                    action_tried = True
                    break
                except Exception:
                    continue
            if not action_tried:
                treat_as_char_result = {"warning": "TreatAsChar 액션 미지원"}
        except Exception as e:
            treat_as_char_result = {"error": str(e)}

    result = {"status": "ok", "rows": rows, "cols": cols, "filled": filled, "header_styled": bool(header_style)}
    if col_width_warning:
        result["col_width_warning"] = col_width_warning
    if row_height_warning:
        result["row_height_warning"] = row_height_warning
    if in_cell_for_table:
        result["in_cell_nested"] = True
    if treat_as_char_result is not None:
        result["treat_as_char"] = treat_as_char_result
    return result


@register("create_approval_box")
def create_approval_box(hwp, params):
    """결재란 4×N 표 자동 생성 (기안/검토/결재 levels)."""
    levels = params.get("levels", ["기안", "검토", "결재"])
    position = params.get("position", "right")
    cols = len(levels) + 1
    rows = 4
    data = [["구분"] + levels]
    data.append(["직급"] + ["" for _ in levels])
    data.append(["성명"] + ["" for _ in levels])
    data.append(["서명"] + ["" for _ in levels])
    col_widths = [18] + [25 for _ in levels]
    row_heights = [8, 8, 12, 12]

    # REGISTRY lookup 으로 table_create_from_data 호출 (순환 의존 회피)
    from hwp_core import REGISTRY
    table_handler = REGISTRY.get("table_create_from_data")
    if table_handler:
        table_handler(hwp, {
            "data": data,
            "col_widths": col_widths,
            "row_heights": row_heights,
            "alignment": position,
            "header_style": True,
        })

    try:
        from hwp_editor import set_cell_background_color
        cells = [{"tab": i, "color": "#E8E8E8"} for i in range(cols)]
        set_cell_background_color(hwp, 0, cells)
    except Exception as e:
        print(f"[WARN] Approval box style: {e}", file=sys.stderr)
    return {"status": "ok", "rows": rows, "cols": cols, "levels": levels}


@register("table_insert_from_csv")
def table_insert_from_csv(hwp, params):
    """CSV/Excel 파일을 표로 삽입."""
    validate_params(params, ["file_path"], "table_insert_from_csv")
    csv_path = validate_file_path(params["file_path"], must_exist=True)
    from ref_reader import read_reference
    ref = read_reference(csv_path)
    if ref.get("format") not in ("csv", "excel"):
        raise ValueError(f"CSV 또는 Excel 파일만 지원합니다. (현재: {ref.get('format')})")
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
                hwp.insert_text(str(val))
                filled += 1
            if c < len(row) - 1 or r < rows - 1:
                hwp.TableRightCell()
    _exit_table_safely(hwp)
    return {"status": "ok", "file": os.path.basename(csv_path), "rows": rows, "cols": cols, "filled": filled}
