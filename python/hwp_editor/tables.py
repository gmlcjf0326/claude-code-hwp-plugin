"""hwp_editor.tables — 표 네비게이션 + 채우기 + 서식.

함수:
- _goto_cell                   : 표 내 셀 이동 (sequential Tab)
- _navigate_to_tab             : target tab 까지 안전 이동 (역방향 = reset)
- _hex_to_rgb                  : #RRGGBB → (r,g,b) 헬퍼
- fill_table_cells_by_tab      : tab 인덱스 기반 셀 채우기 (style 지원)
- smart_fill_table_cells       : 서식 감지 + 재적용 (v0.7.4.8 Fix D5)
- fill_table_cells_by_label    : 라벨 기반 → resolve_labels_to_tabs → by_tab
- verify_after_fill            : 채우기 후 실제 값 검증
- set_cell_background_color    : 셀 배경색 (어두우면 auto 흰색+bold)
- set_table_border_style       : 셀/표 테두리 (4면 각각 type/width/color)

v0.7.5.0 Issue 3: 역방향 이동 → MovePos(2) reset + 정방향만
v0.7.4.8 Fix D5: format reapply (read-only 아님)
"""
import sys

from .text_style import insert_text_with_style
from .char_para import get_char_shape, get_para_shape


def _goto_cell(hwp, table_idx, cell_positions, target_cell_idx):
    """Navigate to a specific cell by its sequential index using Tab."""
    hwp.get_into_nth_table(table_idx)

    for _ in range(target_cell_idx):
        try:
            hwp.TableRightCell()
        except Exception:
            break


def _navigate_to_tab(hwp, table_idx, target_tab, current_tab):
    """셀 네비게이션 공통 로직. 새 current_tab을 반환.
    v0.7.5.0 Issue 3: 역방향 이동 시 Cancel() 대신 MovePos(2) (문서 처음) 로 완전 reset
    후 정방향만 사용. 병합 셀 있는 표에서도 안전.
    """
    if target_tab < current_tab:
        # 역방향 → 문서 처음으로 복귀 후 표 재진입 + 정방향만
        try:
            hwp.MovePos(2)  # MoveDocBegin
        except Exception as e:
            print(f"[WARN] MovePos reset: {e}", file=sys.stderr)
        try:
            hwp.get_into_nth_table(table_idx)
        except Exception as e:
            print(f"[WARN] get_into_nth_table reenter: {e}", file=sys.stderr)
        moves = target_tab
    else:
        moves = target_tab - current_tab
    for _ in range(moves):
        try:
            hwp.TableRightCell()
        except Exception as e:
            print(f"[WARN] TableRightCell: {e}", file=sys.stderr)
            break
    return target_tab


def _hex_to_rgb(hex_color):
    """#RRGGBB 헥스 색상을 (r, g, b) 튜플로 변환."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def fill_table_cells_by_tab(hwp, table_idx, cells):
    """Fill table cells using Tab index navigation.

    Handles merged cells correctly by using sequential Tab traversal
    instead of row/col coordinate navigation.

    cells: list of {"tab": int, "text": str, "style": {...} (optional)}
    """
    # Filter out cells with invalid tab values
    cells = [c for c in cells if isinstance(c.get("tab"), int) and c["tab"] >= 0]

    if not cells:
        return {"filled": 0, "failed": 0, "errors": []}

    result = {"filled": 0, "failed": 0, "errors": []}

    # Sort by tab index for sequential forward navigation
    sorted_cells = sorted(cells, key=lambda c: c["tab"])

    try:
        hwp.get_into_nth_table(table_idx)
        current_tab = 0

        for cell in sorted_cells:
            try:
                target_tab = cell["tab"]
                text = str(cell.get("text", ""))

                current_tab = _navigate_to_tab(hwp, table_idx, target_tab, current_tab)

                # 선택 영역 대체 — style 지정 시 명시적 서식, 미지정 시 기존 서식 상속
                hwp.HAction.Run("SelectAll")
                cell_style = cell.get("style")
                if cell_style:
                    insert_text_with_style(hwp, text, cell_style)
                    # 셀 정렬은 텍스트 삽입 후 TableCellAlign 액션으로 적용
                    if "align" in cell_style:
                        align_action = {
                            "left": "TableCellAlignLeftCenter",
                            "center": "TableCellAlignCenterCenter",
                            "right": "TableCellAlignRightCenter",
                            "justify": "TableCellAlignLeftCenter",
                        }
                        action = align_action.get(cell_style["align"], "TableCellAlignLeftCenter")
                        hwp.HAction.Run(action)
                else:
                    hwp.insert_text(text)

                # 셀 수직 정렬/여백 설정 (HCell)
                vert_align = cell.get("vert_align")  # "top"|"middle"|"bottom"
                if vert_align:
                    try:
                        va_map = {"top": 0, "middle": 1, "bottom": 2}
                        pset_cell = hwp.HParameterSet.HCell
                        hwp.HAction.GetDefault("CellShape", pset_cell.HSet)
                        pset_cell.VertAlign = va_map.get(vert_align, 0)
                        hwp.HAction.Execute("CellShape", pset_cell.HSet)
                    except Exception as e:
                        print(f"[WARN] VertAlign: {e}", file=sys.stderr)

                result["filled"] += 1

            except Exception as e:
                result["failed"] += 1
                result["errors"].append(
                    f"Table{table_idx} tab{cell.get('tab')} failed: {e}"
                )
                print(f"[WARN] Tab cell fill error: {e}", file=sys.stderr)

    finally:
        # 표 안전 탈출 (MovePos(3)으로 문서 끝 이동)
        try:
            if hwp.is_cell():
                hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] Table exit (fill_by_tab): {e}", file=sys.stderr)

    return result


def smart_fill_table_cells(hwp, table_idx, cells):
    """서식 감지 후 자동 적용하는 표 셀 채우기.

    각 셀에 진입 → 기존 서식 읽기 → 서식 보존하며 텍스트 삽입.
    적용된 서식 정보도 함께 반환하여 AI가 서식을 "볼 수 있게" 함.

    cells: [{"tab": int, "text": str}, ...]
    """
    cells = [c for c in cells if isinstance(c.get("tab"), int) and c["tab"] >= 0]
    if not cells:
        return {"filled": 0, "failed": 0, "errors": [], "formats_applied": []}

    result = {"filled": 0, "failed": 0, "errors": [], "formats_applied": []}
    sorted_cells = sorted(cells, key=lambda c: c["tab"])

    try:
        hwp.get_into_nth_table(table_idx)
        current_tab = 0

        for cell in sorted_cells:
            try:
                target_tab = cell["tab"]
                text = str(cell.get("text", ""))

                current_tab = _navigate_to_tab(hwp, table_idx, target_tab, current_tab)

                detected_char = get_char_shape(hwp)
                detected_para = get_para_shape(hwp)

                hwp.HAction.Run("SelectAll")
                hwp.insert_text(text)

                # v0.7.4.8 Fix D5: 감지한 format 을 실제로 재적용 (이전: read-only)
                format_reapplied = False
                try:
                    # 셀 안의 모든 텍스트 재선택
                    hwp.HAction.Run("SelectAll")
                    if isinstance(detected_char, dict):
                        char_style = {}
                        if "font_name" in detected_char and detected_char["font_name"]:
                            char_style["font_name"] = detected_char["font_name"]
                        if "font_size" in detected_char and detected_char["font_size"]:
                            char_style["font_size"] = detected_char["font_size"]
                        if detected_char.get("bold"):
                            char_style["bold"] = True
                        if detected_char.get("italic"):
                            char_style["italic"] = True
                        if detected_char.get("color") and detected_char["color"] != [0, 0, 0]:
                            char_style["color"] = detected_char["color"]
                        if char_style:
                            # CharShape ParameterSet 적용 — 현재 선택된 텍스트에
                            try:
                                _pset = hwp.HParameterSet.HCharShape
                                hwp.HAction.GetDefault("CharShape", _pset.HSet)
                                if "font_name" in char_style:
                                    _pset.FaceNameHangul = char_style["font_name"]
                                    _pset.FaceNameLatin = char_style["font_name"]
                                if "font_size" in char_style:
                                    _pset.Height = int(float(char_style["font_size"]) * 100)
                                if char_style.get("bold"):
                                    _pset.Bold = 1
                                if char_style.get("italic"):
                                    _pset.Italic = 1
                                if char_style.get("color"):
                                    c = char_style["color"]
                                    try:
                                        _pset.TextColor = hwp.RGBColor(c[0], c[1], c[2])
                                    except Exception:
                                        pass
                                hwp.HAction.Execute("CharShape", _pset.HSet)
                                format_reapplied = True
                            except Exception as fe:
                                print(f"[WARN] format reapply failed tab{target_tab}: {fe}",
                                      file=sys.stderr)
                    # 선택 해제
                    try:
                        hwp.HAction.Run("Cancel")
                    except Exception:
                        hwp.MovePos(3)
                except Exception as reapply_e:
                    print(f"[WARN] format re-selection failed tab{target_tab}: {reapply_e}",
                          file=sys.stderr)

                result["filled"] += 1
                result["formats_applied"].append({
                    "tab": target_tab,
                    "char": detected_char,
                    "para": detected_para,
                    "reapplied": format_reapplied,  # v0.7.4.8: 재적용 성공 여부
                })

            except Exception as e:
                result["failed"] += 1
                result["errors"].append(f"Table{table_idx} tab{cell.get('tab')} failed: {e}")
                print(f"[WARN] Smart fill error: {e}", file=sys.stderr)

    finally:
        try:
            if hwp.is_cell():
                hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] Table exit (smart_fill): {e}", file=sys.stderr)

    return result


def fill_table_cells_by_label(hwp, table_idx, cells):
    """라벨 기반으로 표 셀을 채운다.

    cells: [{"label": str, "text": str, "direction": "right"|"below" (optional)}, ...]

    1. resolve_labels_to_tabs()로 tab 인덱스 확보
    2. fill_table_cells_by_tab()으로 실제 채우기
    3. 매칭 실패한 라벨은 errors에 포함
    """
    # Lazy import: hwp_analyzer (utility package)
    from hwp_analyzer import resolve_labels_to_tabs

    resolution = resolve_labels_to_tabs(hwp, table_idx, cells)
    resolved = resolution.get("resolved", [])
    errors = resolution.get("errors", [])

    result = {"filled": 0, "failed": len(errors), "errors": list(errors)}

    if resolved:
        tab_cells = [{"tab": r["tab"], "text": r["text"]} for r in resolved]
        tab_result = fill_table_cells_by_tab(hwp, table_idx, tab_cells)
        result["filled"] += tab_result["filled"]
        result["failed"] += tab_result["failed"]
        result["errors"].extend(tab_result["errors"])

    # Include matched labels info for debugging
    result["matched"] = [
        {"label": r["matched_label"], "tab": r["tab"]} for r in resolved
    ]

    return result


def verify_after_fill(hwp, table_idx, expected_cells):
    """채우기 후 실제 값을 map_table_cells로 대조하여 검증.

    expected_cells: [{"tab": int, "text": str}, ...]
    Returns: {"verified": int, "mismatched": int, "details": [...]}
    """
    # Lazy import: hwp_analyzer (utility package)
    from hwp_analyzer import map_table_cells

    actual = map_table_cells(hwp, table_idx)
    actual_map = {c["tab"]: c["text"] for c in actual.get("cell_map", [])}

    result = {"verified": 0, "mismatched": 0, "details": []}
    for cell in expected_cells:
        tab = cell["tab"]
        expected_text = str(cell.get("text", "")).strip()
        actual_text = actual_map.get(tab, "").strip()

        if expected_text in actual_text or actual_text in expected_text:
            result["verified"] += 1
        else:
            result["mismatched"] += 1
            result["details"].append({
                "tab": tab,
                "expected": expected_text[:50],
                "actual": actual_text[:50],
            })

    return result


def set_cell_background_color(hwp, table_idx, cells):
    """표 셀의 배경색을 설정합니다.

    table_idx: 표 인덱스 (-1이면 현재 위치한 표)
    cells: [{"tab": int, "color": "#RRGGBB"}, ...]
    """
    if not cells:
        return {"status": "ok", "colored": 0}

    result = {"colored": 0, "failed": 0, "errors": []}
    sorted_cells = sorted(cells, key=lambda c: c["tab"])

    try:
        if table_idx >= 0:
            hwp.get_into_nth_table(table_idx)
        current_tab = 0

        for cell in sorted_cells:
            try:
                target_tab = cell["tab"]
                color_hex = cell.get("color", "#E8E8E8")
                r, g, b = _hex_to_rgb(color_hex)

                # 셀 이동
                if table_idx >= 0:
                    current_tab = _navigate_to_tab(hwp, table_idx, target_tab, current_tab)
                else:
                    # 현재 표에서 탭 이동
                    moves = target_tab - current_tab
                    if moves < 0:
                        hwp.get_into_nth_table(0)
                        moves = target_tab
                        current_tab = 0
                    for _ in range(moves):
                        hwp.TableRightCell()
                    current_tab = target_tab

                # 셀 배경색 설정 (pyhwpx 내장 cell_fill 사용)
                hwp.cell_fill((r, g, b))

                # 어두운 배경색이면 자동으로 흰색 글자 + Bold + 가운데 정렬
                brightness = (r * 299 + g * 587 + b * 114) / 1000
                if brightness < 160 and cell.get("auto_text_style", True):
                    try:
                        hwp.HAction.Run("SelectAll")
                        act = hwp.HAction
                        cs = hwp.HParameterSet.HCharShape
                        act.GetDefault("CharShape", cs.HSet)
                        cs.TextColor = hwp.RGBColor(255, 255, 255)
                        cs.Bold = 1
                        act.Execute("CharShape", cs.HSet)
                        hwp.HAction.Run("TableCellAlignCenterCenter")
                    except Exception:
                        pass

                result["colored"] += 1

            except Exception as e:
                result["failed"] += 1
                result["errors"].append(f"tab {cell.get('tab')}: {e}")
                print(f"[WARN] Cell color error: {e}", file=sys.stderr)

    finally:
        # 표 안전 탈출 (MovePos(3)으로 문서 끝 이동)
        try:
            if hwp.is_cell():
                hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] Table exit (set_cell_color): {e}", file=sys.stderr)

    return {"status": "ok", **result}


def set_table_border_style(hwp, table_idx, cells=None, style=None):
    """표 테두리 스타일을 설정합니다.

    table_idx: 표 인덱스
    cells: 특정 셀만 적용 시 [{"tab": int}, ...] (None이면 표 전체)
    style: {
        "line_type": int,       # 0=없음, 1=실선, 2=파선, 3=점선, 4=1점쇄선, 5=2점쇄선
        "line_width": int,      # pt 단위
        "color": "#RRGGBB",     # 테두리 색상
        "edges": ["left","right","top","bottom"],  # 적용할 방향 (생략 시 전체)
    }
    """
    if style is None:
        style = {}
    line_type = style.get("line_type", 1)
    line_width = style.get("line_width", 0)
    border_color = style.get("color")  # "#RRGGBB" 또는 None
    edges = style.get("edges", ["left", "right", "top", "bottom"])  # 기본: 전체

    try:
        hwp.get_into_nth_table(table_idx)

        if cells:
            # 특정 셀만 테두리 적용
            sorted_cells = sorted(cells, key=lambda c: c["tab"])
            current_tab = 0
            modified = 0
            for cell in sorted_cells:
                try:
                    target_tab = cell["tab"]
                    current_tab = _navigate_to_tab(hwp, table_idx, target_tab, current_tab)
                    act = hwp.HAction
                    pset = hwp.HParameterSet.HCellBorderFill
                    act.GetDefault("CellBorderFill", pset.HSet)
                    # 방향별 테두리 설정
                    edge_map = {"left": "Left", "right": "Right", "top": "Top", "bottom": "Bottom"}
                    for edge_name in edges:
                        prop = edge_map.get(edge_name)
                        if prop:
                            setattr(pset, f"BorderType{prop}", line_type)
                            if line_width:
                                setattr(pset, f"BorderWidth{prop}", line_width)
                            if border_color:
                                try:
                                    r, g, b = _hex_to_rgb(border_color)
                                    color_attr = f"BorderColor{prop}" if prop != "Left" else "BorderCorlorLeft"  # typo in COM
                                    setattr(pset, color_attr, hwp.RGBColor(r, g, b))
                                except Exception as e:
                                    print(f"[WARN] BorderColor {prop}: {e}", file=sys.stderr)
                    act.Execute("CellBorderFill", pset.HSet)
                    modified += 1
                except Exception as e:
                    print(f"[WARN] Border error tab {cell.get('tab')}: {e}", file=sys.stderr)
            return {"status": "ok", "modified": modified}
        else:
            # 표 전체: 표 블록 선택 후 적용
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.Run("TableCellBlock")
            act = hwp.HAction
            pset = hwp.HParameterSet.HCellBorderFill
            act.GetDefault("CellBorderFill", pset.HSet)
            edge_map = {"left": "Left", "right": "Right", "top": "Top", "bottom": "Bottom"}
            for edge_name in edges:
                prop = edge_map.get(edge_name)
                if prop:
                    setattr(pset, f"BorderType{prop}", line_type)
                    if line_width:
                        setattr(pset, f"BorderWidth{prop}", line_width)
                    if border_color:
                        try:
                            r, g, b = _hex_to_rgb(border_color)
                            color_attr = f"BorderColor{prop}" if prop != "Left" else "BorderCorlorLeft"
                            setattr(pset, color_attr, hwp.RGBColor(r, g, b))
                        except Exception as e:
                            print(f"[WARN] BorderColor {prop}: {e}", file=sys.stderr)
            act.Execute("CellBorderFill", pset.HSet)
            return {"status": "ok", "applied": "whole_table"}
    finally:
        try:
            if hwp.is_cell():
                hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] Table exit (set_border): {e}", file=sys.stderr)
