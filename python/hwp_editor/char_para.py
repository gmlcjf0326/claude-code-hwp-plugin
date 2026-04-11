"""hwp_editor.char_para — 문자/단락 서식 조회.

함수:
- get_char_shape             : 현재 커서 위치의 글자 서식 dict 반환
- get_para_shape             : 현재 커서 위치의 단락 서식 dict 반환 (15+ 필드)
- get_cell_format            : 특정 표 셀의 char+para 서식 조회
- get_table_format_summary   : 표 전체의 서식 요약 (샘플 셀 5개+마지막)

v0.6.0: ParaShape 11 개 추가 속성 지원
v0.6.7: RatioHangul/SpacingHangul 자간/장평 반환
"""
import sys


def get_char_shape(hwp):
    """현재 커서 위치의 글자 서식 정보를 반환."""
    act = hwp.HAction
    pset = hwp.HParameterSet.HCharShape
    act.GetDefault("CharShape", pset.HSet)

    font_hangul = ""
    font_latin = ""
    try:
        font_hangul = pset.FaceNameHangul or ""
        font_latin = pset.FaceNameLatin or ""
    except Exception:
        pass

    # 자간: SpacingHangul (언어별 분리, 한글 기준)
    # 장평: RatioHangul (언어별 분리, 한글 기준)
    char_spacing = 0
    width_ratio = 100
    try:
        char_spacing = pset.SpacingHangul
    except Exception:
        pass
    try:
        width_ratio = pset.RatioHangul
    except Exception:
        pass

    return {
        "font_name_hangul": font_hangul,
        "font_name_latin": font_latin,
        "font_size": pset.Height / 100.0,  # HWP 단위 → pt
        "bold": bool(pset.Bold),
        "italic": bool(pset.Italic),
        "underline": pset.UnderlineType,
        "strikeout": pset.StrikeOutType,
        "color": pset.TextColor,
        "char_spacing": char_spacing,
        "width_ratio": width_ratio,
    }


def get_para_shape(hwp):
    """현재 커서 위치의 단락 서식 정보를 반환.

    pyhwpx HParaShape 실제 속성명 (dir() 확인 결과):
    - AlignType (정렬), LineSpacing, LineSpacingType
    - PrevSpacing (문단 앞), NextSpacing (문단 뒤)
    - Indentation (들여쓰기), LeftMargin, RightMargin
    """
    act = hwp.HAction
    pset = hwp.HParameterSet.HParaShape
    act.GetDefault("ParagraphShape", pset.HSet)

    align_names = {0: "left", 1: "center", 2: "right", 3: "justify"}
    spacing_type_names = {0: "percent", 1: "fixed", 2: "multiple"}

    alignment = 0
    try:
        alignment = pset.AlignType
    except Exception:
        pass

    line_spacing = 160
    try:
        line_spacing = pset.LineSpacing
    except Exception:
        pass

    line_spacing_type = 0
    try:
        line_spacing_type = pset.LineSpacingType
    except Exception:
        pass

    space_before = 0
    try:
        space_before = pset.PrevSpacing / 100.0
    except Exception:
        pass

    space_after = 0
    try:
        space_after = pset.NextSpacing / 100.0
    except Exception:
        pass

    indent = 0
    try:
        val = pset.Indentation
        if val:
            indent = val / 100.0
    except Exception:
        pass

    left_margin = 0
    try:
        left_margin = pset.LeftMargin / 100.0
    except Exception:
        pass

    right_margin = 0
    try:
        right_margin = pset.RightMargin / 100.0
    except Exception:
        pass

    # 추가 속성 11개 (v0.6.0 양식 정밀 분석용)
    extra = {}
    for attr, key in [
        ("PagebreakBefore", "page_break_before"),
        ("KeepWithNext", "keep_with_next"),
        ("WidowOrphan", "widow_orphan"),
        ("KeepLinesTogether", "keep_lines_together"),
        ("AutoSpaceEAsianEng", "auto_space_eAsian_eng"),
        ("AutoSpaceEAsianNum", "auto_space_eAsian_num"),
        ("BreakLatinWord", "break_latin_word"),
        ("LineWrap", "line_wrap"),
        ("SnapToGrid", "snap_to_grid"),
        ("HeadingType", "heading_type"),
        ("Condense", "condense"),
    ]:
        try:
            val = getattr(pset, attr)
            extra[key] = bool(val) if isinstance(val, int) and val in (0, 1) else val
        except Exception:
            pass

    result = {
        "align": align_names.get(alignment, "left"),
        "line_spacing": line_spacing,
        "line_spacing_type": spacing_type_names.get(line_spacing_type, "percent"),
        "space_before": space_before,
        "space_after": space_after,
        "indent": indent,
        "indent_type": "들여쓰기" if indent > 0 else ("내어쓰기" if indent < 0 else "없음"),
        "left_margin": left_margin,
        "right_margin": right_margin,
        "first_line_start": round(left_margin + indent, 1),  # 첫 줄 시작위치
    }
    result.update(extra)
    return result


def get_cell_format(hwp, table_idx, cell_tab):
    """특정 표 셀의 글자+단락 서식을 조회.

    table_idx: 표 인덱스
    cell_tab: Tab 인덱스 (hwp_map_table_cells로 확인)
    Returns: {"char": {...}, "para": {...}, "text_preview": str}
    """
    try:
        hwp.get_into_nth_table(table_idx)
        for _ in range(cell_tab):
            hwp.TableRightCell()

        text_preview = ""
        try:
            hwp.HAction.Run("SelectAll")
            text_preview = hwp.GetTextFile("TEXT", "saveblock").strip()[:100]
            hwp.HAction.Run("Cancel")
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

        char = get_char_shape(hwp)
        para = get_para_shape(hwp)

        return {
            "table_index": table_idx,
            "cell_tab": cell_tab,
            "text_preview": text_preview,
            "char": char,
            "para": para,
        }
    finally:
        try:
            if hwp.is_cell():
                hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] Table exit (get_cell_format): {e}", file=sys.stderr)


def get_table_format_summary(hwp, table_idx, sample_tabs=None):
    """표 전체의 서식 요약을 반환. sample_tabs 미지정 시 첫 5개 + 마지막 셀."""
    # Lazy import: hwp_analyzer (utility package)
    from hwp_analyzer import map_table_cells

    cell_data = map_table_cells(hwp, table_idx)
    cell_map = cell_data.get("cell_map", [])

    if not cell_map:
        return {"table_index": table_idx, "cell_formats": [], "error": "표에 셀이 없습니다"}

    if sample_tabs is None:
        total = len(cell_map)
        tabs = list(range(min(5, total)))
        if total > 5:
            tabs.append(total - 1)
        sample_tabs = tabs

    formats = []
    for tab in sample_tabs:
        if tab >= len(cell_map):
            continue
        try:
            fmt = get_cell_format(hwp, table_idx, tab)
            fmt["text_preview"] = cell_map[tab]["text"][:50] if tab < len(cell_map) else ""
            formats.append(fmt)
        except Exception as e:
            formats.append({"cell_tab": tab, "error": str(e)})

    return {
        "table_index": table_idx,
        "total_cells": len(cell_map),
        "sampled_cells": len(formats),
        "cell_formats": formats,
    }
