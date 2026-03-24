"""HWP Document Editor - Fill and modify HWP documents.
Uses pyhwpx Hwp() only. Raw win32com is forbidden.
All file paths must use os.path.abspath().

Cell navigation uses sequential Tab (TableRightCell) traversal,
which handles merged cells better than row/col coordinate addressing.
"""
import sys
import os


def insert_text_with_color(hwp, text, rgb=None):
    """텍스트를 지정 색상으로 삽입. rgb=(r,g,b) 또는 None(기본색)"""
    if not rgb:
        hwp.insert_text(text)
        return

    act = hwp.HAction
    pset = hwp.HParameterSet.HCharShape
    try:
        act.GetDefault("CharShape", pset.HSet)
        pset.TextColor = hwp.RGBColor(rgb[0], rgb[1], rgb[2])
        act.Execute("CharShape", pset.HSet)
        hwp.insert_text(text)
    finally:
        # 색상 복원 (기본 검정) — 에러 시에도 반드시 실행
        try:
            act.GetDefault("CharShape", pset.HSet)
            pset.TextColor = hwp.RGBColor(0, 0, 0)
            act.Execute("CharShape", pset.HSet)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)


def insert_text_with_style(hwp, text, style=None):
    """서식 지정 텍스트 삽입.
    style: {
        "color": [r,g,b],        # 글자 색상
        "bold": True/False,       # 굵게
        "italic": True/False,     # 기울임
        "underline": True/False,  # 밑줄
        "font_size": 12.0,        # 글자 크기 (pt)
        "font_name": "맑은 고딕", # 글꼴
        "bg_color": [r,g,b],      # 배경 색상
        "strikeout": True/False,  # 취소선
        "char_spacing": -5,       # 자간 (%, 기본 0)
        "width_ratio": 90,        # 장평 (%, 기본 100)
        "font_name_hanja": "바탕",# 한자 글꼴
        "font_name_japanese": "", # 일본어 글꼴
    }
    삽입 후 원래 서식으로 복원.
    """
    if not style:
        hwp.insert_text(text)
        return

    act = hwp.HAction
    pset = hwp.HParameterSet.HCharShape

    # 현재 서식 저장
    act.GetDefault("CharShape", pset.HSet)
    saved_color = pset.TextColor
    saved_bold = pset.Bold
    saved_italic = pset.Italic
    saved_underline = pset.UnderlineType
    saved_height = pset.Height
    saved_strikeout = pset.StrikeOutType
    saved_char_spacing = None
    saved_width_ratio = None
    try:
        saved_char_spacing = pset.SpacingHangul
    except Exception:
        pass
    try:
        saved_width_ratio = pset.RatioHangul
    except Exception:
        pass

    # 새 서식 적용
    act.GetDefault("CharShape", pset.HSet)

    if "color" in style:
        c = style["color"]
        pset.TextColor = hwp.RGBColor(c[0], c[1], c[2])
    if "bold" in style:
        pset.Bold = 1 if style["bold"] else 0
    if "italic" in style:
        pset.Italic = 1 if style["italic"] else 0
    if "underline" in style:
        pset.UnderlineType = 1 if style["underline"] else 0
    if "font_size" in style:
        pset.Height = int(style["font_size"] * 100)  # pt → HWP 단위
    if "font_name" in style:
        pset.FaceNameHangul = style["font_name"]
        pset.FaceNameLatin = style["font_name"]
    if "bg_color" in style:
        bg = style["bg_color"]
        pset.ShadeColor = hwp.RGBColor(bg[0], bg[1], bg[2])
    if "strikeout" in style:
        pset.StrikeOutType = 1 if style["strikeout"] else 0
    if "char_spacing" in style:
        try:
            pset.SpacingHangul = int(style["char_spacing"])
            pset.SpacingLatin = int(style["char_spacing"])
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "width_ratio" in style:
        try:
            pset.RatioHangul = int(style["width_ratio"])
            pset.RatioLatin = int(style["width_ratio"])
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "font_name_hanja" in style:
        try:
            pset.FaceNameHanja = style["font_name_hanja"]
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "font_name_japanese" in style:
        try:
            pset.FaceNameJapanese = style["font_name_japanese"]
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

    act.Execute("CharShape", pset.HSet)

    hwp.insert_text(text)

    # 원래 서식 복원
    act.GetDefault("CharShape", pset.HSet)
    pset.TextColor = saved_color
    pset.Bold = saved_bold
    pset.Italic = saved_italic
    pset.UnderlineType = saved_underline
    pset.Height = saved_height
    pset.StrikeOutType = saved_strikeout
    if saved_char_spacing is not None:
        try:
            pset.SpacingHangul = saved_char_spacing
            pset.SpacingLatin = saved_char_spacing
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if saved_width_ratio is not None:
        try:
            pset.RatioHangul = saved_width_ratio
            pset.RatioLatin = saved_width_ratio
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    act.Execute("CharShape", pset.HSet)


def set_paragraph_style(hwp, style=None):
    """현재 커서 위치의 단락 서식을 설정.
    style: {
        "align": "left"|"center"|"right"|"justify",  # 정렬
        "line_spacing": 160,    # 줄간격 (%)
        "space_before": 0,      # 문단 앞 간격 (pt)
        "space_after": 0,       # 문단 뒤 간격 (pt)
        "indent": 0,            # 들여쓰기 (pt)
    }
    """
    if not style:
        return

    act = hwp.HAction
    pset = hwp.HParameterSet.HParaShape

    act.GetDefault("ParaShape", pset.HSet)

    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
    if "align" in style:
        try:
            pset.AlignType = align_map.get(style["align"], 0)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "line_spacing" in style:
        try:
            pset.LineSpacingType = style.get("line_spacing_type", 0)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        try:
            pset.LineSpacing = int(style["line_spacing"])
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "space_before" in style:
        try:
            pset.PrevSpacing = int(style["space_before"] * 100)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "space_after" in style:
        try:
            pset.NextSpacing = int(style["space_after"] * 100)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "indent" in style:
        try:
            pset.Indentation = int(style["indent"] * 100)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "left_margin" in style:
        try:
            pset.LeftMargin = int(style["left_margin"] * 100)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
    if "right_margin" in style:
        try:
            pset.RightMargin = int(style["right_margin"] * 100)
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

    act.Execute("ParaShape", pset.HSet)


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
    act.GetDefault("ParaShape", pset.HSet)

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

    return {
        "align": align_names.get(alignment, "left"),
        "line_spacing": line_spacing,
        "line_spacing_type": spacing_type_names.get(line_spacing_type, "percent"),
        "space_before": space_before,
        "space_after": space_after,
        "indent": indent,              # 첫 줄 들여쓰기 (양수=들여쓰기, 음수=내어쓰기)
        "left_margin": left_margin,    # 왼쪽 여백 = 나머지 줄 시작위치
        "right_margin": right_margin,  # 오른쪽 여백
        # 첫 줄 시작위치 = left_margin + indent
    }


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
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)


def get_table_format_summary(hwp, table_idx, sample_tabs=None):
    """표 전체의 서식 요약을 반환. sample_tabs 미지정 시 첫 5개 + 마지막 셀."""
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


def _goto_cell(hwp, table_idx, cell_positions, target_cell_idx):
    """Navigate to a specific cell by its sequential index using Tab."""
    hwp.get_into_nth_table(table_idx)

    for _ in range(target_cell_idx):
        try:
            hwp.TableRightCell()
        except Exception:
            break


def _navigate_to_tab(hwp, table_idx, target_tab, current_tab):
    """셀 네비게이션 공통 로직. 새 current_tab을 반환."""
    moves = target_tab - current_tab
    if moves < 0:
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
        hwp.get_into_nth_table(table_idx)
        moves = target_tab
    for _ in range(moves):
        hwp.TableRightCell()
    return target_tab


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
                else:
                    hwp.insert_text(text)
                result["filled"] += 1

            except Exception as e:
                result["failed"] += 1
                result["errors"].append(
                    f"Table{table_idx} tab{cell.get('tab')} failed: {e}"
                )
                print(f"[WARN] Tab cell fill error: {e}", file=sys.stderr)

    finally:
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

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

                result["filled"] += 1
                result["formats_applied"].append({
                    "tab": target_tab,
                    "char": detected_char,
                    "para": detected_para,
                })

            except Exception as e:
                result["failed"] += 1
                result["errors"].append(f"Table{table_idx} tab{cell.get('tab')} failed: {e}")
                print(f"[WARN] Smart fill error: {e}", file=sys.stderr)

    finally:
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

    return result


def insert_markdown(hwp, md_text):
    """마크다운 텍스트를 한글 서식으로 변환하여 삽입.

    지원: # 제목, **굵게**, *기울임*, - 목록, | 표 |, > 인용, --- 구분선.
    BUG-5 fix: 마크다운 표 파싱 추가.
    """
    import re

    lines = md_text.split('\n')
    inserted = 0
    i = 0

    while i < len(lines):
        stripped = lines[i].strip()

        if not stripped:
            hwp.insert_text('\r\n')
            i += 1
            continue

        # 수평선 (---, ***, ___)
        if re.match(r'^[-*_]{3,}$', stripped):
            hwp.insert_text('─' * 40 + '\r\n')
            inserted += 1
            i += 1
            continue

        # 마크다운 표 (| 로 시작하는 연속된 줄)
        if stripped.startswith('|') and '|' in stripped[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                row_text = lines[i].strip()
                # 구분선(|---|---|) 건너뛰기
                if re.match(r'^\|[\s\-:]+\|', row_text):
                    i += 1
                    continue
                # 셀 파싱
                # H3 fix: 빈 셀 유지 (앞뒤 빈 요소만 제거)
                raw_cells = row_text.split('|')
                cells = [c.strip() for c in raw_cells[1:-1]]  # | 앞뒤 빈 요소 제거, 중간 빈 셀 유지
                table_lines.append(cells)
                i += 1
            # 표 생성
            if table_lines:
                rows = len(table_lines)
                cols = max(len(row) for row in table_lines)
                hwp.create_table(rows, cols)
                for r, row in enumerate(table_lines):
                    for c in range(cols):
                        val = row[c] if c < len(row) else ''
                        if val:
                            if r == 0:
                                insert_text_with_style(hwp, val, {"bold": True})
                            else:
                                hwp.insert_text(val)
                        if c < cols - 1 or r < rows - 1:
                            hwp.TableRightCell()
                try:
                    hwp.Cancel()
                except Exception:
                    pass
                hwp.insert_text('\r\n')
                inserted += 1
            continue

        # 인용문 (>)
        if stripped.startswith('>'):
            quote_text = stripped.lstrip('>').strip()
            hwp.insert_text('  │ ' + quote_text + '\r\n')
            inserted += 1
            i += 1
            continue

        # 제목 (# ~ ###)
        heading_match = re.match(r'^(#{1,3})\s+(.+)$', stripped)
        if heading_match:
            level = len(heading_match.group(1))
            title_text = heading_match.group(2)
            sizes = {1: 22, 2: 16, 3: 13}
            insert_text_with_style(hwp, title_text + '\r\n', {
                "bold": True,
                "font_size": sizes.get(level, 13),
            })
            inserted += 1
            i += 1
            continue

        # 목록 (- 또는 *)
        list_match = re.match(r'^[\-\*]\s+(.+)$', stripped)
        if list_match:
            hwp.insert_text('  ◦ ' + list_match.group(1) + '\r\n')
            inserted += 1
            i += 1
            continue

        # 번호 목록
        numbered_match = re.match(r'^(\d+)\.\s+(.+)$', stripped)
        if numbered_match:
            hwp.insert_text('  ' + numbered_match.group(1) + '. ' + numbered_match.group(2) + '\r\n')
            inserted += 1
            i += 1
            continue

        # 인라인 서식 처리 (**굵게**, *기울임*)
        parts = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                insert_text_with_style(hwp, part[2:-2], {"bold": True})
            elif part.startswith('*') and part.endswith('*'):
                insert_text_with_style(hwp, part[1:-1], {"italic": True})
            else:
                hwp.insert_text(part)
        hwp.insert_text('\r\n')
        inserted += 1
        i += 1

    return {"status": "ok", "lines_inserted": inserted}


def auto_map_reference_to_table(hwp, table_idx, ref_headers, ref_row):
    """참고자료의 헤더와 표의 라벨을 자동 매칭하여 채울 데이터 생성.

    ref_headers: ["기업명", "대표자", "전화번호", ...]
    ref_row: ["(주)플랜아이", "이명기", "042-934-3508", ...]

    Returns: {"mappings": [{header, matched_label, tab, text}, ...], "unmapped": [...]}
    """
    from hwp_analyzer import map_table_cells, _match_label

    cell_data = map_table_cells(hwp, table_idx)
    cell_map = cell_data.get("cell_map", [])

    mappings = []
    unmapped = []

    for i, header in enumerate(ref_headers):
        if i >= len(ref_row):
            break
        value = ref_row[i]
        if not value or not header:
            continue

        matched = False
        for j, cell in enumerate(cell_map):
            is_match, is_exact, ratio = _match_label(cell["text"], header)
            if is_match and (is_exact or ratio > 0.5):
                target_tab = j + 1
                if target_tab < len(cell_map):
                    mappings.append({
                        "header": header,
                        "matched_label": cell["text"].strip()[:30],
                        "tab": target_tab,
                        "text": str(value),
                    })
                    matched = True
                    break
        if not matched:
            unmapped.append({"header": header, "value": str(value)})

    return {"mappings": mappings, "unmapped": unmapped, "total_matched": len(mappings)}


def insert_picture(hwp, file_path, width=0, height=0):
    """현재 커서 위치에 이미지 삽입.

    file_path: 이미지 파일 경로
    width/height: mm 단위 (0이면 원본 크기)
    """
    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {file_path}")

    # pyhwpx insert_picture: Width/Height는 HWPUNIT (1mm = 283.46 HWP 단위)
    hwp.insert_picture(file_path,
                        Width=int(width * 283.46) if width else 0,
                        Height=int(height * 283.46) if height else 0)
    return {"status": "ok", "file_path": file_path, "width_mm": width, "height_mm": height}


def fill_table_cells_by_label(hwp, table_idx, cells):
    """라벨 기반으로 표 셀을 채운다.

    cells: [{"label": str, "text": str, "direction": "right"|"below" (optional)}, ...]

    1. resolve_labels_to_tabs()로 tab 인덱스 확보
    2. fill_table_cells_by_tab()으로 실제 채우기
    3. 매칭 실패한 라벨은 errors에 포함
    """
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


def fill_document(hwp, fill_data):
    """Fill document with AI-generated content.

    fill_data format:
    {
        "file_path": "...",           # optional: open file first
        "fields": {"name": "value"},  # field-based fill
        "tables": [                   # table-based fill
            {
                "index": 0,
                "cells": [
                    {"row": 0, "col": 0, "text": "value"},
                    ...
                ]
            }
        ]
    }
    """
    # Open file if specified
    if "file_path" in fill_data:
        file_path = os.path.abspath(fill_data["file_path"])
        hwp.open(file_path)

    result = {"filled": 0, "failed": 0, "errors": []}

    # Fill fields
    if "fields" in fill_data and fill_data["fields"]:
        try:
            hwp.put_field_text(fill_data["fields"])
            result["filled"] += len(fill_data["fields"])
        except Exception as e:
            result["errors"].append(f"Field fill failed: {e}")
            result["failed"] += len(fill_data["fields"])

    # Fill tables - each cell independently (re-enter table each time)
    if "tables" in fill_data:
        for table_data in fill_data["tables"]:
            table_idx = table_data.get("index", 0)
            cells = table_data.get("cells", [])

            # Split: tab-based cells vs row/col cells
            tab_cells = [c for c in cells if "tab" in c]
            rowcol_cells = [c for c in cells if "tab" not in c]

            if tab_cells:
                tab_result = fill_table_cells_by_tab(hwp, table_idx, tab_cells)
                result["filled"] += tab_result["filled"]
                result["failed"] += tab_result["failed"]
                result["errors"].extend(tab_result["errors"])

            for cell in rowcol_cells:
                try:
                    row = cell.get("row", 0)
                    col = cell.get("col", 0)
                    text = str(cell.get("text", cell.get("value", "")))

                    # Enter table fresh each time
                    hwp.get_into_nth_table(table_idx)

                    try:
                        # Navigate: first go down, then right
                        # Use HAction.Run for more reliable navigation
                        for _ in range(row):
                            hwp.HAction.Run("TableLowerCell")
                        for _ in range(col):
                            hwp.HAction.Run("TableRightCell")

                        # 선택 영역 대체 — 기존 서식 상속
                        hwp.HAction.Run("SelectAll")
                        hwp.insert_text(text)
                        result["filled"] += 1

                    finally:
                        try:
                            hwp.Cancel()
                        except Exception:
                            pass

                except Exception as e:
                    result["failed"] += 1
                    result["errors"].append(
                        f"Table{table_idx} ({row},{col}) failed: {e}"
                    )
                    print(f"[WARN] Cell fill error: {e}", file=sys.stderr)

    return result


def _hex_to_rgb(hex_color):
    """#RRGGBB 헥스 색상을 (r, g, b) 튜플로 변환."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


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

                # 셀 배경색 설정 (CellShape → BackColor)
                act = hwp.HAction
                pset = hwp.HParameterSet.HCellShape
                act.GetDefault("CellShape", pset.HSet)
                pset.BackColor = hwp.RGBColor(r, g, b)
                act.Execute("CellShape", pset.HSet)
                result["colored"] += 1

            except Exception as e:
                result["failed"] += 1
                result["errors"].append(f"tab {cell.get('tab')}: {e}")
                print(f"[WARN] Cell color error: {e}", file=sys.stderr)

    finally:
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)

    return {"status": "ok", **result}


def set_table_border_style(hwp, table_idx, cells=None, style=None):
    """표 테두리 스타일을 설정합니다.

    table_idx: 표 인덱스
    cells: 특정 셀만 적용 시 [{"tab": int}, ...] (None이면 표 전체)
    style: {"line_type": int, "line_width": int} — HWP 테두리 속성
           line_type: 0=없음, 1=실선, 2=파선, 3=점선, 4=1점쇄선, 5=2점쇄선
           line_width: pt 단위 (기본 0.12mm → 약 0.4pt)
    """
    if style is None:
        style = {}
    line_type = style.get("line_type", 1)  # 기본: 실선
    line_width = style.get("line_width", 0)

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
                    # 4방향 테두리 설정
                    for attr in ["Left", "Right", "Top", "Bottom"]:
                        line = getattr(pset, f"{attr}Border", None)
                        if line:
                            line.Type = line_type
                            if line_width:
                                line.Width = line_width
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
            for attr in ["Left", "Right", "Top", "Bottom"]:
                line = getattr(pset, f"{attr}Border", None)
                if line:
                    line.Type = line_type
                    if line_width:
                        line.Width = line_width
            act.Execute("CellBorderFill", pset.HSet)
            return {"status": "ok", "applied": "whole_table"}
    finally:
        try:
            hwp.Cancel()
        except Exception as e:
            print(f"[WARN] {e}", file=sys.stderr)
