"""Table post-processing — 표 cell 크기/정렬/border 일반화 (Phase 5E).

Phase 4 의 hardcode (font_size 9pt 모든 cells) 를 대체 — **rule 기반 동적 적응**.

사용자 핵심 원칙 (v0.7.10+):
> "수천개의 양식을 모두 학습할 수 없기에 그때그때 양식에 따라서 달라짐을 알아야합니다"

기능:
1. `auto_fit_font_size`: cell 의 text 길이에 따라 font_size 동적 결정
   - ≤4 char: 12pt
   - 5-7 char: 11pt
   - 8-10 char: 10pt
   - 11-15 char: 9pt
   - 16+ char: 8pt
2. `auto_align`: text 타입 기반 정렬
   - 숫자 (단위 포함): right
   - 짧은 한글 (≤5 char): center
   - 긴 텍스트: left
3. 표 width 균등 분배 (기존 `hwp_table_distribute_width` 활용)
4. `smart_fill_table_auto`: type 기반 자동 cell mapping (classify_table_type 결과 활용)
"""
from __future__ import annotations

import re
import sys
from typing import Any, Dict, List, Optional


_NUMBER_RE = re.compile(r"^[\d,\.]+(?:\s*(?:%|원|만원|억원|조원|달러|명|대|건|개|년|월|일))?\s*$")


def auto_font_size(text: str) -> int:
    """Text 길이 기반 동적 font size (rule)."""
    n = len(text)
    if n <= 4:
        return 12
    if n <= 7:
        return 11
    if n <= 10:
        return 10
    if n <= 15:
        return 9
    return 8


def auto_align(text: str) -> str:
    """Text 타입 기반 동적 정렬 (rule)."""
    stripped = text.strip()
    if not stripped:
        return "center"
    # 숫자 (단위 포함) → right
    if _NUMBER_RE.match(stripped):
        return "right"
    # 짧은 한글/영문 (≤5 char) → center
    if len(stripped) <= 5:
        return "center"
    # 긴 텍스트 → left
    return "left"


def apply_auto_style(cells: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Cells 리스트의 각 cell 에 `style.font_size` + `style.align` 자동 적용.

    Hardcode 없음 — cell text 기반 dynamic.
    """
    result = []
    for c in cells:
        new_c = dict(c)
        text = str(c.get("text", ""))
        style = dict(c.get("style") or {})
        if "font_size" not in style:
            style["font_size"] = auto_font_size(text)
        if "align" not in style:
            style["align"] = auto_align(text)
        new_c["style"] = style
        result.append(new_c)
    return result


def smart_fill_table_auto(hwp, table_idx: int, cells: List[Dict[str, Any]], table_type: Optional[str] = None) -> Dict[str, Any]:
    """Type 기반 auto-fit 표 채우기 (Phase 5E).

    1. apply_auto_style 로 cells 에 font_size/align 자동 적용
    2. fill_table_cells_by_tab 호출
    3. (선택) table_distribute_width 호출

    Args:
        hwp: pyhwpx Hwp 인스턴스
        table_idx: 0-based table index
        cells: list of {tab, text, style?} — style 없어도 auto 적용
        table_type: optional — `classify_table_type` 결과 (patent/market_size/...)

    Returns:
        {status, filled, failed, style_applied}
    """
    try:
        from hwp_editor.tables import fill_table_cells_by_tab
    except Exception as e:
        return {"status": "error", "error": f"fill_table_cells_by_tab import: {e}"}

    # Auto style 적용
    styled_cells = apply_auto_style(cells)

    # fill_table_cells_by_tab 호출
    result = fill_table_cells_by_tab(hwp, table_idx, styled_cells)

    # (선택) width 균등 분배 — type 이 year-based 표면 유용
    width_applied = False
    if table_type in ("market_size", "sales_plan", "revenue", "personnel_plan"):
        try:
            hwp.get_into_nth_table(table_idx)
            hwp.HAction.Run("TableDistributeCellWidth")
            if hwp.is_cell():
                hwp.MovePos(3)
            width_applied = True
        except Exception as e:
            print(f"[WARN] smart_fill_table_auto width: {e}", file=sys.stderr)

    return {
        "status": "ok",
        "filled": result.get("filled", 0),
        "failed": result.get("failed", 0),
        "errors": result.get("errors", []),
        "style_applied": True,
        "width_applied": width_applied,
        "table_type": table_type,
    }
