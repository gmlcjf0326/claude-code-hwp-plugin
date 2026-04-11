"""hwp_core.analysis.metadata — 페이지/커서 메타 정보 read-only 핸들러.

Handlers:
- get_page_setup      : F7 용지편집 정보 — 용지 크기/방향/여백
- get_cursor_context  : 현재 커서 위치의 char/para shape + 페이지/위치
"""
from .. import register  # noqa: F401 — Phase 6 gotcha: 두 점!


@register("get_page_setup")
def get_page_setup(hwp, params):
    """F7 용지편집 정보 — 용지 크기/방향/여백/사용 가능 영역."""
    try:
        d = hwp.get_pagedef_as_dict()
        pw = d.get("용지폭", 210)
        ph = d.get("용지길이", 297)
        lm = d.get("왼쪽", 30)
        rm = d.get("오른쪽", 30)
        tm = d.get("위쪽", 20)
        bm = d.get("아래쪽", 15)
        hm = d.get("머리말", 15)
        fm = d.get("꼬리말", 15)
        orient = d.get("용지방향", 0)
        binding = d.get("제본여백", 0)
        return {
            "status": "ok",
            "paper_width_mm": pw,
            "paper_height_mm": ph,
            "orientation": "landscape" if orient == 1 else "portrait",
            "top_margin_mm": tm,
            "bottom_margin_mm": bm,
            "left_margin_mm": lm,
            "right_margin_mm": rm,
            "header_margin_mm": hm,
            "footer_margin_mm": fm,
            "binding_margin_mm": binding,
            "usable_width_mm": round(pw - lm - rm, 1),
            "usable_height_mm": round(ph - tm - bm, 1),
        }
    except Exception as e:
        raise RuntimeError(f"용지 설정 읽기 실패: {e}")


@register("get_cursor_context")
def get_cursor_context(hwp, params):
    """실제 커서 위치의 서식 + 주변 페이지/줄/컬럼 정보 반환."""
    from hwp_editor import get_char_shape, get_para_shape
    context = {"status": "ok"}
    try:
        context["char_shape"] = get_char_shape(hwp)
    except Exception as e:
        context["char_shape"] = {"error": str(e)}
    try:
        context["para_shape"] = get_para_shape(hwp)
    except Exception as e:
        context["para_shape"] = {"error": str(e)}
    try:
        pos = hwp.GetPos()
        context["position"] = list(pos) if pos else None
    except Exception:
        context["position"] = None
    try:
        context["total_pages"] = hwp.PageCount
    except Exception:
        context["total_pages"] = None
    try:
        # KeyIndicator: (섹션, 페이지, 줄, 컬럼, 삽입/수정, 줄번호)
        ki = hwp.KeyIndicator()
        context["current_page"] = ki[1] if ki else None
    except Exception:
        context["current_page"] = None
    return context
