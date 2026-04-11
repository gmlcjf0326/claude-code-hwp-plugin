"""hwp_core.formatting.page_layout — 페이지 / 다단 / 셀 속성 / 머리글-바닥글.

Handlers:
- set_page_setup      : 여백/용지 크기/방향 (PageDef + ApplyTo=2 경로)
- set_cell_property   : 셀 여백/텍스트 방향/수직 정렬/보호
- set_header_footer   : 머리글/바닥글 CreateAction 방식
- set_column          : 다단 설정 (count/gap/line_type)

v0.7.3.4 Fix #1: pyhwpx set_pagedef(dict) silent fail → 직접 setattr + ApplyTo=2 경로
"""
import sys

from .. import register  # 두 점!
from .._helpers import validate_params, _exit_table_safely  # 두 점!


@register("set_page_setup")
def set_page_setup(hwp, params):
    """페이지 설정 — 여백/용지 크기/방향.

    v0.7.3.4 Fix #1: pyhwpx set_pagedef(dict) 는 GetDefault 호출 안해서 silent fail.
    직접 setattr + ApplyTo=2 + Execute 경로 사용.
    """
    try:
        d = {}
        if "top_margin" in params:
            d["TopMargin"] = params["top_margin"]
        if "bottom_margin" in params:
            d["BottomMargin"] = params["bottom_margin"]
        if "left_margin" in params:
            d["LeftMargin"] = params["left_margin"]
        if "right_margin" in params:
            d["RightMargin"] = params["right_margin"]
        if "header_margin" in params:
            d["HeaderLen"] = params["header_margin"]
        if "footer_margin" in params:
            d["FooterLen"] = params["footer_margin"]
        if "orientation" in params:
            d["Landscape"] = 1 if params["orientation"] == "landscape" else 0
        if "paper_width" in params:
            d["PaperWidth"] = params["paper_width"]
        if "paper_height" in params:
            d["PaperHeight"] = params["paper_height"]

        if not d:
            return {"status": "ok", "applied": []}

        act = hwp.HAction
        hsec = hwp.HParameterSet.HSecDef
        act.GetDefault("PageSetup", hsec.HSet)
        pd = hsec.PageDef

        def _mm2hu(v):
            return hwp.MiliToHwpUnit(v)

        for key, val in d.items():
            try:
                if key in ("Landscape", "GutterType"):
                    setattr(pd, key, val)
                else:
                    setattr(pd, key, _mm2hu(val))
            except Exception as e:
                print(f"[WARN] PageDef.{key}={val}: {e}", file=sys.stderr)

        hsec.HSet.SetItem("ApplyTo", 2)
        ok = act.Execute("PageSetup", hsec.HSet)
        return {"status": "ok" if ok else "partial", "applied": list(d.keys())}
    except Exception as e:
        return {"status": "error", "error": f"페이지 설정 실패: {e}"}


@register("set_cell_property")
def set_cell_property(hwp, params):
    """셀 속성 설정 — 여백/텍스트 방향/수직 정렬/보호."""
    validate_params(params, ["table_index", "tab"], "set_cell_property")
    try:
        from hwp_editor import _navigate_to_tab
        hwp.get_into_nth_table(params["table_index"])
        _navigate_to_tab(hwp, params["table_index"], params["tab"], 0)
        pset = hwp.HParameterSet.HCell
        hwp.HAction.GetDefault("CellShape", pset.HSet)
        if "vert_align" in params:
            va_map = {"top": 0, "middle": 1, "bottom": 2}
            pset.VertAlign = va_map.get(params["vert_align"], 0)
        if "margin_left" in params:
            pset.MarginLeft = hwp.MiliToHwpUnit(params["margin_left"])
        if "margin_right" in params:
            pset.MarginRight = hwp.MiliToHwpUnit(params["margin_right"])
        if "margin_top" in params:
            pset.MarginTop = hwp.MiliToHwpUnit(params["margin_top"])
        if "margin_bottom" in params:
            pset.MarginBottom = hwp.MiliToHwpUnit(params["margin_bottom"])
        if "text_direction" in params:
            pset.TextDirection = int(params["text_direction"])
        if "protected" in params:
            pset.Protected = 1 if params["protected"] else 0
        hwp.HAction.Execute("CellShape", pset.HSet)
        return {"status": "ok", "tab": params["tab"]}
    except Exception as e:
        raise RuntimeError(f"셀 속성 설정 실패: {e}")
    finally:
        _exit_table_safely(hwp)


@register("set_header_footer")
def set_header_footer(hwp, params):
    """머리글/바닥글 설정 — CreateAction 방식."""
    hf_type = params.get("type", "header")
    text = params.get("text", "")
    style = params.get("style")
    try:
        act = hwp.CreateAction("HeaderFooter")
        ps = act.CreateSet()
        act.GetDefault(ps)
        ps.SetItem("Type", 0 if hf_type == "header" else 1)
        result = act.Execute(ps)
        if not result:
            raise RuntimeError("HeaderFooter Execute 실패")
        if text and style:
            from hwp_editor import insert_text_with_style, set_paragraph_style as _set_ps
            insert_text_with_style(hwp, text, style)
            if "align" in style:
                _set_ps(hwp, {"align": style["align"]})
        elif text:
            hwp.insert_text(text)
        hwp.HAction.Run("CloseEx")
        return {"status": "ok", "type": hf_type, "text": text}
    except Exception as e:
        try:
            hwp.HAction.Run("CloseEx")
        except Exception as ex:
            print(f"[WARN] CloseEx recovery: {ex}", file=sys.stderr)
        raise RuntimeError(f"머리글/바닥글 설정 실패: {e}")


@register("set_column")
def set_column(hwp, params):
    """다단 설정."""
    count = params.get("count", 2)
    gap = params.get("gap", 10)
    line_type = params.get("line_type", 0)
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HColDef
        act.GetDefault("MultiColumn", pset.HSet)
        pset.Count = int(count)
        pset.SameSize = 1
        pset.SameGap = hwp.MiliToHwpUnit(gap)
        pset.LineType = int(line_type)
        pset.type = 1
        act.Execute("MultiColumn", pset.HSet)
        return {"status": "ok", "count": count, "gap": gap}
    except Exception as e:
        raise RuntimeError(f"다단 설정 실패: {e}")
