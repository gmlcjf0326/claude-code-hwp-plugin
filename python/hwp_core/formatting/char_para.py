"""hwp_core.formatting.char_para — 문자/단락 서식 조회 + 설정.

Handlers:
- get_char_shape       : 커서 위치 글자 서식 조회 (hwp_editor 위임)
- get_para_shape       : 커서 위치 단락 서식 조회 (hwp_editor 위임)
- set_paragraph_style  : ParaShape 전체 + Border + indent/margin (245L 메가 함수)
                         v0.7.9 fix: SetItem 2-tier (양수 indent silent fail 방지)
"""
import sys

from .. import register  # 두 점!
from .._helpers import validate_params  # 두 점!


@register("get_char_shape")
def get_char_shape(hwp, params):
    """현재 커서 위치 글자 서식 조회."""
    from hwp_editor import get_char_shape as _get
    return _get(hwp)


@register("get_para_shape")
def get_para_shape(hwp, params):
    """현재 커서 위치 단락 서식 조회."""
    from hwp_editor import get_para_shape as _get
    return _get(hwp)


@register("set_paragraph_style")
def set_paragraph_style(hwp, params):
    """단락 서식 설정 — ParaShape + Border + indent/margin.

    v0.6.7: hwp_editor 와 인라인 풀 동기화 (8개 속성 추가).
    v0.7.2.1: ParaShape 정밀 옵션 multi-fallback.
    v0.7.3.1: Indentation 직접 attribute (SetItem 'Indent' 오류 정정).
    v0.7.3.4: LeftMargin=0 skip (ParameterSet reset 방지).
    v0.7.9: SetItem 2-tier 복원 (양수 indent silent fail 방지).
    """
    validate_params(params, ["style"], "set_paragraph_style")
    s = params["style"]
    # first_line_indent ↔ indent alias
    if "first_line_indent" in s and "indent" not in s:
        s["indent"] = s["first_line_indent"]

    act = hwp.HAction
    pset = hwp.HParameterSet.HParaShape
    act.GetDefault("ParagraphShape", pset.HSet)
    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
    _need_execute = False

    if "align" in s:
        pset.AlignType = align_map.get(s["align"], 0)
        _need_execute = True
    if "line_spacing" in s:
        pset.LineSpacingType = s.get("line_spacing_type", 0)
        pset.LineSpacing = int(s["line_spacing"])
        _need_execute = True
    if "space_before" in s:
        pset.PrevSpacing = int(s["space_before"] * 100)
        _need_execute = True
    if "space_after" in s:
        pset.NextSpacing = int(s["space_after"] * 100)
        _need_execute = True
    if "page_break_before" in s:
        pset.PagebreakBefore = 1 if s["page_break_before"] else 0
        _need_execute = True
    if "keep_with_next" in s:
        pset.KeepWithNext = 1 if s["keep_with_next"] else 0
        _need_execute = True
    if "widow_orphan" in s:
        pset.WidowOrphan = 1 if s["widow_orphan"] else 0
        _need_execute = True
    # v0.6.7: 8개 추가
    if "line_wrap" in s:
        try:
            pset.LineWrap = int(s["line_wrap"])
            _need_execute = True
        except Exception as e:
            print(f"[WARN] LineWrap: {e}", file=sys.stderr)
    if "snap_to_grid" in s:
        try:
            pset.SnapToGrid = 1 if s["snap_to_grid"] else 0
            _need_execute = True
        except Exception as e:
            print(f"[WARN] SnapToGrid: {e}", file=sys.stderr)
    if "auto_space_eAsian_eng" in s:
        try:
            pset.AutoSpaceEAsianEng = 1 if s["auto_space_eAsian_eng"] else 0
            _need_execute = True
        except Exception as e:
            print(f"[WARN] AutoSpaceEAsianEng: {e}", file=sys.stderr)
    if "auto_space_eAsian_num" in s:
        try:
            pset.AutoSpaceEAsianNum = 1 if s["auto_space_eAsian_num"] else 0
            _need_execute = True
        except Exception as e:
            print(f"[WARN] AutoSpaceEAsianNum: {e}", file=sys.stderr)
    if "break_latin_word" in s:
        try:
            pset.BreakLatinWord = int(s["break_latin_word"])
            _need_execute = True
        except Exception as e:
            print(f"[WARN] BreakLatinWord: {e}", file=sys.stderr)
    if "heading_type" in s:
        try:
            pset.HeadingType = int(s["heading_type"])
            _need_execute = True
        except Exception as e:
            print(f"[WARN] HeadingType: {e}", file=sys.stderr)
    if "keep_lines_together" in s:
        try:
            pset.KeepLinesTogether = 1 if s["keep_lines_together"] else 0
            _need_execute = True
        except Exception as e:
            print(f"[WARN] KeepLinesTogether: {e}", file=sys.stderr)
    if "condense" in s:
        try:
            pset.Condense = int(s["condense"])
            _need_execute = True
        except Exception as e:
            print(f"[WARN] Condense: {e}", file=sys.stderr)

    # v0.6.7: 문단 테두리 4면
    _border_edges = {"left": "Left", "right": "Right", "top": "Top", "bottom": "Bottom"}
    for edge_key, edge_attr in _border_edges.items():
        border_key = f"border_{edge_key}"
        if border_key in s and isinstance(s[border_key], dict):
            bspec = s[border_key]
            try:
                if "type" in bspec:
                    setattr(pset, f"BorderType{edge_attr}", int(bspec["type"]))
                if "width" in bspec:
                    setattr(pset, f"BorderWidth{edge_attr}", float(bspec["width"]))
                if "color" in bspec:
                    c = bspec["color"].lstrip("#")
                    if len(c) == 6:
                        r = int(c[0:2], 16)
                        g = int(c[2:4], 16)
                        b = int(c[4:6], 16)
                        setattr(pset, f"BorderColor{edge_attr}", hwp.RGBColor(r, g, b))
                _need_execute = True
            except Exception as e:
                print(f"[WARN] Border{edge_attr}: {e}", file=sys.stderr)
    if "border_color" in s:
        try:
            c = s["border_color"].lstrip("#")
            if len(c) == 6:
                r = int(c[0:2], 16)
                g = int(c[2:4], 16)
                b = int(c[4:6], 16)
                rgb = hwp.RGBColor(r, g, b)
                for edge_attr in _border_edges.values():
                    setattr(pset, f"BorderColor{edge_attr}", rgb)
                _need_execute = True
        except Exception as e:
            print(f"[WARN] BorderColor: {e}", file=sys.stderr)
    if "border_shadowing" in s:
        try:
            pset.BorderShadowing = 1 if s["border_shadowing"] else 0
            _need_execute = True
        except Exception as e:
            print(f"[WARN] BorderShadowing: {e}", file=sys.stderr)

    # ParaShape 정밀 옵션
    if "first_line_indent_hwpunit" in s:
        try:
            pset.Indentation = int(s["first_line_indent_hwpunit"])
            _need_execute = True
        except Exception as e:
            print(f"[WARN] first_line_indent_hwpunit: {e}", file=sys.stderr)
    if s.get("hanging_indent"):
        try:
            cur_indent = getattr(pset, "Indentation", 0)
            if cur_indent > 0:
                pset.Indentation = -abs(int(cur_indent))
            _need_execute = True
        except Exception as e:
            print(f"[WARN] hanging_indent: {e}", file=sys.stderr)
    if "paragraph_heading_type" in s:
        try:
            pht_map = {"none": 0, "outline": 1, "number": 2}
            pht_val = pht_map.get(s["paragraph_heading_type"], 0)
            try:
                pset.HSet.SetItem("HeadingType", pht_val)
            except Exception:
                pset.HeadingType = pht_val
            _need_execute = True
        except Exception as e:
            print(f"[WARN] paragraph_heading_type: {e}", file=sys.stderr)
    if "word_spacing" in s:
        try:
            ws = int(s["word_spacing"])
            try:
                pset.HSet.SetItem("WordSpacing", ws)
            except Exception:
                pset.WordSpacing = ws
            _need_execute = True
        except Exception as e:
            print(f"[WARN] word_spacing: {e}", file=sys.stderr)
    if "line_weight" in s:
        try:
            lw = int(s["line_weight"])
            try:
                pset.HSet.SetItem("LineWeight", lw)
            except Exception:
                pset.LineWeight = lw
            _need_execute = True
        except Exception as e:
            print(f"[WARN] line_weight: {e}", file=sys.stderr)

    # indent / left_margin / right_margin (HWPUNIT, 양수/음수)
    indent_val = None
    left_margin_val = None
    right_margin_val = None
    if "indent" in s:
        indent_val = float(s["indent"])
    if "left_margin" in s:
        left_margin_val = float(s["left_margin"])
    if "right_margin" in s:
        right_margin_val = float(s["right_margin"])
    # 음수 indent 자동 보정 (v0.6.7 규칙)
    if indent_val is not None and indent_val < 0 and left_margin_val is None:
        left_margin_val = abs(indent_val)

    # v0.7.9 fix: v0.7.3 SetItem 2-tier 복원 (양수 indent silent fail 방지)
    _need_para_execute = False
    if indent_val is not None:
        indent_hwpunit = int(indent_val * 200)
        try:
            try:
                pset.HSet.SetItem("Indent", indent_hwpunit)  # Tier 1: SetItem (더 안정적)
            except Exception:
                pset.Indentation = indent_hwpunit              # Tier 2: 직접 속성
            _need_para_execute = True
        except Exception as e:
            print(f"[WARN] indent={indent_val}: {e}", file=sys.stderr)
    if left_margin_val is not None:
        lm_hwpunit = int(left_margin_val * 200)
        try:
            try:
                pset.HSet.SetItem("LeftMargin", lm_hwpunit)
            except Exception:
                pset.LeftMargin = lm_hwpunit
            _need_para_execute = True
        except Exception as e:
            print(f"[WARN] left_margin={left_margin_val}: {e}", file=sys.stderr)
    if right_margin_val is not None:
        rm_hwpunit = int(right_margin_val * 200)
        try:
            try:
                pset.HSet.SetItem("RightMargin", rm_hwpunit)
            except Exception:
                pset.RightMargin = rm_hwpunit
            _need_para_execute = True
        except Exception as e:
            print(f"[WARN] right_margin={right_margin_val}: {e}", file=sys.stderr)

    if _need_execute or _need_para_execute:
        act.Execute("ParagraphShape", pset.HSet)

    # v0.7.3.4 Fix #2: LeftMargin=0 skip (양수 indent silent fail 방지)
    _para_kwargs = {}
    if left_margin_val is not None and left_margin_val != 0:
        _para_kwargs["LeftMargin"] = left_margin_val
    if right_margin_val is not None and right_margin_val != 0:
        _para_kwargs["RightMargin"] = right_margin_val
    if indent_val is not None:
        _para_kwargs["Indentation"] = indent_val
    if _para_kwargs:
        try:
            hwp.set_para(**_para_kwargs)
        except Exception as e:
            print(f"[INFO] set_para fallback: {e}", file=sys.stderr)
    return {"status": "ok"}
