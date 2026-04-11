"""hwp_core.formatting.styles — 스타일 적용 + 문서 프리셋 + 스타일 프로파일 일괄 적용.

Handlers:
- apply_style             : 스타일 적용 ('제목1', '본문', '개요1' 등)
- apply_document_preset   : 문서 프리셋 (공문서/사업계획서 등)
- apply_style_profile     : 스타일 프로파일 일괄 적용 (body_style 전체)

v0.7.3.5: set_paragraph_style R2 경로 직접 인라인 (recursive dispatch 회피)
v0.7.5.4 P0-2: runAutoFixLoop 에서 더 이상 호출 안 함 (override 부작용 방지)
"""
import sys

from .. import register  # 두 점!
from .._helpers import validate_params  # 두 점!


@register("apply_style")
def apply_style(hwp, params):
    """스타일 적용 — '제목1', '본문', '개요1' 등."""
    style_name = params.get("style_name", "본문")
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HStyle
        act.GetDefault("Style", pset.HSet)
        pset.HSet.SetItem("StyleName", style_name)
        act.Execute("Style", pset.HSet)
        return {"status": "ok", "style": style_name}
    except Exception as e:
        raise RuntimeError(f"스타일 적용 실패: {e}")


@register("apply_document_preset")
def apply_document_preset(hwp, params):
    """문서 프리셋 적용 — 공문서/사업계획서 등."""
    validate_params(params, ["preset_name"], "apply_document_preset")
    from presets import DOCUMENT_PRESETS
    preset_name = params["preset_name"]
    if preset_name not in DOCUMENT_PRESETS:
        return {"error": f"프리셋 '{preset_name}' 없음. 사용 가능: {list(DOCUMENT_PRESETS.keys())}"}
    preset = DOCUMENT_PRESETS[preset_name]
    # 1. 용지 설정 (REGISTRY lookup)
    page = preset.get("page", {})
    if page:
        from hwp_core import REGISTRY
        page_handler = REGISTRY.get("set_page_setup")
        if page_handler:
            page_handler(hwp, {
                "top_margin": page.get("top", 20),
                "bottom_margin": page.get("bottom", 15),
                "left_margin": page.get("left", 20),
                "right_margin": page.get("right", 20),
            })
    # 2. 본문 서식
    body = preset.get("body", {})
    if body:
        from hwp_editor import set_paragraph_style as _set_ps
        para_params = {}
        if "line_spacing" in body:
            para_params["line_spacing"] = body["line_spacing"]
        if "align" in body:
            para_params["align"] = body["align"]
        if para_params:
            _set_ps(hwp, para_params)
    return {"status": "ok", "preset": preset_name, "applied": preset}


@register("apply_style_profile")
def apply_style_profile(hwp, params):
    """스타일 프로파일 일괄 적용 (v0.7.1).

    v0.7.3.5: set_paragraph_style R2 경로 직접 인라인 (recursive dispatch 는 empty-doc silent fail).
    v0.7.5.4 P0-2: runAutoFixLoop 에서 더 이상 호출 안 함 (override 부작용 방지).
    """
    validate_params(params, ["profile"], "apply_style_profile")
    profile = params["profile"]
    target = params.get("target", "all")

    body = profile.get("body_style", {}) if isinstance(profile, dict) else {}
    applied_para = 0
    applied_char = 0

    # === ParaShape 적용 ===
    para = body.get("para", {}) if isinstance(body, dict) else {}
    if para:
        try:
            _p = {}
            if "align" in para:
                _p["align"] = para["align"]
            if "line_spacing" in para:
                _p["line_spacing"] = para["line_spacing"]
            if "space_before" in para:
                _p["space_before"] = para["space_before"]
            if "space_after" in para:
                _p["space_after"] = para["space_after"]
            if "left_margin" in para:
                _p["left_margin"] = para["left_margin"]
            if "right_margin" in para:
                _p["right_margin"] = para["right_margin"]
            if "first_line_indent" in para:
                _p["first_line_indent"] = para["first_line_indent"]
            elif "indent" in para:
                _p["indent"] = para["indent"]
            if "AlignType" in para:
                _p["align"] = ["left", "center", "right", "justify"][int(para["AlignType"])] if 0 <= int(para["AlignType"]) <= 3 else "left"
            if "LineSpacing" in para:
                _p["line_spacing"] = int(para["LineSpacing"])

            if _p:
                _s = _p
                if "first_line_indent" in _s and "indent" not in _s:
                    _s["indent"] = _s["first_line_indent"]
                _act = hwp.HAction
                _pset = hwp.HParameterSet.HParaShape
                _act.GetDefault("ParagraphShape", _pset.HSet)
                _align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                _need_exec = False
                if "align" in _s:
                    _pset.AlignType = _align_map.get(_s["align"], 0)
                    _need_exec = True
                if "line_spacing" in _s:
                    _pset.LineSpacingType = 0
                    _pset.LineSpacing = int(_s["line_spacing"])
                    _need_exec = True
                if "space_before" in _s:
                    _pset.PrevSpacing = int(float(_s["space_before"]) * 100)
                    _need_exec = True
                if "space_after" in _s:
                    _pset.NextSpacing = int(float(_s["space_after"]) * 100)
                    _need_exec = True
                _ind = float(_s["indent"]) if "indent" in _s else None
                _lm = float(_s["left_margin"]) if "left_margin" in _s else None
                _rm = float(_s["right_margin"]) if "right_margin" in _s else None
                if _ind is not None and _ind < 0 and _lm is None:
                    _lm = abs(_ind)
                if _ind is not None:
                    _pset.Indentation = int(_ind * 200)
                    _need_exec = True
                if _lm is not None:
                    _pset.LeftMargin = int(_lm * 200)
                    _need_exec = True
                if _rm is not None:
                    _pset.RightMargin = int(_rm * 200)
                    _need_exec = True
                if _need_exec:
                    _act.Execute("ParagraphShape", _pset.HSet)
                _kw = {}
                if _lm is not None and _lm != 0:
                    _kw["LeftMargin"] = _lm
                if _rm is not None and _rm != 0:
                    _kw["RightMargin"] = _rm
                if _ind is not None:
                    _kw["Indentation"] = _ind
                if _kw:
                    try:
                        hwp.set_para(**_kw)
                    except Exception as e:
                        print(f"[INFO] set_para fallback: {e}", file=sys.stderr)
                applied_para += 1
        except Exception as e:
            print(f"[WARN] apply_style_profile para: {e}", file=sys.stderr)

    # === CharShape 적용 ===
    char = body.get("char", {}) if isinstance(body, dict) else {}
    if char:
        try:
            act = hwp.HAction
            cset = hwp.HParameterSet.HCharShape
            act.GetDefault("CharShape", cset.HSet)
            if "font_size" in char:
                try:
                    cset.Height = int(float(char["font_size"]) * 100)
                except Exception:
                    pass
            if "bold" in char:
                try:
                    cset.Bold = 1 if char["bold"] else 0
                except Exception:
                    pass
            if "italic" in char:
                try:
                    cset.Italic = 1 if char["italic"] else 0
                except Exception:
                    pass
            if "color" in char and isinstance(char.get("color"), str) and char["color"].startswith("#"):
                try:
                    c = char["color"]
                    r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
                    cset.TextColor = hwp.RGBColor(r, g, b)
                except Exception:
                    pass
            act.Execute("CharShape", cset.HSet)
            applied_char += 1
        except Exception as e:
            print(f"[WARN] apply_style_profile char: {e}", file=sys.stderr)

    return {
        "status": "ok",
        "applied_paragraphs": applied_para,
        "applied_chars": applied_char,
        "target": target,
    }
