"""hwp_editor.text_style — 텍스트 삽입 + 문자/단락 서식 설정.

함수:
- insert_text_with_color  : 색상 지정 텍스트 삽입 (간단)
- insert_text_with_style  : 20+ 속성 지원 서식 텍스트 삽입 (bold/italic/color/font/size/underline/...)
- set_paragraph_style     : 단락 서식 설정 (align/line_spacing/margins/indent/...)

v0.6.7: FaceNameOther 추가, ParaShape 전체 8+ 필드 sync
"""
import sys


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
        "color": [r,g,b],          # 글자 색상
        "bold": True/False,         # 굵게
        "italic": True/False,       # 기울임
        "underline": True/False,    # 밑줄 (bool → 실선)
        "underline_type": 0-7,      # 밑줄 종류 (0=없음,1=실선,2=이중,3=점선,4=파선,5=1점쇄선,6=물결,7=굵은실선)
        "underline_color": [r,g,b], # 밑줄 색상
        "font_size": 12.0,          # 글자 크기 (pt)
        "font_name": "맑은 고딕",   # 글꼴 (한글+라틴 동시)
        "font_name_latin": "Arial", # 라틴 전용 글꼴
        "bg_color": [r,g,b],        # 배경 색상
        "strikeout": True/False,    # 취소선 (bool → 단일)
        "strikeout_type": 0-3,      # 취소선 종류 (0=없음,1=단일,2=이중,3=굵은)
        "strikeout_color": [r,g,b], # 취소선 색상
        "char_spacing": -5,         # 자간 (%, 기본 0)
        "width_ratio": 90,          # 장평 (%, 기본 100)
        "font_name_hanja": "바탕",  # 한자 글꼴
        "font_name_japanese": "",   # 일본어 글꼴
        "font_name_other": "",      # 기타(라틴 외) 글꼴 (v0.6.7)
        "superscript": True/False,  # 위 첨자
        "subscript": True/False,    # 아래 첨자
        "outline": True/False,      # 외곽선
        "shadow": True/False,       # 그림자
        "emboss": True/False,       # 양각
        "engrave": True/False,      # 음각
        "small_caps": True/False,   # 작은 대문자
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
    saved = {}
    saved['TextColor'] = pset.TextColor
    saved['Bold'] = pset.Bold
    saved['Italic'] = pset.Italic
    saved['UnderlineType'] = pset.UnderlineType
    saved['Height'] = pset.Height
    saved['StrikeOutType'] = pset.StrikeOutType
    for attr in ['SpacingHangul', 'RatioHangul', 'SuperScript', 'SubScript',
                 'OutLineType', 'ShadowType', 'Emboss', 'Engrave', 'SmallCaps',
                 'UnderlineColor', 'StrikeOutColor', 'ShadeColor']:
        try:
            saved[attr] = getattr(pset, attr)
        except Exception:
            saved[attr] = None

    # 새 서식 적용
    act.GetDefault("CharShape", pset.HSet)

    if "color" in style:
        c = style["color"]
        pset.TextColor = hwp.RGBColor(c[0], c[1], c[2])
    if "bold" in style:
        pset.Bold = 1 if style["bold"] else 0
    if "italic" in style:
        pset.Italic = 1 if style["italic"] else 0
    if "underline_type" in style:
        pset.UnderlineType = int(style["underline_type"])
    elif "underline" in style:
        pset.UnderlineType = 1 if style["underline"] else 0
    if "underline_color" in style:
        uc = style["underline_color"]
        try:
            pset.UnderlineColor = hwp.RGBColor(uc[0], uc[1], uc[2])
        except Exception as e:
            print(f"[WARN] UnderlineColor: {e}", file=sys.stderr)
    if "font_size" in style:
        pset.Height = int(style["font_size"] * 100)  # pt → HWP 단위
    if "font_name" in style:
        pset.FaceNameHangul = style["font_name"]
        pset.FaceNameLatin = style["font_name"]
    if "font_name_latin" in style:
        pset.FaceNameLatin = style["font_name_latin"]
    if "bg_color" in style:
        bg = style["bg_color"]
        pset.ShadeColor = hwp.RGBColor(bg[0], bg[1], bg[2])
    if "strikeout_type" in style:
        pset.StrikeOutType = int(style["strikeout_type"])
    elif "strikeout" in style:
        pset.StrikeOutType = 1 if style["strikeout"] else 0
    if "strikeout_color" in style:
        sc = style["strikeout_color"]
        try:
            pset.StrikeOutColor = hwp.RGBColor(sc[0], sc[1], sc[2])
        except Exception as e:
            print(f"[WARN] StrikeOutColor: {e}", file=sys.stderr)
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
    if "font_name_other" in style:
        # v0.6.7: 기타(라틴 외) 글꼴 — FaceNameOther 1건 추가
        try:
            pset.FaceNameOther = style["font_name_other"]
        except Exception as e:
            print(f"[WARN] FaceNameOther: {e}", file=sys.stderr)
    # 위/아래 첨자
    if "superscript" in style:
        try:
            pset.SuperScript = 1 if style["superscript"] else 0
        except Exception as e:
            print(f"[WARN] SuperScript: {e}", file=sys.stderr)
    if "subscript" in style:
        try:
            pset.SubScript = 1 if style["subscript"] else 0
        except Exception as e:
            print(f"[WARN] SubScript: {e}", file=sys.stderr)
    # 외곽선/그림자/양각/음각/작은대문자
    if "outline" in style:
        try:
            pset.OutLineType = 1 if style["outline"] else 0
        except Exception as e:
            print(f"[WARN] OutLineType: {e}", file=sys.stderr)
    if "shadow" in style:
        try:
            pset.ShadowType = 1 if style["shadow"] else 0
        except Exception as e:
            print(f"[WARN] ShadowType: {e}", file=sys.stderr)
    if "emboss" in style:
        try:
            pset.Emboss = 1 if style["emboss"] else 0
        except Exception as e:
            print(f"[WARN] Emboss: {e}", file=sys.stderr)
    if "engrave" in style:
        try:
            pset.Engrave = 1 if style["engrave"] else 0
        except Exception as e:
            print(f"[WARN] Engrave: {e}", file=sys.stderr)
    if "small_caps" in style:
        try:
            pset.SmallCaps = 1 if style["small_caps"] else 0
        except Exception as e:
            print(f"[WARN] SmallCaps: {e}", file=sys.stderr)
    # 그림자 색상/오프셋
    if "shadow_color" in style:
        try:
            sc = style["shadow_color"]
            pset.ShadowColor = hwp.RGBColor(sc[0], sc[1], sc[2])
        except Exception as e:
            print(f"[WARN] ShadowColor: {e}", file=sys.stderr)
    if "shadow_offset_x" in style:
        try:
            pset.ShadowOffsetX = int(style["shadow_offset_x"])
        except Exception as e:
            print(f"[WARN] ShadowOffsetX: {e}", file=sys.stderr)
    if "shadow_offset_y" in style:
        try:
            pset.ShadowOffsetY = int(style["shadow_offset_y"])
        except Exception as e:
            print(f"[WARN] ShadowOffsetY: {e}", file=sys.stderr)
    # 밑줄/취소선 모양
    if "underline_shape" in style:
        try:
            pset.UnderlineShape = int(style["underline_shape"])
        except Exception as e:
            print(f"[WARN] UnderlineShape: {e}", file=sys.stderr)
    if "strikeout_shape" in style:
        try:
            pset.StrikeOutShape = int(style["strikeout_shape"])
        except Exception as e:
            print(f"[WARN] StrikeOutShape: {e}", file=sys.stderr)
    # 커닝
    if "use_kerning" in style:
        try:
            pset.UseKerning = 1 if style["use_kerning"] else 0
        except Exception as e:
            print(f"[WARN] UseKerning: {e}", file=sys.stderr)

    act.Execute("CharShape", pset.HSet)

    # B1 (v0.6.6): 외곽 try/finally — insert_text 예외 시에도 CharShape 복원 보장
    try:
        hwp.insert_text(text)
    finally:
        # 원래 서식 복원 (예외 시에도 반드시 실행)
        try:
            act.GetDefault("CharShape", pset.HSet)
            pset.TextColor = saved['TextColor']
            pset.Bold = saved['Bold']
            pset.Italic = saved['Italic']
            pset.UnderlineType = saved['UnderlineType']
            pset.Height = saved['Height']
            pset.StrikeOutType = saved['StrikeOutType']
            for attr in ['SpacingHangul', 'RatioHangul', 'SuperScript', 'SubScript',
                         'OutLineType', 'ShadowType', 'Emboss', 'Engrave', 'SmallCaps',
                         'UnderlineColor', 'StrikeOutColor', 'ShadeColor']:
                if saved.get(attr) is not None:
                    try:
                        setattr(pset, attr, saved[attr])
                        if attr == 'SpacingHangul':
                            pset.SpacingLatin = saved[attr]
                        if attr == 'RatioHangul':
                            pset.RatioLatin = saved[attr]
                    except Exception as e:
                        print(f"[WARN] Restore {attr}: {e}", file=sys.stderr)
            act.Execute("CharShape", pset.HSet)
        except Exception as e:
            print(f"[WARN] CharShape restore failed: {e}", file=sys.stderr)


def set_paragraph_style(hwp, style=None):
    """현재 커서 위치의 단락 서식을 설정.
    v0.7.5.0 Issue 5: 반환 {ok, applied: [field,...], failed: [{field, reason},...]}
    style: {
        "align": "left"|"center"|"right"|"justify",  # 정렬
        "line_spacing": 160,    # 줄간격 (%)
        "space_before": 0,      # 문단 앞 간격 (pt)
        "space_after": 0,       # 문단 뒤 간격 (pt)
        "indent": 0,            # 들여쓰기 (pt)
    }
    """
    _applied: list = []
    _failed: list = []

    def _try(field, fn):
        try:
            fn()
            _applied.append(field)
        except Exception as e:
            _failed.append({"field": field, "reason": str(e)})
            print(f"[WARN] {field}: {e}", file=sys.stderr)

    if not style:
        return {"ok": True, "applied": [], "failed": []}

    act = hwp.HAction
    pset = hwp.HParameterSet.HParaShape

    act.GetDefault("ParagraphShape", pset.HSet)

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
    # 문단 앞 페이지 나누기
    if "page_break_before" in style:
        try:
            pset.PagebreakBefore = 1 if style["page_break_before"] else 0
        except Exception as e:
            print(f"[WARN] PagebreakBefore: {e}", file=sys.stderr)
    # 다음 문단과 함께
    if "keep_with_next" in style:
        try:
            pset.KeepWithNext = 1 if style["keep_with_next"] else 0
        except Exception as e:
            print(f"[WARN] KeepWithNext: {e}", file=sys.stderr)
    # 과부/고아 방지
    if "widow_orphan" in style:
        try:
            pset.WidowOrphan = 1 if style["widow_orphan"] else 0
        except Exception as e:
            print(f"[WARN] WidowOrphan: {e}", file=sys.stderr)
    # 줄 바꿈
    if "line_wrap" in style:
        try:
            pset.LineWrap = int(style["line_wrap"])
        except Exception as e:
            print(f"[WARN] LineWrap: {e}", file=sys.stderr)
    # 그리드에 맞춤
    if "snap_to_grid" in style:
        try:
            pset.SnapToGrid = 1 if style["snap_to_grid"] else 0
        except Exception as e:
            print(f"[WARN] SnapToGrid: {e}", file=sys.stderr)
    # 한영 자동 간격
    if "auto_space_eAsian_eng" in style:
        try:
            pset.AutoSpaceEAsianEng = 1 if style["auto_space_eAsian_eng"] else 0
        except Exception as e:
            print(f"[WARN] AutoSpaceEAsianEng: {e}", file=sys.stderr)
    if "auto_space_eAsian_num" in style:
        try:
            pset.AutoSpaceEAsianNum = 1 if style["auto_space_eAsian_num"] else 0
        except Exception as e:
            print(f"[WARN] AutoSpaceEAsianNum: {e}", file=sys.stderr)
    # 영문 줄바꿈
    if "break_latin_word" in style:
        try:
            pset.BreakLatinWord = int(style["break_latin_word"])
        except Exception as e:
            print(f"[WARN] BreakLatinWord: {e}", file=sys.stderr)
    # 제목 수준
    if "heading_type" in style:
        try:
            pset.HeadingType = int(style["heading_type"])
        except Exception as e:
            print(f"[WARN] HeadingType: {e}", file=sys.stderr)
    # 줄 함께 유지
    if "keep_lines_together" in style:
        try:
            pset.KeepLinesTogether = 1 if style["keep_lines_together"] else 0
        except Exception as e:
            print(f"[WARN] KeepLinesTogether: {e}", file=sys.stderr)
    # 문단 압축
    if "condense" in style:
        try:
            pset.Condense = int(style["condense"])
        except Exception as e:
            print(f"[WARN] Condense: {e}", file=sys.stderr)

    try:
        _exec_ok = bool(act.Execute("ParagraphShape", pset.HSet))
    except Exception as e:
        _failed.append({"field": "Execute", "reason": str(e)})
        _exec_ok = False
    return {"ok": _exec_ok and not _failed, "applied": _applied, "failed": _failed}
