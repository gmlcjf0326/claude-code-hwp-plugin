"""hwp_core.text_editing.insertions — 텍스트/제목/본문 삽입 handlers.

Handlers:
- insert_text                : 텍스트 삽입 + auto_indent + outline_level
- insert_heading             : 제목 삽입 + numbering + OutlineLevel (v0.6.9)
- insert_body_after_heading  : 소제목 아래 본문 (TOC skip + heading inherit)
                               v0.7.7: 3-tier fuzzy matcher
                               v0.7.9: set_style 바탕글 + 뎁스별 자동 들여쓰기
                               메가 함수 354L — Phase 10 defer
- find_and_append            : 찾아서 뒤에 추가 (mode='body' 시 body 핸들러 위임)
- extend_section             : 섹션 확장 — 제목 찾고 직후 텍스트 삽입
"""
import re
import sys

from .. import register  # 두 점!
from .._helpers import validate_params, _execute_all_replace, _exit_table_safely  # 두 점!
from ._internal import (
    _with_auto_dismiss,
    _find_heading_positions,
    _apply_indent_at_caret,
    _detect_heading_depth,
)


@register("insert_text")
def insert_text(hwp, params):
    """텍스트 삽입 + auto_indent + outline_level."""
    validate_params(params, ["text"], "insert_text")
    # 표 안에 커서가 있으면 먼저 탈출
    try:
        if hwp.is_cell():
            _exit_table_safely(hwp)
    except Exception:
        pass
    text = params["text"]
    # 전처리: 줄바꿈 정규화 + 마커 앞 자동 줄바꿈
    text = text.replace("\r\n", "\n")
    _markers = r'[○□■◆●•◦※➤❶-❿▶▷►]'
    _roman = r'(?:Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)'
    text = re.sub(rf'(?<=[^\n])({_markers})', r'\n\1', text)
    text = re.sub(rf'(?<=[^\n])({_roman}\.)', r'\n\1', text)
    text = re.sub(r'  {3,}', '\n     ', text)
    text = text.replace("\n", "\r\n")
    if not text.endswith("\r\n"):
        text += "\r\n"

    original_text = params["text"]
    style = params.get("style")
    color = params.get("color")  # [r, g, b] 하위 호환
    if style:
        from hwp_editor import insert_text_with_style
        insert_text_with_style(hwp, text, style)
    elif color:
        from hwp_editor import insert_text_with_color
        insert_text_with_color(hwp, text, tuple(color))
    else:
        hwp.insert_text(text)

    # 후처리: 마커/번호 → ParagraphShapeIndentAtCaret 자동 내어쓰기
    auto_indent = params.get("auto_indent", True)
    outline_level = params.get("outline_level")
    if outline_level is None and auto_indent:
        _apply_indent_at_caret(hwp, original_text)

    # v0.6.9: outline_level 지정 시 직전 단락 OutlineLevel 설정
    if outline_level is not None:
        try:
            hwp.HAction.Run("MovePrevPara")
            ol_int = int(outline_level)
            success = False
            # 시도 1: ParameterSet.HSet.SetItem
            try:
                act = hwp.HAction
                pset = hwp.HParameterSet.HParaShape
                act.GetDefault("ParagraphShape", pset.HSet)
                pset.HSet.SetItem("OutlineLevel", ol_int)
                act.Execute("ParagraphShape", pset.HSet)
                success = True
            except Exception as e1:
                print(f"[INFO] insert_text OutlineLevel SetItem failed: {e1}", file=sys.stderr)
            # 시도 2: hwp.set_style("개요 N+1")
            if not success:
                try:
                    hwp.set_style(f"개요 {ol_int + 1}")
                    success = True
                except Exception as e2:
                    print(f"[INFO] insert_text set_style 개요 {ol_int + 1} failed: {e2}", file=sys.stderr)
            # 시도 3: pset.OutlineLevel 직접 attribute
            if not success:
                try:
                    act = hwp.HAction
                    pset = hwp.HParameterSet.HParaShape
                    act.GetDefault("ParagraphShape", pset.HSet)
                    pset.OutlineLevel = ol_int
                    act.Execute("ParagraphShape", pset.HSet)
                    success = True
                except Exception as e3:
                    print(f"[WARN] insert_text OutlineLevel all alternatives failed: {e3}", file=sys.stderr)
            hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] insert_text OutlineLevel (level={outline_level}): {e}", file=sys.stderr)
    return {"status": "ok"}


@register("insert_heading")
def insert_heading(hwp, params):
    """제목 삽입 + numbering + OutlineLevel (v0.6.9)."""
    validate_params(params, ["text", "level"], "insert_heading")
    from hwp_editor import insert_text_with_style
    level = min(max(params["level"], 1), 9)
    sizes = {1: 22, 2: 18, 3: 15, 4: 13, 5: 11, 6: 10, 7: 10, 8: 10, 9: 10}
    text = params["text"]
    numbering = params.get("numbering")
    number = params.get("number", 1)
    auto_outline_level = bool(params.get("auto_outline_level", False))
    outline_level_only = bool(params.get("outline_level_only", False))

    if numbering and not outline_level_only:
        roman = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"]
        korean = ["가", "나", "다", "라", "마", "바", "사", "아", "자", "차"]
        circle = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
        idx = max(0, min(number - 1, 9))
        if numbering == "roman":
            text = f"{roman[idx]}. {text}"
        elif numbering == "decimal":
            text = f"{number}. {text}"
        elif numbering == "korean":
            text = f"{korean[idx]}. {text}"
        elif numbering == "circle":
            text = f"{circle[idx]} {text}"
        elif numbering == "paren_decimal":
            text = f"{number}) {text}"
        elif numbering == "paren_korean":
            text = f"{korean[idx]}) {text}"

    insert_text_with_style(hwp, text + "\r\n", {
        "bold": True,
        "font_size": sizes.get(level, 11),
    })

    applied_outline_level = None
    applied_via = None
    if auto_outline_level or outline_level_only:
        try:
            hwp.HAction.Run("MovePrevPara")
            ol_int = level - 1
            # 시도 1: SetItem
            try:
                act = hwp.HAction
                pset = hwp.HParameterSet.HParaShape
                act.GetDefault("ParagraphShape", pset.HSet)
                pset.HSet.SetItem("OutlineLevel", ol_int)
                act.Execute("ParagraphShape", pset.HSet)
                applied_outline_level = ol_int
                applied_via = "SetItem"
            except Exception as e1:
                print(f"[INFO] insert_heading SetItem failed: {e1}", file=sys.stderr)
            # 시도 2: set_style
            if applied_outline_level is None:
                try:
                    hwp.set_style(f"개요 {level}")
                    applied_outline_level = ol_int
                    applied_via = "set_style"
                except Exception as e2:
                    print(f"[INFO] insert_heading set_style failed: {e2}", file=sys.stderr)
            # 시도 3: direct attribute
            if applied_outline_level is None:
                try:
                    act = hwp.HAction
                    pset = hwp.HParameterSet.HParaShape
                    act.GetDefault("ParagraphShape", pset.HSet)
                    pset.OutlineLevel = ol_int
                    act.Execute("ParagraphShape", pset.HSet)
                    applied_outline_level = ol_int
                    applied_via = "direct_attribute"
                except Exception as e3:
                    print(f"[WARN] insert_heading OutlineLevel failed: {e3}", file=sys.stderr)
            hwp.MovePos(3)
        except Exception as e:
            print(f"[WARN] insert_heading OutlineLevel: {e}", file=sys.stderr)
    return {
        "status": "ok",
        "level": level,
        "text": text,
        "outline_level": applied_outline_level,
        "applied_via": applied_via,
    }


@register("find_and_append")
def find_and_append(hwp, params):
    """찾아서 뒤에 텍스트 추가. mode='body' 시 insert_body_after_heading 위임."""
    validate_params(params, ["find", "append_text"], "find_and_append")
    find_text = params["find"]
    append_text = params["append_text"]

    # v0.7.5.1: mode="body" 시 insert_body_after_heading 로 위임
    mode = params.get("mode", "all")
    if mode == "body":
        # REGISTRY lookup (dispatch 순환 회피)
        from hwp_core import REGISTRY
        body_handler = REGISTRY.get("insert_body_after_heading")
        if body_handler:
            body_params = dict(params)
            body_params["heading"] = find_text
            body_params["body_text"] = append_text
            return body_handler(hwp, body_params)

    # v0.7.5.1: false not_found 제거 — count 기반
    before = ""
    try:
        before = hwp.get_text_file("TEXT", "")
    except Exception:
        pass

    before_count = before.count(find_text) if find_text else 0
    if before_count == 0:
        return {"status": "not_found", "find": find_text}

    replace_text = find_text + append_text
    _execute_all_replace(hwp, find_text, replace_text)

    after = ""
    try:
        after = hwp.get_text_file("TEXT", "")
    except Exception:
        pass

    after_replace_count = after.count(replace_text) if replace_text else 0
    if after_replace_count >= 1:
        return {
            "status": "ok",
            "find": find_text,
            "appended": True,
            "occurrences": after_replace_count,
            "toc_polluted": after_replace_count > 1,
            "hint": "목차 오염 가능 — mode='body' 사용 권장" if after_replace_count > 1 else "",
        }
    else:
        return {
            "status": "not_found",
            "find": find_text,
            "warning": "AllReplace 실행했으나 텍스트 변화 미확인",
        }


@register("insert_body_after_heading")
def insert_body_after_heading(hwp, params):
    """소제목 아래 본문 단락 삽입 (TOC skip + 문단 분리 + 본문 스타일 inherit).

    v0.7.5.1: TOC skip
    v0.7.5.3: heading inherit (char/para)
    v0.7.5.4: 양수 indent 제거 (pyhwpx silent fail 우회)
    v0.7.7: 3-tier fuzzy heading matcher (정규화/공백유연/핵심어)
    v0.7.9: set_style 바탕글 + 뎁스별 자동 들여쓰기

    354L 메가 함수 — Phase 10 에서 내부 helper 로 refactor 예정.
    """
    validate_params(params, ["heading", "body_text"], "insert_body_after_heading")
    heading = params["heading"]
    body_text = params["body_text"]
    body_style_opt = params.get("body_style", {}) or {}
    skip_toc = bool(params.get("skip_toc", True))
    occurrence = int(params.get("occurrence", -1))  # -1 = auto

    # 1) 전체 텍스트에서 fuzzy heading search
    full_text = ""
    try:
        full_text = hwp.get_text_file("TEXT", "")
    except Exception as e:
        print(f"[WARN] get_text_file: {e}", file=sys.stderr)

    # v0.7.7: 3-tier fuzzy matching
    all_positions = _find_heading_positions(full_text, heading)
    total_matches = len(all_positions)
    match_tier = all_positions[0][2] if all_positions else 0

    if total_matches == 0:
        return {"status": "not_found", "heading": heading, "match_tier": 0}

    # 실제 매칭된 텍스트 (RepeatFind 에 사용)
    find_string = all_positions[0][1]  # 기본값: 첫 매칭 텍스트

    # TOC suffix 패턴 감지 (\t\d+, ...\d, 공백+\d 등)
    toc_pattern = re.compile(r"\t\s*\d+|\s*\.{3,}\s*\d*|\t\s*$|\s+\d{1,3}\s*$")
    match_positions = []
    body_match_positions = []
    toc_match_positions = []
    for char_pos, matched_text, tier in all_positions:
        tail_end = min(len(full_text), char_pos + len(matched_text) + 30)
        tail = full_text[char_pos + len(matched_text):tail_end]
        nl = tail.find("\n")
        cr = tail.find("\r")
        line_end = min(x for x in [nl, cr, len(tail)] if x >= 0) if (nl >= 0 or cr >= 0) else len(tail)
        line_tail = tail[:line_end]
        if toc_pattern.search(line_tail):
            toc_match_positions.append((char_pos, matched_text))
        else:
            body_match_positions.append((char_pos, matched_text))
        match_positions.append((char_pos, matched_text))

    if skip_toc and not body_match_positions:
        return {
            "status": "toc_only",
            "heading": heading,
            "total_matches": total_matches,
            "match_tier": match_tier,
            "toc_matches": len(toc_match_positions),
            "body_matches": 0,
            "hint": "본문에 해당 소제목이 없음 — 원본 양식에 placeholder 부재 또는 TOC 항목만 존재",
        }

    # target_occurrence 계산 + find_string 결정
    if occurrence >= 1:
        target_occurrence = occurrence
        if occurrence <= len(match_positions):
            find_string = match_positions[occurrence - 1][1]
    elif skip_toc:
        first_body = body_match_positions[0]
        target_occurrence = next(
            i + 1 for i, (p, _) in enumerate(match_positions) if p == first_body[0]
        )
        find_string = first_body[1]
    else:
        target_occurrence = 1
        find_string = match_positions[0][1]

    # 2) MoveDocBegin + RepeatFind N 번
    try:
        hwp.HAction.Run("MoveDocBegin")
    except Exception:
        pass
    act = hwp.HAction
    pset = hwp.HParameterSet.HFindReplace
    act.GetDefault("RepeatFind", pset.HSet)
    pset.FindString = find_string  # v0.7.7: 실제 매칭된 텍스트 사용
    pset.Direction = 0
    pset.IgnoreMessage = 1

    found_count = 0
    for _ in range(target_occurrence):
        try:
            ok = act.Execute("RepeatFind", pset.HSet)
        except Exception as e:
            print(f"[WARN] RepeatFind: {e}", file=sys.stderr)
            ok = False
        if not ok:
            break
        found_count += 1

    if found_count < target_occurrence:
        return {
            "status": "not_found",
            "heading": heading,
            "match_tier": match_tier,
            "matched_text": find_string,
            "total_matches": total_matches,
            "required_occurrence": target_occurrence,
            "found_count": found_count,
        }

    # heading 의 char/para shape 캡처 (v0.7.5.3)
    heading_char_shape = {}
    heading_para_shape = {}
    try:
        from hwp_editor import get_char_shape as _get_cs, get_para_shape as _get_ps
        heading_char_shape = _get_cs(hwp) or {}
        heading_para_shape = _get_ps(hwp) or {}
    except Exception as e:
        print(f"[WARN] heading shape capture: {e}", file=sys.stderr)

    # 3) 제목 줄 끝 → 다음 빈 줄 재활용 또는 BreakPara
    # v0.7.9 핵심 fix: 양식의 빈 줄 placeholder 를 재활용
    try:
        hwp.HAction.Run("MoveLineEnd")
    except Exception as e:
        print(f"[WARN] MoveLineEnd: {e}", file=sys.stderr)

    # 다음 줄이 빈 줄인지 체크 (빈 줄이면 그 위치로 이동, 아니면 BreakPara)
    insertion_mode = "break_para"
    try:
        saved_pos = hwp.GetPos()
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveLineBegin")
        ln_start = hwp.GetPos()
        hwp.HAction.Run("MoveLineEnd")
        ln_end = hwp.GetPos()
        if ln_start[1] == ln_end[1] and ln_start[2] == ln_end[2]:
            # 빈 줄 — 그 시작 위치로 복귀
            hwp.SetPos(*ln_start)
            insertion_mode = "fill_empty_line"
        else:
            # 빈 줄 아님 → 원위치 복귀 후 BreakPara
            hwp.SetPos(*saved_pos)
            hwp.HAction.Run("BreakPara")
            insertion_mode = "break_para"
    except Exception as e:
        print(f"[WARN] empty line detection: {e}", file=sys.stderr)
        try:
            hwp.HAction.Run("BreakPara")
        except Exception:
            pass

    # v0.7.9: set_style("바탕글") 복원 + 대화상자 백그라운드 자동 "덮어씀(Y)" 클릭
    try:
        _with_auto_dismiss(hwp, hwp.set_style, "바탕글")
    except Exception as e:
        print(f"[WARN] set_style 바탕글: {e}", file=sys.stderr)
        # fallback: 정렬만 리셋
        try:
            hwp.set_para(AlignType='Left')
        except Exception:
            pass

    # v0.7.9-postfix5 (P3, 일반화 첫걸음): CharShape bold=0/italic=0/underline=0 강제.
    # 양식 placeholder paragraph 가 heading 상속 볼드 CharShape 인 경우 fill_empty_line 시
    # 그 볼드를 본문에 상속 → 이미지 #9 의 `+|의료` 같은 partial bold 버그 발생.
    # 본문 insert 전에 CharShape 강제 재설정하여 양식 독립적 해결.
    # font_size 는 heading 상속 유지 (일부러 설정 안 함).
    try:
        pset = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        pset.Bold = 0
        pset.Italic = 0
        pset.UnderLine = 0
        hwp.HAction.Execute("CharShape", pset.HSet)
    except Exception as e:
        print(f"[WARN] CharShape bold=0 force: {e}", file=sys.stderr)

    # v0.7.9: 4) 본문 스타일 — 뎁스별 들여쓰기 자동 계산
    _match_line = find_string
    if all_positions:
        _mp = all_positions[0][0]  # 첫 매칭의 char_pos
        _line_start = full_text.rfind('\n', 0, _mp) + 1
        _line_end = full_text.find('\n', _mp)
        if _line_end < 0:
            _line_end = len(full_text)
        _match_line = full_text[_line_start:_line_end].strip()
    heading_depth = _detect_heading_depth(_match_line)

    # 뎁스별 left_margin (나머지줄 시작위치) 자동 계산
    INDENT_PER_DEPTH = 10  # pt per depth level
    FIRST_LINE_INDENT = 10  # 첫줄 추가 들여쓰기

    auto_left_margin = max(0, (heading_depth - 1) * INDENT_PER_DEPTH)
    auto_indent = FIRST_LINE_INDENT

    # v0.7.9 리팩터: 양식 서식 최우선 (프리셋은 fallback만)
    heading_align = heading_para_shape.get("align", "left")
    if heading_align not in ("left", "justify"):
        heading_align = "left"

    # inherited_para — 양식 heading의 para 값 전체 상속
    inherited_para = {
        "left_margin": auto_left_margin,
        "indent": auto_indent,
        "align": heading_align,
        "line_spacing": heading_para_shape.get("line_spacing", 160),
        "line_spacing_type": heading_para_shape.get("line_spacing_type", "percent"),
        "space_before": heading_para_shape.get("space_before", 0),
        "space_after": heading_para_shape.get("space_after", 0),
        "snap_to_grid": heading_para_shape.get("snap_to_grid", False),
    }

    # 양식 > 프리셋 (프리셋은 양식에 없는 값만 fallback)
    user_para = body_style_opt.get("para", {}) or {}
    for k, v in user_para.items():
        if k in ("left_margin", "indent"):
            continue  # 뎁스 자동 계산값 보호
        if k not in inherited_para or inherited_para.get(k) is None:
            inherited_para[k] = v
    default_para = inherited_para

    # inherited_char — 양식 heading의 char 값 전체 상속 (자간/장평 포함)
    heading_font_size = heading_char_shape.get("font_size", 11) or 11
    try:
        heading_font_size = float(heading_font_size)
    except Exception:
        heading_font_size = 11
    heading_hangul = (
        heading_char_shape.get("font_name_hangul", "")
        or heading_char_shape.get("font_name", "")
        or "맑은 고딕"
    )
    heading_latin = heading_char_shape.get("font_name_latin", "") or heading_hangul
    body_font_size = max(10, round(heading_font_size - 1))
    inherited_char = {
        "font_name": heading_hangul,
        "font_name_latin": heading_latin,
        "font_size": body_font_size,
        "bold": False,
        "italic": False,
        "color": [0, 0, 0],
        "char_spacing": heading_char_shape.get("char_spacing", 0),
        "width_ratio": heading_char_shape.get("width_ratio", 100),
    }

    user_char = body_style_opt.get("char", {}) or {}
    for k, v in user_char.items():
        if k not in inherited_char or not inherited_char.get(k):
            inherited_char[k] = v
    default_char = inherited_char

    # 5) paragraph style — hwp.set_para() + ParagraphShape Execute (align)
    align_int_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
    _para_kwargs = {}
    if "left_margin" in default_para:
        _para_kwargs["LeftMargin"] = float(default_para["left_margin"])
    if "indent" in default_para:
        indent_val = float(default_para["indent"])
        _para_kwargs["Indentation"] = indent_val
        if indent_val < 0 and "left_margin" not in default_para:
            _para_kwargs["LeftMargin"] = abs(indent_val)
    if "right_margin" in default_para:
        _para_kwargs["RightMargin"] = float(default_para["right_margin"])
    if "line_spacing" in default_para:
        _para_kwargs["LineSpacing"] = int(default_para["line_spacing"])
    if default_para.get("snap_to_grid"):
        _para_kwargs["SnapToGrid"] = 1
    if "space_before" in default_para and default_para["space_before"]:
        _para_kwargs["PrevSpacing"] = float(default_para["space_before"])
    if "space_after" in default_para and default_para["space_after"]:
        _para_kwargs["NextSpacing"] = float(default_para["space_after"])

    def _apply_align():
        """AlignType 을 ParagraphShape Execute 로 직접 설정."""
        if "align" not in default_para:
            return
        try:
            _act = hwp.HAction
            _pset = hwp.HParameterSet.HParaShape
            _act.GetDefault("ParagraphShape", _pset.HSet)
            _pset.AlignType = align_int_map.get(default_para["align"], 0)
            _act.Execute("ParagraphShape", _pset.HSet)
        except Exception as e:
            print(f"[WARN] align execute: {e}", file=sys.stderr)

    ps_result = {"ok": True}
    try:
        if _para_kwargs:
            hwp.set_para(**_para_kwargs)
        _apply_align()
    except Exception as e:
        print(f"[WARN] set_para: {e}", file=sys.stderr)
        ps_result = {"ok": False, "error": str(e)}

    def _reapply_para():
        """CharShape Execute / BreakPara 후 단락 서식 재적용."""
        try:
            if _para_kwargs:
                hwp.set_para(**_para_kwargs)
            _apply_align()
        except Exception:
            pass

    # 6) 본문 삽입 — 줄바꿈은 BreakPara + 매 단락 서식 재적용
    inserted = 0
    failed_lines = 0
    try:
        from hwp_editor import insert_text_with_style as _ins
    except Exception:
        _ins = None

    lines = body_text.split("\n")
    for i, line in enumerate(lines):
        if line:
            try:
                if _ins is not None:
                    _ins(hwp, line, default_char)
                    _reapply_para()
                else:
                    hwp.insert_text(line)
                inserted += len(line)
                # v0.7.9: 마커/번호 자동 내어쓰기 (IndentAtCaret)
                _apply_indent_at_caret(hwp, line)
            except Exception as e:
                print(f"[WARN] insert_text_with_style: {e}", file=sys.stderr)
                try:
                    hwp.insert_text(line)
                    inserted += len(line)
                except Exception:
                    failed_lines += 1
        if i < len(lines) - 1:
            try:
                hwp.HAction.Run("BreakPara")
                _reapply_para()
            except Exception as e:
                print(f"[WARN] BreakPara inner: {e}", file=sys.stderr)

    return {
        "status": "ok" if failed_lines == 0 else "partial",
        "heading": heading,
        "match_tier": match_tier,
        "matched_text": find_string,
        "total_matches": total_matches,
        "occurrence_used": target_occurrence,
        "insertion_mode": insertion_mode,
        "heading_depth": heading_depth,
        "inserted_chars": inserted,
        "failed_lines": failed_lines,
        "para_style_result": ps_result,
        "heading_style_detected": {
            "char": heading_char_shape,
            "para": heading_para_shape,
        },
        "body_style_applied": {"char": default_char, "para": default_para},
    }


@register("extend_section")
def extend_section(hwp, params):
    """섹션 확장 — 제목을 찾아 그 직후에 텍스트 삽입 (v0.7.1)."""
    validate_params(params, ["section_identifier", "content"], "extend_section")
    section_id = params["section_identifier"]  # {by: "title|index", value: ...}
    content = params["content"]
    preserve_format = bool(params.get("preserve_format", True))

    if isinstance(section_id, dict) and section_id.get("by") == "title":
        title = section_id.get("value", "")
        try:
            hwp.HAction.Run("MoveDocBegin")
            act = hwp.HAction
            pset = hwp.HParameterSet.HFindReplace
            act.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = title
            pset.Direction = 0
            pset.IgnoreMessage = 1
            if not act.Execute("RepeatFind", pset.HSet):
                return {"status": "error", "error": f"섹션 제목을 찾을 수 없습니다: {title}"}
            hwp.HAction.Run("MoveLineEnd")
            hwp.HAction.Run("BreakPara")
        except Exception as e:
            return {"status": "error", "error": f"섹션 위치 이동 실패: {e}"}

    try:
        for line in content.split("\n"):
            if line.strip():
                hwp.insert_text(line)
                hwp.HAction.Run("BreakPara")
        return {
            "status": "ok",
            "section_identifier": section_id,
            "inserted_paragraphs": len([l for l in content.split("\n") if l.strip()]),
            "preserve_format": preserve_format,
        }
    except Exception as e:
        return {"status": "error", "error": f"텍스트 삽입 실패: {e}"}
