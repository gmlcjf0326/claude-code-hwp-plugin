"""HWP Core — Content insertion handlers.

v0.7.6.0 P1-5: 28 핸들러 (단순한 콘텐츠 삽입 + 메트릭 + 작성요령 + PDF 클론).
v0.7.9 Phase 8: 파일 분할 대신 **섹션 주석** 유지 (28 micro-handler 는 분할 이득 < 노이즈).

섹션 구분 (PYTHON_INDEX.md 에서 anchor 로 참조):
  § BASIC INSERTIONS     — line 14~163   (13 handlers)
  § METRICS & SCAN       — line 165~203  (3 handlers)
  § INDENT               — line 205~229  (1 handler, outdent 는 line 769)
  § COMPLEX INSERTIONS   — line 231~461  (6 handlers)
  § GUIDE TEXT (v0.7.7)  — line 475~710  (3 handlers)
  § PDF CLONE            — line 721~767  (1 handler)
  § OUTDENT              — line 769~     (1 handler, indent 짝)

각 handler 는 대부분 hwp.HAction.Run(액션명) 호출만 함. 복잡한 핸들러 (extract_guide_text,
clone_pdf_to_hwp) 는 전담 모듈 (content_gen, pdf_clone) 로 위임.
"""
import re
import sys
from . import register
from ._helpers import validate_params


# ══════════════════════════════════════════════════════════════════════════
# SECTION: BASIC INSERTIONS (13 handlers)
# ══════════════════════════════════════════════════════════════════════════
# insert_page_break, break_section, break_column, insert_date_code,
# insert_auto_num, insert_memo, insert_line, insert_caption, insert_hyperlink,
# insert_footnote, insert_endnote, insert_markdown, insert_picture
# ══════════════════════════════════════════════════════════════════════════

@register("insert_page_break")
def insert_page_break(hwp, params):
    """페이지 나누기."""
    try:
        hwp.HAction.Run("BreakPage")
        return {"status": "ok"}
    except Exception as e:
        raise RuntimeError(f"페이지 나누기 실패: {e}")


@register("break_section")
def break_section(hwp, params):
    """섹션 나누기."""
    hwp.BreakSection()
    return {"status": "ok", "type": "section"}


@register("break_column")
def break_column(hwp, params):
    """단 나누기."""
    hwp.BreakColumn()
    return {"status": "ok", "type": "column"}


@register("insert_date_code")
def insert_date_code(hwp, params):
    """날짜 코드 삽입."""
    try:
        hwp.InsertDateCode()
    except Exception:
        hwp.HAction.Run("InsertDateCode")
    return {"status": "ok"}


@register("insert_auto_num")
def insert_auto_num(hwp, params):
    """자동 번호 삽입."""
    hwp.HAction.Run("InsertAutoNum")
    return {"status": "ok"}


@register("insert_memo")
def insert_memo(hwp, params):
    """필드 메모 삽입."""
    hwp.HAction.Run("InsertFieldMemo")
    return {"status": "ok"}


@register("insert_line")
def insert_line(hwp, params):
    """선 삽입 (대화상자 없이)."""
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HDrawLineAttr
        act.GetDefault("DrawLine", pset.HSet)
        act.Execute("DrawLine", pset.HSet)
        return {"status": "ok"}
    except Exception as e:
        raise RuntimeError(f"선 삽입 실패: {e}")


@register("insert_caption")
def insert_caption(hwp, params):
    """캡션 삽입 (표/그림 제목). side: 0=왼쪽 1=오른쪽 2=위 3=아래."""
    text = params.get("text", "")
    side = params.get("side", 3)
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HCaption
        act.GetDefault("InsertCaption", pset.HSet)
        pset.Side = int(side)
        act.Execute("InsertCaption", pset.HSet)
        if text:
            hwp.insert_text(text)
        return {"status": "ok", "text": text, "side": side}
    except Exception as e:
        raise RuntimeError(f"캡션 삽입 실패: {e}")


@register("insert_hyperlink")
def insert_hyperlink(hwp, params):
    """하이퍼링크 삽입."""
    validate_params(params, ["url"], "insert_hyperlink")
    url = params["url"]
    text = params.get("text", url)
    try:
        hwp.insert_hyperlink(url, text)
    except TypeError:
        hwp.insert_hyperlink(url)
    return {"status": "ok", "url": url, "text": text}


@register("insert_footnote")
def insert_footnote(hwp, params):
    """각주 삽입."""
    try:
        hwp.HAction.Run("InsertFootnote")
        text = params.get("text")
        if text:
            hwp.insert_text(text)
        hwp.HAction.Run("CloseEx")
        return {"status": "ok", "type": "footnote"}
    except Exception as e:
        try:
            hwp.HAction.Run("CloseEx")
        except Exception:
            pass
        raise RuntimeError(f"각주 삽입 실패: {e}")


@register("insert_endnote")
def insert_endnote(hwp, params):
    """미주 삽입."""
    try:
        hwp.HAction.Run("InsertEndnote")
        text = params.get("text")
        if text:
            hwp.insert_text(text)
        hwp.HAction.Run("CloseEx")
        return {"status": "ok", "type": "endnote"}
    except Exception as e:
        try:
            hwp.HAction.Run("CloseEx")
        except Exception:
            pass
        raise RuntimeError(f"미주 삽입 실패: {e}")


@register("insert_markdown")
def insert_markdown(hwp, params):
    """마크다운 → HWP 구문 삽입 (hwp_editor 위임)."""
    validate_params(params, ["text"], "insert_markdown")
    from hwp_editor import insert_markdown as _insert
    return _insert(hwp, params["text"])


@register("insert_picture")
def insert_picture_handler(hwp, params):
    """이미지 삽입 (hwp_editor 위임, v0.7.3 #7+#8 treat_as_char/embedded 지원)."""
    validate_params(params, ["file_path"], "insert_picture")
    from hwp_editor import insert_picture as _insert
    return _insert(
        hwp,
        params["file_path"],
        params.get("width", 0),
        params.get("height", 0),
        treat_as_char=params.get("treat_as_char"),
        embedded=params.get("embedded"),
    )


# ══════════════════════════════════════════════════════════════════════════
# SECTION: METRICS & SCAN (3 handlers)
# ══════════════════════════════════════════════════════════════════════════
# privacy_scan, list_controls, word_count
# ══════════════════════════════════════════════════════════════════════════

@register("privacy_scan")
def privacy_scan(hwp, params):
    """개인정보 스캔 (주민번호/전화/이메일 등)."""
    validate_params(params, ["text"], "privacy_scan")
    from privacy_scanner import scan_privacy
    return scan_privacy(params["text"])


@register("list_controls")
def list_controls(hwp, params):
    """HeadCtrl 순회로 모든 컨트롤(표/그림/머리말 등) 나열 (v0.6.6 B2)."""
    from hwp_traversal import traverse_all_ctrls
    filter_ids = params.get("filter")
    max_visits = params.get("max_visits", 5000)
    return traverse_all_ctrls(hwp, include_ids=filter_ids, max_visits=max_visits)


@register("word_count")
def word_count(hwp, params):
    """글자수/단어수/문단수/페이지수 (v0.6.6 B3 extract_all_text)."""
    from hwp_editor import extract_all_text
    text = ""
    try:
        text = extract_all_text(hwp, max_iters=10000, strip_each=False, separator="")
    except Exception as e:
        print(f"[WARN] word_count extract: {e}", file=sys.stderr)
    chars_total = len(text)
    chars_no_space = len(text.replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", ""))
    words = len(text.split())
    paragraphs = text.count("\n") + 1
    return {
        "status": "ok",
        "chars_total": chars_total,
        "chars_no_space": chars_no_space,
        "words": words,
        "paragraphs": paragraphs,
        "pages": hwp.PageCount,
    }


# ══════════════════════════════════════════════════════════════════════════
# SECTION: INDENT / OUTDENT (2 handlers, outdent is at the end of file)
# ══════════════════════════════════════════════════════════════════════════

@register("indent")
def indent_para(hwp, params):
    """들여쓰기: LeftMargin 증가."""
    depth = params.get("depth", 10)  # pt 단위
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HParaShape
        act.GetDefault("ParagraphShape", pset.HSet)
        current_left = 0
        try:
            current_left = pset.LeftMargin or 0
        except Exception:
            pass
        new_left = current_left + int(depth * 100)
        pset.LeftMargin = new_left
        act.Execute("ParagraphShape", pset.HSet)
        return {"status": "ok", "left_margin_pt": new_left / 100}
    except Exception as e:
        raise RuntimeError(f"들여쓰기 실패: {e}")


# ══════════════════════════════════════════════════════════════════════════
# SECTION: COMPLEX INSERTIONS (6 handlers)
# ══════════════════════════════════════════════════════════════════════════
# insert_textbox, draw_line, image_extract, insert_page_num, generate_toc,
# create_gantt_chart
# ══════════════════════════════════════════════════════════════════════════

@register("insert_textbox")
def insert_textbox(hwp, params):
    """글상자 삽입 — multi-fallback (pyhwpx / HShapeObject / DrawObjTextBoxNew)."""
    x = params.get("x", 0)
    y = params.get("y", 0)
    width = params.get("width", 60)
    height = params.get("height", 30)
    text = params.get("text", "")

    # 시도 1: pyhwpx 헬퍼
    for helper_name in ("create_textbox", "insert_textbox", "create_text_box"):
        if hasattr(hwp, helper_name):
            try:
                fn = getattr(hwp, helper_name)
                fn(x=int(x * 283.465), y=int(y * 283.465),
                   width=int(width * 283.465), height=int(height * 283.465))
                if text:
                    hwp.insert_text(text)
                try:
                    hwp.MovePos(3)
                except Exception:
                    pass
                return {"status": "ok", "method": f"pyhwpx_{helper_name}",
                        "x": x, "y": y, "width": width, "height": height}
            except Exception as e:
                print(f"[INFO] {helper_name} failed: {e}", file=sys.stderr)
                continue

    # 시도 2: HShapeObject
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HShapeObject
        act.GetDefault("InsertDrawObj", pset.HSet)
        try:
            pset.HSet.SetItem("ShapeType", 1)
        except Exception:
            pass
        try:
            pset.HorzRelTo = 0
            pset.VertRelTo = 0
            pset.HorzOffset = int(x * 283.465)
            pset.VertOffset = int(y * 283.465)
            pset.Width = int(width * 283.465)
            pset.Height = int(height * 283.465)
        except Exception as e:
            print(f"[INFO] HShapeObject attr: {e}", file=sys.stderr)
        act.Execute("InsertDrawObj", pset.HSet)
        if text:
            hwp.insert_text(text)
        try:
            hwp.MovePos(3)
        except Exception:
            pass
        return {"status": "ok", "method": "HShapeObject_no_ShapeType",
                "x": x, "y": y, "width": width, "height": height}
    except Exception as e:
        print(f"[INFO] HShapeObject path failed: {e}", file=sys.stderr)

    # 시도 3: DrawObjTextBoxNew
    try:
        hwp.HAction.Run("DrawObjTextBoxNew")
        if text:
            hwp.insert_text(text)
        try:
            hwp.MovePos(3)
        except Exception:
            pass
        return {"status": "ok", "method": "DrawObjTextBoxNew_no_position",
                "warning": "위치/크기 미지정"}
    except Exception as e:
        raise RuntimeError(f"글상자 생성 모든 fallback 실패: {e}")


@register("draw_line")
def draw_line(hwp, params):
    """선 그리기 (두께/색상/스타일)."""
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HDrawLineAttr
        act.GetDefault("DrawLine", pset.HSet)
        if "width" in params:
            pset.Width = int(params["width"])
        if "color" in params:
            c = params["color"]
            if isinstance(c, str):
                r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
                pset.Color = hwp.RGBColor(r, g, b)
            elif isinstance(c, list):
                pset.Color = hwp.RGBColor(c[0], c[1], c[2])
        if "style" in params:
            pset.style = int(params["style"])
        act.Execute("DrawLine", pset.HSet)
        return {"status": "ok"}
    except Exception as e:
        raise RuntimeError(f"선 그리기 실패: {e}")


@register("image_extract")
def image_extract(hwp, params):
    """문서 내 이미지 추출 (save_all_pictures + HWPX ZIP fallback)."""
    import os
    validate_params(params, ["output_dir"], "image_extract")
    output_dir = os.path.abspath(params["output_dir"])
    os.makedirs(output_dir, exist_ok=True)
    temp_dir = os.path.join(os.getcwd(), "temp", "binData")
    os.makedirs(temp_dir, exist_ok=True)
    extracted_ok = False
    try:
        hwp.save_all_pictures(output_dir)
        extracted_ok = True
    except Exception:
        try:
            import zipfile
            temp_hwpx = os.path.join(output_dir, "_temp.hwpx")
            hwp.save_as(temp_hwpx, "HWPX")
            if os.path.exists(temp_hwpx):
                with zipfile.ZipFile(temp_hwpx, 'r') as z:
                    for name in z.namelist():
                        if name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                            z.extract(name, output_dir)
                os.remove(temp_hwpx)
                extracted_ok = True
        except Exception as e2:
            raise RuntimeError(f"이미지 추출 실패: {e2}")
    files = []
    for root, dirs, fnames in os.walk(output_dir):
        for fname in fnames:
            if fname.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.wmf', '.emf')):
                rel = os.path.relpath(os.path.join(root, fname), output_dir)
                files.append(rel)
    return {"status": "ok", "output_dir": output_dir, "extracted": len(files), "files": files}


@register("insert_page_num")
def insert_page_num(hwp, params):
    """페이지 번호 삽입 (plain/dash/paren)."""
    fmt = params.get("format", "plain")
    prefix_suffix = {"dash": ("- ", " -"), "paren": ("(", ")"), "plain": ("", "")}
    prefix, suffix = prefix_suffix.get(fmt, ("", ""))
    if prefix:
        hwp.insert_text(prefix)
    hwp.HAction.Run("InsertPageNum")
    if suffix:
        hwp.insert_text(suffix)
    return {"status": "ok", "format": fmt}


@register("generate_toc")
def generate_toc(hwp, params):
    """목차 자동 생성 (제목 패턴 감지 + 삽입)."""
    import re
    from hwp_editor import extract_all_text
    text_blob = extract_all_text(hwp, max_iters=1000, strip_each=True, separator="\n")
    texts = text_blob.split("\n") if text_blob else []
    toc_items = []
    heading_patterns = [
        (r'^(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)[.\s]', 1),
        (r'^(\d+)\.\s', 2),
        (r'^(가|나|다|라|마|바|사)\.\s', 3),
    ]
    for t in texts:
        for pattern, level in heading_patterns:
            if re.match(pattern, t):
                toc_items.append({"level": level, "text": t[:60]})
                break
    if params.get("insert", True):
        from hwp_editor import insert_text_with_style
        insert_text_with_style(hwp, "목   차\r\n", {"bold": True, "font_size": 16})
        hwp.insert_text("\r\n")
        for item in toc_items:
            indent = "  " * (item["level"] - 1)
            hwp.insert_text(f"{indent}{item['text']}\r\n")
        hwp.insert_text("\r\n")
    return {"status": "ok", "toc_items": len(toc_items), "items": toc_items[:30]}


@register("create_gantt_chart")
def create_gantt_chart(hwp, params):
    """간트 차트 표 생성 (tasks × months)."""
    validate_params(params, ["tasks", "months"], "create_gantt_chart")
    tasks = params["tasks"]
    months = params["months"]
    month_label = params.get("month_label", "M+N")
    header = ["세부 업무", "수행내용"]
    for i in range(months):
        if month_label == "M+N":
            header.append(f"M+{i}" if i > 0 else "M")
        else:
            header.append(f"{i+1}월")
    header.append("비중(%)")
    data = [header]
    active_cells = []
    for task_idx, task in enumerate(tasks):
        row = [task.get("name", ""), task.get("desc", "")]
        start = task.get("start", 1)
        end = task.get("end", 1)
        for m in range(months):
            if start <= m + 1 <= end:
                row.append("■")
                tab = (task_idx + 1) * len(header) + 2 + m
                active_cells.append(tab)
            else:
                row.append("")
        row.append(str(task.get("weight", "")))
        data.append(row)
    rows = len(data)
    cols = len(data[0])
    hwp.create_table(rows, cols)
    from hwp_editor import insert_text_with_style
    filled = 0
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            if val:
                if r == 0:
                    insert_text_with_style(hwp, str(val), {"bold": True})
                else:
                    hwp.insert_text(str(val))
                filled += 1
            if c < len(row) - 1 or r < rows - 1:
                hwp.TableRightCell()
    from ._helpers import _exit_table_safely
    _exit_table_safely(hwp)
    try:
        from hwp_editor import set_cell_background_color
        style_cells = [{"tab": i, "color": "#666666"} for i in range(cols)]
        style_cells += [{"tab": t, "color": "#C0C0C0"} for t in active_cells]
        set_cell_background_color(hwp, -1, style_cells)
    except Exception as e:
        print(f"[WARN] {e}", file=sys.stderr)
    return {"status": "ok", "rows": rows, "cols": cols, "filled": filled, "active_cells": len(active_cells)}


# ---------------------------------------------------------------------------
# v0.7.7 — 작성요령 추출 (1C)
# ---------------------------------------------------------------------------

_GUIDE_TABLE_KEYWORDS = ["작성요령", "유의사항", "참고", "주의사항"]

# ══════════════════════════════════════════════════════════════════════════
# SECTION: GUIDE TEXT (v0.7.7) — 3 handlers
# ══════════════════════════════════════════════════════════════════════════
# extract_guide_text, delete_guide_text, verify_after_fill
# 양식의 "작성요령" 박스 내용 추출/삭제 + 채우기 검증
# ══════════════════════════════════════════════════════════════════════════

# 제약 조건 파싱 패턴
_CONSTRAINT_PAGE_LIMIT = re.compile(r'(\d+)\s*페이지\s*이내')
_CONSTRAINT_REQUIRED = re.compile(r'(반드시|필히|필수|필수적)\s*(기재|포함|작성|서술|제시)')
_CONSTRAINT_FORMAT = re.compile(r'(표\s*형식|도표|그림|사진|그래프|차트)\s*(포함|첨부|작성|제시)?')


@register("extract_guide_text")
def extract_guide_text(hwp, params):
    """작성요령 박스 내용을 구조화 추출 (v0.7.7).

    delete_guide_text 전에 호출하여 각 섹션별 요구사항을 보존.
    반환: guides[] — 각각 {heading, raw_text, constraints, source}
    """
    guides = []
    try:
        from hwp_analyzer import analyze_document as _analyze
        from ._state import get_current_doc_path
        doc_path = get_current_doc_path()
        if not doc_path:
            return {"status": "error", "error": "열린 문서가 없습니다.", "guides": [], "count": 0}
        analysis = _analyze(hwp, doc_path, already_open=True)
        tables = analysis.get("tables", []) if isinstance(analysis, dict) else []
        full_text = analysis.get("full_text", "") or ""

        # 1. 작성요령 표에서 추출
        for t in tables:
            headers = t.get("headers", []) or []
            data = t.get("data", []) or []
            first_cell = str(headers[0] or "").strip() if headers else ""
            if not first_cell and data and data[0]:
                first_cell = str(data[0][0] or "").strip()

            is_guide = any(kw in first_cell for kw in _GUIDE_TABLE_KEYWORDS)
            if not is_guide:
                continue

            # 표 전체 텍스트 합치기
            raw_parts = []
            for row in ([headers] + data):
                for cell in row:
                    cell_text = str(cell or "").strip()
                    if cell_text:
                        raw_parts.append(cell_text)
            raw_text = "\n".join(raw_parts)

            # 가장 가까운 상위 제목 찾기
            heading = _find_nearest_heading(full_text, t.get("index", 0), tables)

            # 제약 조건 파싱
            constraints = _parse_constraints(raw_text)

            guides.append({
                "heading": heading,
                "table_index": t.get("index", 0),
                "raw_text": raw_text[:1000],  # 최대 1000자
                "constraints": constraints,
                "source": "table",
            })

        # 2. 텍스트 패턴에서 추출 (< 작성요령 > 등)
        guide_text_patterns = [
            "< 작성요령 >", "＜ 작성요령 ＞", "<작성요령>",
            "< 유의사항 >", "＜ 유의사항 ＞", "<유의사항>",
            "※ 작성요령", "※ 유의사항",
        ]
        for pat in guide_text_patterns:
            start = 0
            while True:
                idx = full_text.find(pat, start)
                if idx < 0:
                    break
                # 패턴 뒤 ~ 다음 제목 or 500자까지 추출
                after = full_text[idx:idx + 500]
                heading = _find_heading_before(full_text, idx)
                constraints = _parse_constraints(after)
                guides.append({
                    "heading": heading,
                    "raw_text": after[:500],
                    "constraints": constraints,
                    "source": "text_pattern",
                    "pattern": pat,
                })
                start = idx + len(pat)

    except Exception as e:
        import traceback
        print(f"[WARN] extract_guide_text: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        return {"status": "ok", "guides": guides, "count": len(guides),
                "debug_error": str(e)}

    return {"status": "ok", "guides": guides, "count": len(guides)}


def _parse_constraints(text):
    """작성요령 텍스트에서 제약 조건 파싱."""
    constraints = {}

    # 페이지 제한
    m = _CONSTRAINT_PAGE_LIMIT.search(text)
    if m:
        constraints["max_pages"] = int(m.group(1))

    # 필수 기재 키워드
    required = []
    for m in _CONSTRAINT_REQUIRED.finditer(text):
        # 매칭 주변 ±30자에서 핵심어 추출
        ctx_start = max(0, m.start() - 30)
        ctx = text[ctx_start:m.end() + 20]
        required.append(ctx.strip())
    if required:
        constraints["required_mentions"] = required[:5]

    # 형식 요구
    formats = []
    for m in _CONSTRAINT_FORMAT.finditer(text):
        formats.append(m.group(1).strip())
    if formats:
        constraints["format_hints"] = list(set(formats))

    return constraints


def _find_nearest_heading(full_text, table_index, tables):
    """표 index 기준으로 가장 가까운 상위 제목 텍스트 추출."""
    # 제목 패턴: "1. ", "가. ", "(1) ", "다. " 등
    heading_pattern = re.compile(
        r'^[\s]*(?:\d+\.\s|[가-힣]\.\s|\(\d+\)\s|\d+\.\d+\s)',
        re.MULTILINE
    )
    # 본문 텍스트에서 표 앞쪽의 마지막 제목 찾기
    # 간이 방법: full_text 을 줄 단위로 역순 탐색
    lines = full_text.split('\n')
    # 표의 대략적 위치: 표 이전에 나타나는 마지막 제목
    # table_index 를 순서 힌트로 사용
    last_heading = ""
    table_count_seen = 0
    for line in lines:
        stripped = line.strip()
        if heading_pattern.match(stripped):
            last_heading = stripped
        # 간이 표 감지 (탭 구분 데이터)
        if '\t' in line and len(line.split('\t')) >= 3:
            table_count_seen += 1
            if table_count_seen > table_index:
                break
    return last_heading[:100] if last_heading else "(알 수 없음)"


def _find_heading_before(full_text, char_pos):
    """char_pos 이전의 가장 가까운 제목 텍스트."""
    heading_pattern = re.compile(
        r'(?:^\s*(?:\d+\.\s|[가-힣]\.\s|\(\d+\)\s).*)',
        re.MULTILINE
    )
    text_before = full_text[:char_pos]
    matches = list(heading_pattern.finditer(text_before))
    if matches:
        return matches[-1].group().strip()[:100]
    return "(알 수 없음)"


@register("delete_guide_text")
def delete_guide_text(hwp, params):
    """작성요령/가이드 텍스트 자동 삭제 (v0.7.5.4 P3-3 / v0.7.7 extract_first 추가)."""
    from ._helpers import _execute_all_replace

    # v0.7.7: 삭제 전 추출 옵션
    extract_first = bool(params.get("extract_first", False))
    extracted_guides = None
    if extract_first:
        extracted_guides = extract_guide_text(hwp, {})

    scope = params.get("scope", "text")
    default_patterns = [
        "< 작성요령 >", "＜ 작성요령 ＞", "<작성요령>",
        "< 유의사항 >", "＜ 유의사항 ＞", "<유의사항>",
        "※ 작성요령", "※ 유의사항",
    ]
    default_table_keywords = ["작성요령", "유의사항", "참고", "주의사항"]
    patterns = params.get("patterns", default_patterns)
    table_keywords = params.get("table_keywords", default_table_keywords)

    deleted_text_patterns = 0
    deleted_tables = 0
    table_details = []
    hwp.MovePos(2)

    if scope in ("text", "both"):
        for pat in patterns:
            replaced = _execute_all_replace(hwp, pat, "", False)
            if replaced:
                deleted_text_patterns += 1

    if scope in ("table", "both"):
        try:
            from hwp_analyzer import analyze_document as _analyze
            analysis = _analyze(hwp, None, already_open=True)
            tables = analysis.get("tables", []) if isinstance(analysis, dict) else []
            candidates = []
            for t in tables:
                headers = t.get("headers", []) or []
                data = t.get("data", []) or []
                first_cell = ""
                if headers:
                    first_cell = str(headers[0] or "").strip()
                elif data and data[0]:
                    first_cell = str(data[0][0] or "").strip()
                for kw in table_keywords:
                    if kw in first_cell:
                        candidates.append({
                            "index": t.get("index", 0),
                            "first_cell": first_cell,
                            "keyword": kw,
                            "rows": t.get("rows", 0),
                            "cols": t.get("cols", 0),
                        })
                        break
            for cand in reversed(candidates):
                try:
                    hwp.MovePos(2)
                    hwp.get_into_nth_table(cand["index"])
                    hwp.HAction.Run("TableDeleteTable")
                    deleted_tables += 1
                    table_details.append(cand)
                except Exception as e:
                    print(f"[WARN] delete table {cand['index']}: {e}", file=sys.stderr)
        except Exception as e:
            print(f"[WARN] delete_guide_text table scan: {e}", file=sys.stderr)

    result = {
        "status": "ok",
        "scope": scope,
        "deleted_text_patterns": deleted_text_patterns,
        "deleted_tables": deleted_tables,
        "table_details": table_details,
        "patterns": patterns,
        "table_keywords": table_keywords,
    }
    if extracted_guides is not None:
        result["extracted_guides"] = extracted_guides.get("guides", [])
    return result


@register("verify_after_fill")
def verify_after_fill(hwp, params):
    """표 채우기 후 검증 (hwp_editor 위임)."""
    from hwp_editor import verify_after_fill as _verify
    validate_params(params, ["table_index", "expected_cells"], "verify_after_fill")
    return _verify(hwp, params["table_index"], params["expected_cells"])


# ══════════════════════════════════════════════════════════════════════════
# SECTION: PDF CLONE (1 handler)
# ══════════════════════════════════════════════════════════════════════════
# clone_pdf_to_hwp — pdf_clone 패키지 위임
# ══════════════════════════════════════════════════════════════════════════

@register("clone_pdf_to_hwp")
def clone_pdf_to_hwp(hwp, params):
    """PDF → HWP 클론 (v0.7.4.2). Lazy import OCR deps."""
    import os
    from ._state import set_current_doc_path
    validate_params(params, ["pdf_path", "output_path"], "clone_pdf_to_hwp")
    pdf_path = validate_file_path(params["pdf_path"], must_exist=True)
    output_path = validate_file_path(params["output_path"], must_exist=False)
    options = params.get("options", {}) or {}
    try:
        from pdf_clone import clone_pdf_to_hwp as _clone_fn
    except ImportError as e:
        return {
            "status": "error",
            "error_type": "missing_dependency",
            "error": f"pdf_clone 모듈 import 실패: {e}",
            "guide": (
                "PDF OCR 기본 의존성 설치:\n"
                "  pip install pdfplumber Pillow opencv-python-headless numpy\n"
                "스캔 PDF (PaddleOCR):\n"
                "  pip install 'paddlepaddle>=3.0.0' 'paddleocr>=3.0.0'"
            ),
        }
    try:
        _clone_result = _clone_fn(hwp, pdf_path, output_path, options)
        if isinstance(_clone_result, dict) and _clone_result.get("status") in ("ok", "partial"):
            if os.path.exists(output_path):
                set_current_doc_path(output_path)
        return _clone_result
    except ImportError as e:
        return {
            "status": "error",
            "error_type": "missing_ocr_engine",
            "error": f"OCR 엔진 import 실패: {e}",
            "guide": (
                "PaddleOCR 설치 (Python 3.13+):\n"
                "  pip install 'paddlepaddle>=3.0.0' 'paddleocr>=3.0.0'\n"
                "최초 실행 시 ~150MB 모델 자동 다운로드."
            ),
        }
    except Exception as e:
        return {
            "status": "error",
            "error_type": "clone_failed",
            "error": f"PDF clone 실패: {e}",
        }


# ══════════════════════════════════════════════════════════════════════════
# SECTION: OUTDENT (indent 짝, 앞 INDENT 섹션과 쌍을 이룸)
# ══════════════════════════════════════════════════════════════════════════

@register("outdent")
def outdent_para(hwp, params):
    """내어쓰기: LeftMargin 감소."""
    depth = params.get("depth", 10)
    try:
        act = hwp.HAction
        pset = hwp.HParameterSet.HParaShape
        act.GetDefault("ParagraphShape", pset.HSet)
        current_left = 0
        try:
            current_left = pset.LeftMargin or 0
        except Exception:
            pass
        new_left = max(0, current_left - int(depth * 100))
        pset.LeftMargin = new_left
        act.Execute("ParagraphShape", pset.HSet)
        return {"status": "ok", "left_margin_pt": new_left / 100}
    except Exception as e:
        raise RuntimeError(f"내어쓰기 실패: {e}")
