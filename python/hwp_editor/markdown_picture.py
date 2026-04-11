"""hwp_editor.markdown_picture — 마크다운, 이미지, 참고자료 매핑, 전체 텍스트 추출.

함수:
- insert_markdown              : 마크다운 → HWP 서식 변환 삽입 (제목/목록/표/인용)
- insert_picture               : 이미지 삽입 (width/height/treat_as_char/embedded)
- auto_map_reference_to_table  : 참고자료 헤더-값 → 표 라벨 자동 매칭
- extract_all_text             : InitScan/GetText/ReleaseScan 안전 전체 텍스트 추출

v0.6.6 B3: scan_context 기반 extract_all_text
v0.7.4.8 Fix B1: auto_map _find_next_empty_cell (off-by-one 해소)
"""
import os
import re
import sys

from .text_style import insert_text_with_style


def insert_markdown(hwp, md_text):
    """마크다운 텍스트를 한글 서식으로 변환하여 삽입.

    지원: # 제목, **굵게**, *기울임*, - 목록, | 표 |, > 인용, --- 구분선.
    BUG-5 fix: 마크다운 표 파싱 추가.
    """
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
                # 셀 파싱 (H3 fix: 빈 셀 유지)
                raw_cells = row_text.split('|')
                cells = [c.strip() for c in raw_cells[1:-1]]
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
                    if hwp.is_cell():
                        hwp.MovePos(3)
                except Exception as e:
                    print(f"[WARN] Table exit after markdown table: {e}", file=sys.stderr)
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


def insert_picture(hwp, file_path, width=0, height=0, treat_as_char=None, embedded=None):
    """현재 커서 위치에 이미지 삽입.

    file_path: 이미지 파일 경로
    width/height: mm 단위 (0이면 원본 크기)
    treat_as_char: 글자처럼 취급 (None=pyhwpx 기본, True/False)
    embedded: 본문 안에 박힘 (None=기본)
    """
    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {file_path}")

    # v0.7.3 #7: pyhwpx 는 소문자 width/height 키 요구
    # v0.7.3 #8: treat_as_char/embedded 옵션 지원
    kwargs = {}
    if width:
        kwargs["width"] = int(width * 283.46)  # mm → HWPUNIT
    if height:
        kwargs["height"] = int(height * 283.46)
    if treat_as_char is not None:
        kwargs["treat_as_char"] = treat_as_char
    if embedded is not None:
        kwargs["embedded"] = embedded
    hwp.insert_picture(file_path, **kwargs)
    return {
        "status": "ok",
        "file_path": file_path,
        "width_mm": width,
        "height_mm": height,
        "treat_as_char": treat_as_char,
        "embedded": embedded,
    }


def auto_map_reference_to_table(hwp, table_idx, ref_headers, ref_row):
    """참고자료의 헤더와 표의 라벨을 자동 매칭하여 채울 데이터 생성.

    ref_headers: ["기업명", "대표자", "전화번호", ...]
    ref_row: ["(주)플랜아이", "이명기", "042-934-3508", ...]

    Returns: {"mappings": [{header, matched_label, tab, text}, ...], "unmapped": [...]}
    """
    # Lazy import: hwp_analyzer (utility package)
    from hwp_analyzer import map_table_cells, _match_label

    cell_data = map_table_cells(hwp, table_idx)
    cell_map = cell_data.get("cell_map", [])

    mappings = []
    unmapped = []

    # v0.7.4.8 Fix B1: off-by-one bug 해소 — j+1 고정 대신 라벨 셀 이후 첫 번째 빈 셀 탐색
    def _find_next_empty_cell(cmap, start_j, max_lookahead=4):
        """cmap[start_j+1 .. start_j+max_lookahead] 중 첫 번째 '빈 셀 또는 값 셀' 반환."""
        for offset in range(1, max_lookahead + 1):
            k = start_j + offset
            if k >= len(cmap):
                break
            candidate_text = str(cmap[k].get("text", "") or "").strip()
            # 빈 셀이면 즉시 선택
            if not candidate_text:
                return k
            # 다른 라벨 후보 감지 (콜론으로 끝나거나 1-10자 한국어 명사 패턴)
            if re.match(r"^[가-힣A-Za-z_\s]{1,12}[:：]$", candidate_text):
                break  # 다음 라벨 만남 → 여기부터는 다른 필드
            # 값 셀로 추정 — 이것도 반환
            return k
        # fallback: j+1
        return start_j + 1 if start_j + 1 < len(cmap) else None

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
                target_tab = _find_next_empty_cell(cell_map, j)
                if target_tab is not None and target_tab < len(cell_map):
                    mappings.append({
                        "header": header,
                        "matched_label": cell["text"].strip()[:30],
                        "tab": target_tab,
                        "text": str(value),
                        "match_type": "exact" if is_exact else "partial",
                        "match_ratio": round(ratio, 2),
                    })
                    matched = True
                    break
        if not matched:
            unmapped.append({"header": header, "value": str(value)})

    return {"mappings": mappings, "unmapped": unmapped, "total_matched": len(mappings)}


def extract_all_text(hwp, max_chars=200000, max_iters=50000, strip_each=False, separator="\n"):
    """InitScan/GetText/ReleaseScan 자동 안전 텍스트 추출.

    hwp_constants.scan_context로 ReleaseScan() finally 보장 (예외 시에도).

    Args:
        hwp: pyhwpx Hwp 인스턴스
        max_chars: 누적 문자 상한 (메모리 보호, 기본 20만자)
        max_iters: GetText 루프 상한 (무한 방지, 기본 5만회)
        strip_each: True면 각 GetText 결과 strip 후 빈 문자열 제외
        separator: 조각 결합 구분자 (기본 "\\n", 빈 문자열도 가능)

    Returns:
        조합된 문자열 (실패 시 빈 문자열)
    """
    from hwp_constants import scan_context

    parts = []
    total_chars = 0

    try:
        with scan_context(hwp):
            for _ in range(max_iters):
                try:
                    state, t = hwp.GetText()
                except Exception as e:
                    print(f"[WARN] extract_all_text GetText failed: {e}", file=sys.stderr)
                    break

                if state <= 0:
                    break

                if not t:
                    continue

                if strip_each:
                    t = t.strip()
                    if not t:
                        continue

                parts.append(t)
                total_chars += len(t)

                if total_chars >= max_chars:
                    print(f"[INFO] extract_all_text: max_chars {max_chars} reached",
                          file=sys.stderr)
                    break
    except Exception as e:
        print(f"[WARN] extract_all_text scan failed: {e}", file=sys.stderr)

    return separator.join(parts)
