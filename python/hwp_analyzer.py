"""HWP Document Analyzer - Extract structure and content from HWP documents.
Uses pyhwpx Hwp() only. Raw win32com is forbidden.
"""
import sys
import os
import re


MAX_TABLES = 50  # 표 스캔 상한 (통장사본 등 반복 표 방지)


# ── 공백 정규화 ──
def _normalize(text):
    """모든 공백(스페이스, 탭, NBSP 등)을 제거하여 비교용 문자열 반환."""
    return re.sub(r"\s+", "", text)


# ── 라벨 별칭(alias) 사전 ──
# key: 정규화된 표준명, value: 정규화된 동의어 리스트
_LABEL_ALIASES = {
    "기업명": ["기업이름", "회사명", "상호명", "상호", "법인명", "업체명", "회사이름"],
    "사업자등록번호": ["사업자번호", "사업자등록NO", "사업자No"],
    "법인등록번호": ["법인번호", "법인등록No", "법인No"],
    "사업장주소": ["주소", "소재지", "본점소재지", "사업장소재지", "회사주소", "기업주소"],
    "대표자성명": ["대표자", "대표자명", "대표이사", "대표자이름", "대표이사명", "성명"],
    "대표전화번호": ["대표전화", "전화번호", "연락처", "대표번호", "전화", "TEL"],
    "홈페이지URL": ["홈페이지", "웹사이트", "URL", "홈페이지주소", "웹주소"],
    "이메일": ["이메일주소", "EMAIL", "E-MAIL"],
    "팩스번호": ["팩스", "FAX", "FAX번호"],
    "설립일": ["설립일자", "설립년월일", "법인설립일"],
    "업종": ["업종명", "주업종"],
    "업태": ["업태명", "주업태"],
    "종업원수": ["직원수", "임직원수", "종업원"],
    "자본금": ["납입자본금", "자본금액"],
    "매출액": ["연매출", "연매출액", "매출"],
}

# 역방향 룩업 테이블 생성: 동의어 -> 표준명
_ALIAS_LOOKUP = {}
for canonical, aliases in _LABEL_ALIASES.items():
    norm_canonical = _normalize(canonical)
    _ALIAS_LOOKUP[norm_canonical] = norm_canonical
    for alias in aliases:
        _ALIAS_LOOKUP[_normalize(alias)] = norm_canonical


def _canonical_label(label):
    """라벨을 정규화하고 표준명으로 변환. 별칭 없으면 정규화된 원본 반환."""
    norm = _normalize(label)
    return _ALIAS_LOOKUP.get(norm.upper(), _ALIAS_LOOKUP.get(norm, norm))


def _match_label(cell_text, search_label):
    """셀 텍스트와 검색 라벨이 같은 의미인지 판단.

    Returns: (is_match, is_exact, ratio)
      - is_match: 매칭 여부
      - is_exact: exact match 여부 (정규화 후 완전 일치)
      - ratio: 매칭률 (0.0~1.0, exact이면 1.0)
    """
    norm_cell = _normalize(cell_text)
    norm_label = _normalize(search_label)

    if not norm_cell or not norm_label:
        return False, False, 0.0

    # 1) 정규화 후 exact match (공백만 달랐던 경우)
    if norm_cell == norm_label:
        return True, True, 1.0

    # 2) 별칭 매칭: 둘 다 같은 표준명으로 매핑되는지
    canon_cell = _canonical_label(cell_text)
    canon_label = _canonical_label(search_label)
    if canon_cell == canon_label:
        return True, True, 1.0

    # 3) 정규화된 문자열 포함 관계 (partial match)
    if norm_label in norm_cell:
        return True, False, len(norm_label) / len(norm_cell)
    if norm_cell in norm_label:
        return True, False, len(norm_cell) / len(norm_label)

    return False, False, 0.0


def analyze_document(hwp, file_path, already_open=False):
    """Analyze an HWP document: pages, tables, fields, text."""
    file_path = os.path.abspath(file_path)
    # 항상 문서를 열어서 활성화 보장 (이미 열려있으면 해당 문서가 포커스됨)
    hwp.open(file_path)
    # 커서를 문서 처음으로 이동
    try:
        hwp.MovePos(2)  # movePOS_START: 문서 처음으로
    except Exception as e:
        print(f"[WARN] MovePos failed: {e}", file=sys.stderr)

    result = {
        "file_path": file_path,
        "file_name": os.path.basename(file_path),
        "file_format": "HWPX" if file_path.lower().endswith(".hwpx") else "HWP",
        "pages": 0,
        "tables": [],
        "fields": [],
        "text_preview": "",
        "full_text": "",
    }

    scan_started = False

    try:
        # Page count
        try:
            result["pages"] = hwp.PageCount
        except Exception as e:
            print(f"[WARN] PageCount failed: {e}", file=sys.stderr)

        # Extract tables (with data for AI context, max MAX_TABLES)
        try:
            # 커서를 문서 시작으로 초기화 (표 탐색 안정성 보장)
            try:
                hwp.MovePos(2)  # 문서 처음으로
            except Exception as e:
                print(f"[WARN] MovePos before table scan failed: {e}", file=sys.stderr)

            table_idx = 0
            prev_data = None
            while table_idx < MAX_TABLES:
                try:
                    hwp.get_into_nth_table(table_idx)
                    df = hwp.table_to_df()
                    current_data = df.values.tolist()
                    # 중복 감지: 이전 표와 동일한 데이터면 같은 표 반복 접근 → 중단
                    if prev_data is not None and current_data == prev_data:
                        print(f"[INFO] Table {table_idx} is duplicate of previous — stopping scan", file=sys.stderr)
                        try:
                            if hwp.is_cell():
                                hwp.MovePos(3)
                        except Exception as e:
                            print(f"[WARN] Table exit (dup detect): {e}", file=sys.stderr)
                        break
                    prev_data = current_data
                    table_info = {
                        "index": table_idx,
                        "rows": len(df) + 1,  # +1 for header
                        "cols": len(df.columns) if len(df) > 0 else 0,
                        "headers": [str(c) for c in df.columns],
                        "data": current_data,
                    }
                    result["tables"].append(table_info)
                    try:
                        if hwp.is_cell():
                            hwp.MovePos(3)
                        hwp.MovePos(2)
                    except Exception as e:
                        print(f"[WARN] Table exit/MovePos failed: {e}", file=sys.stderr)
                    table_idx += 1
                except Exception as e:
                    print(f"[INFO] Table scan stopped at idx {table_idx}: {e}", file=sys.stderr)
                    break
            # BUG-7 fix: 실제 발견된 표 수와 스캔 상한 분리
            result["scanned_tables"] = table_idx
            if table_idx >= MAX_TABLES:
                result["tables_truncated"] = True
                result["tables_truncated_message"] = f"표가 {MAX_TABLES}개 이상일 수 있습니다. 처음 {MAX_TABLES}개만 분석했습니다."
                print(f"[WARN] Table scan capped at {MAX_TABLES}", file=sys.stderr)
        except Exception as e:
            print(f"[WARN] Table extraction failed: {e}", file=sys.stderr)

        # Extract fields
        try:
            field_list = hwp.GetFieldList()
            if field_list:
                fields = field_list.split("\x02") if "\x02" in field_list else [field_list]
                for field in fields:
                    if field.strip():
                        value = ""
                        try:
                            value = hwp.GetFieldText(field.strip()) or ""
                        except Exception as e:
                            print(f"[WARN] GetFieldText failed: {e}", file=sys.stderr)
                        result["fields"].append({
                            "name": field.strip(),
                            "value": value,
                        })
        except Exception as e:
            print(f"[WARN] Field extraction failed: {e}", file=sys.stderr)

        # Extract full text (up to 15,000 chars for AI context)
        try:
            hwp.InitScan(0x0077)
            scan_started = True
            text_parts = []
            total_len = 0
            count = 0
            while total_len < 15000 and count < 5000:
                try:
                    state, text = hwp.GetText()
                    if state <= 0:
                        break
                    # state 1=일반텍스트, 2=표 안 텍스트 등
                    if text and text.strip():
                        text_parts.append(text.strip())
                        total_len += len(text)
                    count += 1
                except Exception:
                    break
            hwp.ReleaseScan()
            scan_started = False

            full = "\n".join(text_parts)
            result["full_text"] = full[:15000]
            result["text_preview"] = full[:500]
        except Exception as e:
            print(f"[WARN] Text extraction failed: {e}", file=sys.stderr)

    finally:
        # Guarantee ReleaseScan if InitScan was called
        if scan_started:
            try:
                hwp.ReleaseScan()
            except Exception as e:
                print(f"[WARN] ReleaseScan failed: {e}", file=sys.stderr)

    return result


def map_table_cells(hwp, table_idx, max_cells=200):
    """Map all navigable cells in a table by Tab traversal.

    Returns a list of cell entries with tab index and the text content
    found at each position. This helps identify which tab index
    corresponds to which cell in tables with merged cells.
    """
    cell_map = []

    try:
        hwp.get_into_nth_table(table_idx)
    except Exception as e:
        return {"error": f"Cannot enter table {table_idx}: {e}", "cell_map": []}

    prev_pos = None

    for i in range(max_cells):
        try:
            cur = hwp.GetPos()
            pos = (cur[0], cur[1], cur[2]) if cur else None

            # Detect if we've looped back to the start
            if i > 0 and pos == prev_pos:
                break

            # Read cell text (select all in cell, get text, then deselect)
            cell_text = ""
            try:
                hwp.HAction.Run("SelectAll")
                cell_text = hwp.GetTextFile("TEXT", "saveblock").strip()
            except Exception:
                cell_text = ""
            finally:
                try:
                    hwp.HAction.Run("Cancel")
                except Exception as e:
                    print(f"[WARN] Cancel in map_table_cells: {e}", file=sys.stderr)

            cell_map.append({
                "tab": i,
                "text": cell_text[:100],  # Truncate long text
                "pos": list(pos) if pos else None,
            })

            prev_pos = pos
            hwp.TableRightCell()
        except Exception:
            break

    try:
        if hwp.is_cell():
            hwp.MovePos(3)
    except Exception as e:
        print(f"[WARN] Table exit (map_table_cells): {e}", file=sys.stderr)

    return {
        "table_index": table_idx,
        "total_cells": len(cell_map),
        "cell_map": cell_map,
    }


def _group_cells_into_rows(cell_map):
    """셀 맵을 행 단위로 그룹화한다.

    행 경계 감지: list_id가 감소하면 새 행 시작.
    (병합 셀이 재방문되면 list_id가 이전 값으로 돌아감)
    """
    rows = []
    current_row = []
    prev_list_id = -1

    for cell in cell_map:
        list_id = cell["pos"][0] if cell.get("pos") else -1
        if list_id <= prev_list_id and current_row:
            rows.append(current_row)
            current_row = []
        current_row.append(cell)
        prev_list_id = list_id

    if current_row:
        rows.append(current_row)
    return rows


def _find_label_column(rows, label):
    """label 텍스트가 있는 셀의 (col_index, row_index, is_partial)를 반환.

    공백 정규화 + 별칭 사전으로 매칭. exact 우선, partial은 상위 행 우선.
    """
    # Exact match (정규화 + 별칭 포함)
    for row_idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            is_match, is_exact, _ = _match_label(cell["text"], label)
            if is_match and is_exact:
                return col_idx, row_idx, False
    # Partial match
    if len(_normalize(label)) < 2:
        return None, None, False
    best_col, best_row, best_score = None, None, 0
    for row_idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            is_match, is_exact, ratio = _match_label(cell["text"], label)
            if is_match and not is_exact and ratio > 0:
                score = ratio * max(0.1, 1.0 - row_idx * 0.05)
                if score > best_score:
                    best_col, best_row, best_score = col_idx, row_idx, score
    if best_col is not None:
        return best_col, best_row, True
    return None, None, False


def _find_label_row(rows, row_label):
    """row_label 텍스트가 있는 행의 (row_index, is_partial, matched_text)를 반환.

    공백 정규화 + 별칭 사전으로 매칭. exact 우선, partial fallback.
    """
    # Exact match (정규화 + 별칭 포함)
    for row_idx, row in enumerate(rows):
        for cell in row:
            is_match, is_exact, _ = _match_label(cell["text"], row_label)
            if is_match and is_exact:
                return row_idx, False, cell["text"].strip()
    # Partial match
    if len(_normalize(row_label)) < 2:
        return None, False, ""
    best_row, best_score, best_text = None, 0, ""
    for row_idx, row in enumerate(rows):
        for cell in row:
            is_match, is_exact, ratio = _match_label(cell["text"], row_label)
            if is_match and not is_exact and ratio > 0:
                if ratio > best_score:
                    best_row, best_score, best_text = row_idx, ratio, cell["text"].strip()
    if best_row is not None:
        return best_row, True, best_text
    return None, False, ""


def _find_cell_position_in_rows(rows, flat_idx):
    """flat cell_map 인덱스 → (row_idx, col_idx_in_row) 변환"""
    idx = 0
    for row_idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            if idx == flat_idx:
                return row_idx, col_idx
            idx += 1
    return None, None


def _find_cell_in_flat(cell_map, label):
    """flat cell_map에서 라벨 텍스트 매칭. exact 우선, partial fallback.

    공백 정규화 + 별칭 사전 적용.
    Returns (matched_idx, is_partial).
    """
    if not label:
        return None, False  # 이중 방어: 호출부에서도 체크하지만 안전장치 유지
    # Exact match (정규화 + 별칭 포함)
    for i, cell in enumerate(cell_map):
        is_match, is_exact, _ = _match_label(cell["text"], label)
        if is_match and is_exact:
            return i, False
    # Partial match
    if len(_normalize(label)) < 2:
        return None, False
    best_idx, best_ratio = None, 0
    for i, cell in enumerate(cell_map):
        is_match, is_exact, ratio = _match_label(cell["text"], label)
        if is_match and not is_exact and ratio > 0:
            if ratio > best_ratio:
                best_idx, best_ratio = i, ratio
    if best_idx is not None:
        return best_idx, True
    return None, False


def resolve_labels_to_tabs(hwp, table_idx, labels):
    """라벨 텍스트로 타겟 셀의 tab 인덱스를 찾는다.

    labels: [{"label": "계약금액", "text": "값", "direction": "right"|"below",
              "row_label": "전체기간" (optional)}, ...]

    로직:
    1. map_table_cells()로 전체 셀 맵 수집
    2. row_label이 있으면 → 2D 그리드 교차 매칭 (열 헤더 × 행 라벨)
    3. direction == "below"이면 → 행 그룹 기반 아래 셀 찾기
    4. 그 외(right) → 기존 tab+1 방식
    """
    cell_data = map_table_cells(hwp, table_idx)
    cell_map = cell_data.get("cell_map", [])

    if not cell_map:
        return {
            "resolved": [],
            "errors": ["표에서 셀을 찾을 수 없습니다."],
        }

    rows = _group_cells_into_rows(cell_map)
    resolved = []
    errors = []

    for item in labels:
        label = item.get("label", "").strip()
        text = item.get("text", "")
        direction = item.get("direction", "right")
        row_label = item.get("row_label", "").strip() if item.get("row_label") else ""

        if not label:
            errors.append("빈 라벨이 전달되었습니다.")
            continue

        if row_label:
            # ── 교차 매칭 모드: label(열 헤더) × row_label(행 라벨) ──
            if len(rows) <= 1:
                errors.append(
                    f"라벨 '{label}'+'{row_label}': 행 경계를 감지할 수 없습니다. "
                    "tab 인덱스를 직접 지정하세요."
                )
                continue

            all_texts = [c["text"][:20] for row in rows for c in row][:10]

            col_idx, header_row_idx, col_partial = _find_label_column(rows, label)
            if col_idx is None:
                errors.append(
                    f"열 라벨 '{label}'을(를) 표에서 찾을 수 없습니다. "
                    f"표 내 셀: {all_texts}"
                )
                continue

            target_row_idx, row_partial, row_matched_text = _find_label_row(rows, row_label)
            if target_row_idx is None:
                errors.append(
                    f"행 라벨 '{row_label}'을(를) 표에서 찾을 수 없습니다. "
                    f"표 내 셀: {all_texts}"
                )
                continue

            if col_partial:
                matched_cell = rows[header_row_idx][col_idx]
                print(f"[WARN] 열 라벨 '{label}' partial match: '{matched_cell['text'].strip()}'", file=sys.stderr)
            if row_partial:
                print(f"[WARN] 행 라벨 '{row_label}' partial match: '{row_matched_text}'", file=sys.stderr)

            target_row = rows[target_row_idx]
            if col_idx >= len(target_row):
                errors.append(
                    f"라벨 '{label}'+'{row_label}': 열 인덱스({col_idx})가 "
                    f"해당 행의 셀 수({len(target_row)})를 초과합니다."
                )
                continue

            target = target_row[col_idx]
            entry = {
                "tab": target["tab"],
                "text": text,
                "matched_label": f"{label}×{row_label}",
            }
            if col_partial or row_partial:
                entry["partial_match"] = True
            resolved.append(entry)

        elif direction == "below":
            # ── below 모드: 행 그룹 기반 아래 셀 찾기 ──
            matched_idx, is_partial = _find_cell_in_flat(cell_map, label)
            if matched_idx is None:
                errors.append(
                    f"라벨 '{label}'을(를) 표에서 찾을 수 없습니다. "
                    f"표 내 셀: {[c['text'][:20] for c in cell_map[:10]]}"
                )
                continue
            if is_partial:
                print(
                    f"[WARN] below 라벨 '{label}' partial match: "
                    f"'{cell_map[matched_idx]['text'].strip()}'",
                    file=sys.stderr,
                )

            if len(rows) <= 1:
                errors.append(
                    f"라벨 '{label}' (direction=below): 행 경계를 감지할 수 없어 "
                    "정확한 아래 셀을 찾을 수 없습니다. tab 인덱스를 직접 지정하세요."
                )
                continue
            else:
                # 행 그룹 기반: 같은 열의 다음 행 셀
                label_row_idx, col_idx = _find_cell_position_in_rows(rows, matched_idx)
                if label_row_idx is None:
                    errors.append(f"라벨 '{label}'의 행 위치를 결정할 수 없습니다.")
                    continue
                if label_row_idx + 1 >= len(rows):
                    errors.append(f"라벨 '{label}'의 아래 행이 없습니다.")
                    continue
                next_row = rows[label_row_idx + 1]
                if col_idx >= len(next_row):
                    errors.append(
                        f"라벨 '{label}': 아래 행의 셀 수({len(next_row)})가 "
                        f"열 인덱스({col_idx})보다 적습니다."
                    )
                    continue
                target = next_row[col_idx]
                entry = {
                    "tab": target["tab"],
                    "text": text,
                    "matched_label": label,
                }
                if is_partial:
                    entry["partial_match"] = True
                resolved.append(entry)

        else:
            # ── right 모드: 라벨의 다음 셀 (tab+1) ──
            matched_idx, is_partial = _find_cell_in_flat(cell_map, label)
            if matched_idx is None:
                errors.append(
                    f"라벨 '{label}'을(를) 표에서 찾을 수 없습니다. "
                    f"표 내 셀: {[c['text'][:20] for c in cell_map[:10]]}"
                )
                continue
            if is_partial:
                print(
                    f"[WARN] right 라벨 '{label}' partial match: "
                    f"'{cell_map[matched_idx]['text'].strip()}'",
                    file=sys.stderr,
                )

            target_idx = matched_idx + 1
            if target_idx >= len(cell_map):
                errors.append(
                    f"라벨 '{label}'의 오른쪽 셀이 없습니다 (표 범위 밖)."
                )
                continue

            entry = {
                "tab": cell_map[target_idx]["tab"],
                "text": text,
                "matched_label": label,
            }
            if is_partial:
                entry["partial_match"] = True
            resolved.append(entry)

    return {"resolved": resolved, "errors": errors}
