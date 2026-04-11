"""hwp_analyzer.tables — 표 셀 매핑 + 라벨 해결.

함수:
- map_table_cells             : 표의 모든 셀을 Tab traversal 로 순회 → cell_map
- _group_cells_into_rows      : flat cell_map → 행 단위 그룹
- _find_label_column          : 열 헤더 라벨 매칭
- _find_label_row             : 행 라벨 매칭
- _find_cell_position_in_rows : flat index → (row, col) 변환
- _find_cell_in_flat          : flat cell_map 에서 라벨 검색
- resolve_labels_to_tabs      : 라벨 리스트 → tab 인덱스 해결 (right/below/cross matching)
"""
import sys

from .label import _normalize, _match_label


def map_table_cells(hwp, table_idx, max_cells=200, max_text_len=500):
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

            # v0.7.4.8 Fix B3: 100자 → 500자 (파라미터화). 긴 사업내용 등 보존
            truncated_text = cell_text[:max_text_len]
            cell_map.append({
                "tab": i,
                "text": truncated_text,
                "text_truncated": len(cell_text) > max_text_len,
                "original_length": len(cell_text),
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
        return None, False  # 이중 방어
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
