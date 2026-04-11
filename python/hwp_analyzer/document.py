"""hwp_analyzer.document — HWP 문서 전체 분석.

함수:
- analyze_document : 문서 열기 → 페이지/표/필드/텍스트 추출

v0.6.6+: HeadCtrl 순회로 controls 카탈로그 수집
v0.7.4.9 S2-NEW-1: 표 스캔 실패 시 continue (5회 연속 실패만 진짜 중단)
v0.7.9 fix: already_open=True 지원 (공유 위반 방지)
"""
import os
import sys

from ._constants import MAX_TABLES
from .label import classify_table_type


def analyze_document(hwp, file_path, already_open=False):
    """Analyze an HWP document: pages, tables, fields, text."""
    file_path = os.path.abspath(file_path)
    # already_open=True 이면 hwp.open 스킵 (공유 위반 방지, v0.7.9 fix)
    if not already_open:
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
        # B2 (v0.6.6): HeadCtrl 순회 결과 — 표/그림/머리말/꼬리말/각주/누름틀 위치
        "controls": [],
        "controls_by_type": {},
    }

    scan_started = False

    # B2 (v0.6.6): HeadCtrl 순회 — 표 추출 전에 전체 컨트롤 카탈로그 수집
    try:
        from hwp_traversal import traverse_all_ctrls
        ctrl_result = traverse_all_ctrls(hwp, include_ids=None)
        result["controls"] = ctrl_result.get("controls", [])
        result["controls_by_type"] = ctrl_result.get("by_type", {})
    except Exception as e:
        print(f"[WARN] HeadCtrl traversal failed: {e}", file=sys.stderr)

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
            consecutive_failures = 0  # v0.7.4.9 S2-NEW-1: 연속 실패 5회 초과 시만 진짜 중단
            consecutive_duplicates = 0  # v0.7.4.9: 연속 중복 2회 이상만 진짜 중단
            while table_idx < MAX_TABLES:
                try:
                    hwp.get_into_nth_table(table_idx)
                    df = hwp.table_to_df()
                    current_data = df.values.tolist()
                    # 중복 감지 (v0.7.4.9 완화): 연속 2회 같은 데이터일 때만 중단
                    if prev_data is not None and current_data == prev_data:
                        consecutive_duplicates += 1
                        print(
                            f"[INFO] Table {table_idx} is duplicate (count={consecutive_duplicates}) — "
                            f"{'stopping' if consecutive_duplicates >= 2 else 'skipping'}",
                            file=sys.stderr,
                        )
                        try:
                            if hwp.is_cell():
                                hwp.MovePos(3)
                            hwp.MovePos(2)
                        except Exception as e:
                            print(f"[WARN] Table exit (dup): {e}", file=sys.stderr)
                        if consecutive_duplicates >= 2:
                            break
                        table_idx += 1
                        continue
                    consecutive_duplicates = 0
                    consecutive_failures = 0  # 성공 시 counter 리셋
                    prev_data = current_data
                    table_info = {
                        "index": table_idx,
                        "rows": len(df) + 1,  # +1 for header
                        "cols": len(df.columns) if len(df) > 0 else 0,
                        "headers": [str(c) for c in df.columns],
                        "data": current_data,
                    }
                    table_info["table_type"] = classify_table_type(table_info)
                    result["tables"].append(table_info)
                    try:
                        if hwp.is_cell():
                            hwp.MovePos(3)
                        hwp.MovePos(2)
                    except Exception as e:
                        print(f"[WARN] Table exit/MovePos failed: {e}", file=sys.stderr)
                    table_idx += 1
                except Exception as e:
                    # v0.7.4.9 S2-NEW-1 Fix: break 대신 continue + 안전 장치
                    consecutive_failures += 1
                    print(
                        f"[INFO] Table {table_idx} scan failed (fails={consecutive_failures}): {e}",
                        file=sys.stderr,
                    )
                    # 상태 복구 시도
                    try:
                        if hwp.is_cell():
                            hwp.MovePos(3)
                        hwp.MovePos(2)
                    except Exception:
                        pass
                    if consecutive_failures >= 5:
                        print(
                            f"[WARN] 5 consecutive failures at table_idx {table_idx} — stopping",
                            file=sys.stderr,
                        )
                        break
                    table_idx += 1
                    continue
            # BUG-7 fix: 실제 발견된 표 수와 스캔 상한 분리
            result["scanned_tables"] = table_idx
            if table_idx >= MAX_TABLES:
                result["tables_truncated"] = True
                result["tables_truncated_message"] = (
                    f"표가 {MAX_TABLES}개 이상일 수 있습니다. 처음 {MAX_TABLES}개만 분석했습니다."
                )
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
