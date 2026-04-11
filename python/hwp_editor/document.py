"""hwp_editor.document — 대량 문서 채우기 오케스트레이터.

함수:
- fill_document : AI 가 생성한 fill_data 를 문서에 적용
                  (fields 채우기 + tables 채우기, tab/row-col 방식 혼용)
"""
import os
import sys

from .tables import fill_table_cells_by_tab


def fill_document(hwp, fill_data):
    """Fill document with AI-generated content.

    fill_data format:
    {
        "file_path": "...",           # optional: open file first
        "fields": {"name": "value"},  # field-based fill
        "tables": [                   # table-based fill
            {
                "index": 0,
                "cells": [
                    {"row": 0, "col": 0, "text": "value"},
                    ...
                ]
            }
        ]
    }
    """
    # Open file if specified
    if "file_path" in fill_data:
        file_path = os.path.abspath(fill_data["file_path"])
        hwp.open(file_path)

    result = {"filled": 0, "failed": 0, "errors": []}

    # Fill fields
    if "fields" in fill_data and fill_data["fields"]:
        try:
            hwp.put_field_text(fill_data["fields"])
            result["filled"] += len(fill_data["fields"])
        except Exception as e:
            result["errors"].append(f"Field fill failed: {e}")
            result["failed"] += len(fill_data["fields"])

    # Fill tables - each cell independently (re-enter table each time)
    if "tables" in fill_data:
        for table_data in fill_data["tables"]:
            table_idx = table_data.get("index", 0)
            cells = table_data.get("cells", [])

            # Split: tab-based cells vs row/col cells
            tab_cells = [c for c in cells if "tab" in c]
            rowcol_cells = [c for c in cells if "tab" not in c]

            if tab_cells:
                tab_result = fill_table_cells_by_tab(hwp, table_idx, tab_cells)
                result["filled"] += tab_result["filled"]
                result["failed"] += tab_result["failed"]
                result["errors"].extend(tab_result["errors"])

            for cell in rowcol_cells:
                try:
                    row = cell.get("row", 0)
                    col = cell.get("col", 0)
                    text = str(cell.get("text", cell.get("value", "")))

                    # Enter table fresh each time
                    hwp.get_into_nth_table(table_idx)

                    try:
                        # Navigate: first go down, then right
                        for _ in range(row):
                            hwp.HAction.Run("TableLowerCell")
                        for _ in range(col):
                            hwp.HAction.Run("TableRightCell")

                        # 선택 영역 대체 — 기존 서식 상속
                        hwp.HAction.Run("SelectAll")
                        hwp.insert_text(text)
                        result["filled"] += 1

                    finally:
                        try:
                            if hwp.is_cell():
                                hwp.MovePos(3)
                        except Exception as e:
                            print(f"[WARN] Table exit (fill_document): {e}", file=sys.stderr)

                except Exception as e:
                    result["failed"] += 1
                    result["errors"].append(
                        f"Table{table_idx} ({row},{col}) failed: {e}"
                    )
                    print(f"[WARN] Cell fill error: {e}", file=sys.stderr)

    return result
