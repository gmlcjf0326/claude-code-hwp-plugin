"""참고자료 텍스트 추출기.
지원: .txt, .csv, .xlsx, .json, .md
HWP/HWPX는 hwp_analyzer.analyze_document 사용 (이 모듈에서는 다루지 않음)
"""
import os
import json


def read_reference(file_path, max_chars=30000):
    """참고자료 파일에서 텍스트 추출."""
    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()

    if ext in ('.txt', '.md', '.log'):
        return _read_text(file_path, max_chars)
    elif ext == '.csv':
        return _read_csv(file_path, max_chars)
    elif ext in ('.xlsx', '.xls'):
        return _read_excel(file_path, max_chars)
    elif ext == '.json':
        return _read_json(file_path, max_chars)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}. 지원: .txt, .md, .csv, .xlsx, .json")


def _read_text(path, max_chars):
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read(max_chars)
    return {
        "format": "text",
        "file_name": os.path.basename(path),
        "content": content,
        "char_count": len(content),
    }


def _read_csv(path, max_chars):
    import csv
    rows = []
    total_chars = 0
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        reader = csv.reader(f)
        for row in reader:
            row_text = ','.join(row)
            total_chars += len(row_text)
            if total_chars > max_chars:
                break
            rows.append(row)

    headers = rows[0] if rows else []
    data = rows[1:] if len(rows) > 1 else []
    return {
        "format": "csv",
        "file_name": os.path.basename(path),
        "headers": headers,
        "data": data,
        "row_count": len(data),
    }


def _read_excel(path, max_chars):
    try:
        import openpyxl
    except ImportError:
        raise ImportError("openpyxl이 필요합니다. pip install openpyxl")

    wb = None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheets = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = []
            total_chars = 0
            for row in ws.iter_rows(values_only=True):
                row_data = [str(cell) if cell is not None else "" for cell in row]
                total_chars += sum(len(c) for c in row_data)
                if total_chars > max_chars:
                    break
                rows.append(row_data)

            headers = rows[0] if rows else []
            data = rows[1:] if len(rows) > 1 else []
            sheets.append({
                "sheet_name": sheet_name,
                "headers": headers,
                "data": data,
                "row_count": len(data),
            })

        return {
            "format": "excel",
            "file_name": os.path.basename(path),
            "sheets": sheets,
            "sheet_count": len(sheets),
        }
    finally:
        if wb:
            wb.close()


def _read_json(path, max_chars):
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read(max_chars)

    data = json.loads(content)
    return {
        "format": "json",
        "file_name": os.path.basename(path),
        "data": data,
    }
