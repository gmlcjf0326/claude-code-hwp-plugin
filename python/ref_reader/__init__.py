"""ref_reader — 외부 참고자료 리더 (multi-format).

지원 포맷: .txt, .md, .log, .csv, .xlsx, .xls, .json, .pdf, .html, .htm,
           .xml, .hwp, .hwpx, .docx, .doc, .pptx, .ppt, .rtf, .odt, .odp

공개 API (backward compat):
    from ref_reader import read_reference

v0.7.9 Phase 4: ref_reader.py (528L) → ref_reader/ 패키지 분할
  - dispatcher.py — _check_volume_warning, read_reference (라우터)
  - readers.py    — 1차 포맷 리더 (text/csv/excel/json/pdf/html/xml/hwp_structured)
  - conversion.py — 변환/fallback (LibreOffice, docx, pptx, via_pdf_conversion)
"""
from .dispatcher import read_reference, _check_volume_warning

__all__ = [
    "read_reference",
    "_check_volume_warning",
]
