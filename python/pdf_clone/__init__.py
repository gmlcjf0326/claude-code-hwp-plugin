"""pdf_clone — PDF → HWP 복제 파이프라인.

v0.7.4.2: native PDF only (pdfplumber + PyMuPDF get_text dict)
v0.7.4.3: + PaddleOCR for scanned PDFs, preprocessing, title detection, hybrid per-page dispatch
v0.7.4.4: + pdfplumber find_tables, image extraction, fidelity scoring, column detection warnings
v0.7.4.8: Fix Groups A (visual fidelity) + D (structure detection)
v0.7.4.9: PaddleOCR 3.x 호환 + paragraph overlap filter
v0.7.9 Phase 4: pdf_clone.py (1211L) → pdf_clone/ 패키지 분할

공개 API (backward compat):
    from pdf_clone import clone_pdf_to_hwp

내부 구조:
- `_models.py`  — TextBlock, Paragraph, TableModel, PageLayout dataclasses + 상수
- `ocr.py`      — PaddleOCR 파이프라인 (scanned/hybrid PDF)
- `native.py`   — PyMuPDF/pdfplumber native 추출 (text, tables, images, headers/footers)
- `layout.py`   — 레이아웃 분석 + HWP emit + fidelity 점수 + clone_pdf_to_hwp 메인 진입
"""
from .layout import clone_pdf_to_hwp
from ._models import TextBlock, Paragraph, TableModel, PageLayout

__all__ = [
    "clone_pdf_to_hwp",
    # Dataclasses (backward compat — 일부 테스트가 직접 import)
    "TextBlock",
    "Paragraph",
    "TableModel",
    "PageLayout",
]
