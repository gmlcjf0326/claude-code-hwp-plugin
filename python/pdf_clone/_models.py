"""pdf_clone._models — Dataclasses + 상수.

Dataclasses:
- TextBlock    : 글자 span (bbox, font, size, bold, italic, color, page)
- Paragraph    : 논리적 단락 (여러 TextBlock 을 병합)
- TableModel   : 표 모델 (2D 셀 + has_header + bbox)
- PageLayout   : 페이지 레이아웃 (paragraphs + tables + images + rect)

상수:
- POINTS_TO_MM : 1 pt = 25.4/72 mm 변환
- MAX_IMG_W_MM : 이미지 최대 너비 (180mm)
- MAX_IMG_H_MM : 이미지 최대 높이 (240mm)
- MIN_IMG_MM   : 이미지 최소 크기 (10mm)
"""
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class TextBlock:
    """One span of text with bbox + font metadata."""
    text: str
    bbox: Tuple[float, float, float, float]  # (x0, y0, x1, y1) in PDF points
    font: str = ""
    size: float = 10.0                       # font size in points
    bold: bool = False
    italic: bool = False
    color: int = 0                           # RGB as int (0xRRGGBB)
    page: int = 0


@dataclass
class Paragraph:
    """A logical paragraph built from one or more TextBlocks (possibly multi-line)."""
    text: str
    font_size: float = 10.0
    bold: bool = False
    italic: bool = False
    align: str = "left"
    is_title: bool = False
    bbox: Optional[Tuple[float, float, float, float]] = None
    # v0.7.4.8 Fix Group A: visual fidelity fields
    color: int = 0                 # RGB int (0xRRGGBB) — dominant color across blocks
    font_name: str = ""            # Dominant font name across blocks (empty → 맑은 고딕 fallback)
    indent: float = 0.0            # First-line indent in pt (Fix Group D fills this)
    left_margin: float = 0.0       # Paragraph left margin in pt (Fix Group D fills this)
    line_spacing: int = 160        # Line spacing as percent (default 160%)
    space_before: float = 0.0      # Space before paragraph in pt


@dataclass
class TableModel:
    """Placeholder for v0.7.4.4 — table reconstruction."""
    cells_2d: List[List[str]] = field(default_factory=list)
    has_header: bool = False
    bbox: Optional[Tuple[float, float, float, float]] = None


@dataclass
class PageLayout:
    page_index: int
    paragraphs: List[Paragraph] = field(default_factory=list)
    tables: List[TableModel] = field(default_factory=list)
    images: List[Tuple[str, float, float]] = field(default_factory=list)  # (path, w_mm, h_mm)
    page_rect: Optional[Tuple[float, float]] = None  # (width_pt, height_pt)
    column_detected: bool = False  # v0.7.4.4: 2-column 추정
    list_paragraph_count: int = 0   # v0.7.4.8 Fix D4: bullet/번호 marker 감지된 단락 수


# v0.7.5.0: 1 pt = 25.4/72 mm (exact, was 0.3527777 magic constant)
POINTS_TO_MM = 25.4 / 72.0
MAX_IMG_W_MM = 180.0
MAX_IMG_H_MM = 240.0
MIN_IMG_MM = 10.0
