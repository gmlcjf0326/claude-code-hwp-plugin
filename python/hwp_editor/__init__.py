"""hwp_editor — 저수준 HWP 문서 편집 유틸리티.

pyhwpx Hwp() 만 사용. raw win32com 금지.
파일 경로는 os.path.abspath() 필수.
셀 네비게이션은 TableRightCell 기반 sequential Tab traversal (병합 셀 안전).

공개 API (backward compat — 기존 `from hwp_editor import X` 무변경):
    - 텍스트 삽입: insert_text_with_color, insert_text_with_style, insert_markdown
    - 서식 조회: get_char_shape, get_para_shape, get_cell_format, get_table_format_summary
    - 서식 설정: set_paragraph_style
    - 표 채우기: fill_table_cells_by_tab, smart_fill_table_cells, fill_table_cells_by_label
    - 표 스타일: set_cell_background_color, set_table_border_style
    - 문서: fill_document, verify_after_fill
    - 컨텐츠: insert_picture, auto_map_reference_to_table, extract_all_text

v0.7.9 Phase 5: hwp_editor.py (1441L) → hwp_editor/ 패키지 분할
  - text_style.py       — 텍스트 삽입 + 문자/단락 서식 설정 (insert_text_with_*, set_paragraph_style)
  - char_para.py        — 문자/단락 서식 조회 (get_char_shape, get_para_shape, get_cell_format, get_table_format_summary)
  - tables.py           — 표 편집 (navigation, fill, 배경색, 테두리)
  - markdown_picture.py — 마크다운/이미지/참고자료 매핑/전체 텍스트 추출
  - document.py         — fill_document (대량 채우기 오케스트레이터)
"""
from .text_style import (
    insert_text_with_color,
    insert_text_with_style,
    set_paragraph_style,
)
from .char_para import (
    get_char_shape,
    get_para_shape,
    get_cell_format,
    get_table_format_summary,
)
from .tables import (
    _goto_cell,
    _navigate_to_tab,
    _hex_to_rgb,
    fill_table_cells_by_tab,
    smart_fill_table_cells,
    fill_table_cells_by_label,
    verify_after_fill,
    set_cell_background_color,
    set_table_border_style,
)
from .table_post_process import (  # v0.7.12 Phase 5E
    auto_font_size,
    auto_align,
    apply_auto_style,
    smart_fill_table_auto,
)
from .markdown_picture import (
    insert_markdown,
    insert_picture,
    auto_map_reference_to_table,
    extract_all_text,
)
from .document import fill_document

__all__ = [
    # 텍스트 삽입 + 서식 설정
    "insert_text_with_color",
    "insert_text_with_style",
    "set_paragraph_style",
    # 서식 조회
    "get_char_shape",
    "get_para_shape",
    "get_cell_format",
    "get_table_format_summary",
    # 표 네비게이션 + 채우기
    "_goto_cell",
    "_navigate_to_tab",
    "_hex_to_rgb",
    "fill_table_cells_by_tab",
    "smart_fill_table_cells",
    "fill_table_cells_by_label",
    "verify_after_fill",
    "set_cell_background_color",
    "set_table_border_style",
    # 컨텐츠
    "insert_markdown",
    "insert_picture",
    "auto_map_reference_to_table",
    "extract_all_text",
    # 문서 채우기
    "fill_document",
]
