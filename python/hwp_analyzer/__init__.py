"""hwp_analyzer — HWP 문서 분석 + 라벨 매칭 + 표 타입 분류.

공개 API (backward compat):
    from hwp_analyzer import analyze_document, map_table_cells, classify_table_type
    from hwp_analyzer import _match_label, _canonical_label, _normalize, resolve_labels_to_tabs
    from hwp_analyzer import _LABEL_ALIASES, _ALIAS_LOOKUP, _TABLE_TYPE_KEYWORDS

v0.7.9 Phase 5: hwp_analyzer.py (705L) → hwp_analyzer/ 패키지 분할
  - _constants.py — _LABEL_ALIASES (30+ canonicals, 200+ aliases), _TABLE_TYPE_KEYWORDS, MAX_TABLES
  - label.py      — _normalize, _canonical_label, _match_label, classify_table_type
  - document.py   — analyze_document (179L 메가)
  - tables.py     — map_table_cells, _group_cells_into_rows, _find_label_*, resolve_labels_to_tabs
"""
from ._constants import _LABEL_ALIASES, _ALIAS_LOOKUP, _TABLE_TYPE_KEYWORDS, MAX_TABLES
from .label import _normalize, _canonical_label, _match_label, classify_table_type
from .document import analyze_document
from .tables import (
    map_table_cells,
    _group_cells_into_rows,
    _find_label_column,
    _find_label_row,
    _find_cell_position_in_rows,
    _find_cell_in_flat,
    resolve_labels_to_tabs,
)

__all__ = [
    # 상수
    "MAX_TABLES",
    "_LABEL_ALIASES",
    "_ALIAS_LOOKUP",
    "_TABLE_TYPE_KEYWORDS",
    # 라벨 유틸
    "_normalize",
    "_canonical_label",
    "_match_label",
    "classify_table_type",
    # 문서 분석
    "analyze_document",
    # 표 셀 매핑
    "map_table_cells",
    "_group_cells_into_rows",
    "_find_label_column",
    "_find_label_row",
    "_find_cell_position_in_rows",
    "_find_cell_in_flat",
    "resolve_labels_to_tabs",
]
