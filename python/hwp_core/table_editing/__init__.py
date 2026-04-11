"""hwp_core.table_editing — MCP 표 편집 핸들러 (sub-package).

v0.7.9 Phase 7: table_editing.py (720L) → hwp_core/table_editing/ 분할

sub-modules (부작용 import 로 @register 실행):
- queries.py          — get_table_dimensions, get_cell_format, get_table_format_summary,
                        smart_fill, read_reference (5 handlers)
- navigation.py       — enter_table, exit_table, navigate_cell, insert_row_at_cursor,
                        merge_current_selection (5 handlers)
- structure.py        — table_add_row, table_delete_row, table_add_column, table_delete_column,
                        table_merge_cells, table_split_cell (6 handlers)
- creation.py         — table_create_from_data, create_approval_box, table_insert_from_csv (3 handlers)
- formulas_export.py  — table_formula_sum, table_formula_avg, table_to_csv, table_to_json,
                        table_swap_type, table_distribute_width (6 handlers)

총 25 @register 핸들러.
"""
from . import queries            # noqa: F401
from . import navigation         # noqa: F401
from . import structure          # noqa: F401
from . import creation           # noqa: F401
from . import formulas_export    # noqa: F401
