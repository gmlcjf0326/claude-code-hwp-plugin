"""hwp_core.formatting — MCP 서식 핸들러 (sub-package).

v0.7.9 Phase 6: formatting.py (629L) → hwp_core/formatting/ 분할

sub-modules (부작용 import 로 @register 데코레이터 실행):
- char_para.py       — get_char_shape, get_para_shape, set_paragraph_style (245L 메가)
- page_layout.py     — set_page_setup, set_column, set_cell_property, set_header_footer
- styles.py          — apply_style, apply_document_preset, apply_style_profile
- quick_actions.py   — toggle_checkbox, set_background_picture, set_cell_color,
                       set_table_border, auto_map_reference

총 15 @register 핸들러.
"""
from . import quick_actions   # noqa: F401 - @register 실행
from . import char_para       # noqa: F401
from . import page_layout     # noqa: F401
from . import styles          # noqa: F401
