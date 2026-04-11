"""hwp_core.text_editing — MCP 텍스트 편집 핸들러 (sub-package).

v0.7.9 Phase 7: text_editing.py (1133L) → hwp_core/text_editing/ 분할

sub-modules (부작용 import 로 @register 실행):
- _internal.py   — 비공개 helpers (dialog auto-dismiss + heading detection + indent helpers)
                   _find_hwp_confirm_dialog, _auto_dismiss_hwp_dialog, _with_auto_dismiss,
                   _find_heading_positions, _find_all_in_normalized, _extract_match_from_line,
                   _find_by_regex, _find_core_in_lines, _apply_indent_at_caret, _detect_heading_depth
- search.py      — 검색/치환 (4 handlers)
                   text_search, find_replace, find_replace_multi, find_replace_nth
- insertions.py  — 텍스트/제목/본문 삽입 (5 handlers)
                   insert_text, insert_heading, insert_body_after_heading, find_and_append, extend_section

총 9 @register 핸들러. 모든 sub-module 은 `from .. import register` (두 점).
"""
# _internal.py 는 @register 없으므로 먼저 import 할 필요 없지만, 명시적으로 import
from . import _internal              # noqa: F401
from . import search                  # noqa: F401
from . import insertions              # noqa: F401
from . import placeholder_cleanup     # noqa: F401 — v0.7.11 Phase 5C (cleanup_all_placeholders)
