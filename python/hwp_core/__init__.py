"""HWP Core — Modular HWP operations.

v0.7.6.0 모듈화: hwp_service.py 의 109 RPC 메서드를 기능별 모듈로 분리.

구조:
    hwp_core/
        __init__.py          — REGISTRY + module import
        _helpers.py          — 공유 helper (validate_params, _exit_table_safely 등)
        _state.py            — module-level state (_current_doc_path)
        text_editing.py      — 텍스트 편집 (P1-5)
        table_editing.py     — 표 편집 (P1-5)
        formatting.py        — 서식/스타일 (P1-4)
        analysis.py          — 분석/감지 (P1-4)
        content.py           — 콘텐츠 삽입 (P1-5)
        document.py          — 문서 관리 (P1-3)
        utility.py           — 유틸/프리셋 (P1-3)

dispatcher 패턴:
    hwp_service.py dispatch() 가 REGISTRY.get(method) 로 O(1) lookup.
    miss 시 기존 if-elif fallback (P1 완료 전까지).
"""

# Registry: method_name -> handler(hwp, params) -> dict
REGISTRY: dict = {}


def register(method_name: str):
    """Handler 등록 데코레이터.
    사용: @register("insert_text")
         def insert_text_handler(hwp, params): ...
    """
    def decorator(func):
        REGISTRY[method_name] = func
        return func
    return decorator


# v0.7.6.0 P1-3/4: 모듈 import 시 @register 데코레이터가 REGISTRY 자동 채움
# 각 모듈 import 순서는 중요하지 않음 (dispatch 시 lookup 만 하므로)
from . import utility   # noqa: F401, E402  — ping/get_font_list/get_preset_list
from . import document  # noqa: F401, E402  — get_document_info/document_new/analyze/fill
from . import analysis    # noqa: F401, E402  — get_page_setup/get_cursor_context/extract_style/form_detect/extract_template_structure
from . import formatting  # noqa: F401, E402  — toggle_checkbox/set_background_picture/set_cell_color/set_table_border/auto_map_reference
from . import content     # noqa: F401, E402  — 18 methods (insert_page_break, break_section, break_column, insert_date_code, insert_auto_num, insert_memo, insert_line, insert_caption, insert_hyperlink, insert_footnote, insert_endnote, insert_markdown, insert_picture, privacy_scan, list_controls, word_count, indent, outdent)
from . import text_editing  # noqa: F401, E402  — 9 big methods (insert_text, insert_body_after_heading, insert_heading, text_search, find_replace, find_replace_multi, find_and_append, find_replace_nth, extend_section)
from . import table_editing # noqa: F401, E402  — 24 methods (table_create_from_data, enter/exit_table, merge/split, navigate, add/delete row/column, formula, to_csv/json, swap, distribute, insert_from_csv, create_approval_box)
from . import content_gen   # noqa: F401, E402  — v0.7.8 map_reference_to_sections, build_section_context
from . import form_fill_auto  # noqa: F401, E402  — v0.8.0 Phase 5G (form_fill_auto)

