# Python Backend Index — HWP Studio MCP Server

> 📖 **AI 에이전트 우선 참조 문서**: 코드 작업 전 이 파일을 먼저 읽으세요.
> 목표 기능을 여기서 찾은 다음, 해당 파일만 열어서 작업하세요.
> **전체 파일 스캔 금지** — 인덱스 → 대상 파일 1-2개만 읽으세요.
>
> **생성**: Phase 0 (2026-04-11) · **완성**: Phase 8 (2026-04-11)

---

## 1. 의도별 빠른 찾기 (Quick navigation by intent)

| 하고 싶은 일 | 파일 | 진입 함수 |
|---|---|---|
| 텍스트 삽입 (서식 포함) | `hwp_core/text_editing/insertions.py` | `insert_text` |
| 제목 삽입 (numbering + OutlineLevel) | `hwp_core/text_editing/insertions.py` | `insert_heading` |
| 제목 아래 본문 삽입 | `hwp_core/text_editing/insertions.py` | `insert_body_after_heading` |
| 텍스트 검색 | `hwp_core/text_editing/search.py` | `text_search` |
| Find/Replace | `hwp_core/text_editing/search.py` | `find_replace`, `find_replace_multi`, `find_replace_nth` |
| 문자 서식 조회 | `hwp_editor/char_para.py` | `get_char_shape` |
| 단락 서식 조회 | `hwp_editor/char_para.py` | `get_para_shape` |
| 단락 서식 설정 | `hwp_editor/text_style.py` | `set_paragraph_style` |
| 스타일 지정 텍스트 | `hwp_editor/text_style.py` | `insert_text_with_style` |
| 페이지 설정 | `hwp_core/formatting/page_layout.py` | `set_page_setup` |
| 문자/단락 서식 일괄 적용 | `hwp_core/formatting/styles.py` | `apply_style_profile` |
| 표 생성 (데이터) | `hwp_core/table_editing/creation.py` | `table_create_from_data` |
| 표 셀 매핑 (라벨 → 위치) | `hwp_analyzer/tables.py` | `map_table_cells` |
| 표 셀 채우기 (tab 인덱스) | `hwp_editor/tables.py` | `fill_table_cells_by_tab` |
| 표 셀 배경색 | `hwp_editor/tables.py` | `set_cell_background_color` |
| 표 진입/탈출 | `hwp_core/table_editing/navigation.py` | `enter_table`, `exit_table` |
| 셀 병합/분할 | `hwp_core/table_editing/structure.py` | `table_merge_cells`, `table_split_cell` |
| 문서 분석 (전체) | `hwp_analyzer/document.py` | `analyze_document` |
| 문서 타입 자동 감지 | `hwp_core/analysis/detection.py` | `detect_document_type` |
| 작업량 추정 (모델 추천) | `hwp_core/analysis/detection.py` | `estimate_workload` |
| 5-단계 검증 | `hwp_core/analysis/verification.py` | `verify_5stage` |
| 스타일 프로파일 추출 | `hwp_core/analysis/profile.py` | `extract_style_profile` |
| 양식 템플릿 구조 추출 | `hwp_core/analysis/profile.py` | `extract_template_structure` |
| 양식 원본 스냅샷 | `hwp_core/analysis/profile.py` | `snapshot_template_style` |
| 양식 빈칸 자동 감지 | `hwp_core/analysis/detection.py` | `form_detect` |
| 라벨 alias 사전 | `hwp_analyzer/_constants.py` | `_LABEL_ALIASES` |
| 라벨 매칭 (fuzzy) | `hwp_analyzer/label.py` | `_match_label` |
| 표 타입 자동 분류 | `hwp_analyzer/label.py` | `classify_table_type` |
| 작성요령 추출 | `hwp_core/content.py` (§ GUIDE TEXT) | `extract_guide_text` |
| 작성요령 삭제 | `hwp_core/content.py` (§ GUIDE TEXT) | `delete_guide_text` |
| 마크다운 삽입 | `hwp_editor/markdown_picture.py` | `insert_markdown` |
| 이미지 삽입 | `hwp_editor/markdown_picture.py` | `insert_picture` |
| 참고자료 읽기 (Excel/CSV/PDF) | `ref_reader/dispatcher.py` | `read_reference` |
| PDF → HWP 변환 | `pdf_clone/layout.py` | `clone_pdf_to_hwp` |
| 섹션 매핑 (참고자료 → 섹션) | `hwp_core/content_gen.py` | `map_reference_to_sections` |
| 섹션 컨텍스트 빌드 | `hwp_core/content_gen.py` | `build_section_context` |
| 문서 열기 | `hwp_core/document.py` | `open_document` |
| 문서 저장 | `hwp_core/document.py` | `save_document` |
| 텍스트 전체 추출 | `hwp_editor/markdown_picture.py` | `extract_all_text` |

---

## 2. REGISTRY 메서드 색인 (MCP tools, 112개 A-Z)

> `@register("method_name")` 으로 등록된 모든 핸들러. 모듈 경로는 `hwp_core/` 이하 상대 경로.
> Phase 8 에서 live REGISTRY 로부터 생성 — 최신 상태와 일치.

| method | 모듈 |
|---|---|
| analyze_document | document |
| analyze_writing_patterns | analysis/detection.py |
| apply_document_preset | formatting/styles.py |
| apply_style | formatting/styles.py |
| apply_style_profile | formatting/styles.py |
| auto_map_reference | formatting/quick_actions.py |
| batch_convert | document |
| break_column | content |
| break_section | content |
| build_section_context | content_gen |
| clone_pdf_to_hwp | content |
| close_document | document |
| compare_documents | document |
| create_approval_box | table_editing/creation.py |
| create_gantt_chart | content |
| delete_guide_text | content |
| detect_document_type | analysis/detection.py |
| document_merge | document |
| document_new | document |
| document_split | document |
| draw_line | content |
| enter_table | table_editing/navigation.py |
| estimate_workload | analysis/detection.py |
| exit_table | table_editing/navigation.py |
| export_format | document |
| extend_section | text_editing/insertions.py |
| extract_full_profile | analysis/profile.py |
| extract_guide_text | content |
| extract_style_profile | analysis/profile.py |
| extract_template_structure | analysis/profile.py |
| fill_by_label | document |
| fill_by_tab | document |
| fill_document | document |
| find_and_append | text_editing/insertions.py |
| find_replace | text_editing/search.py |
| find_replace_multi | text_editing/search.py |
| find_replace_nth | text_editing/search.py |
| form_detect | analysis/detection.py |
| generate_toc | content |
| get_cell_format | table_editing/queries.py |
| get_char_shape | formatting/char_para.py |
| get_cursor_context | analysis/metadata.py |
| get_document_info | document |
| get_font_list | utility |
| get_page_setup | analysis/metadata.py |
| get_para_shape | formatting/char_para.py |
| get_preset_list | utility |
| get_selected_text | document |
| get_table_dimensions | table_editing/queries.py |
| get_table_format_summary | table_editing/queries.py |
| image_extract | content |
| indent | content |
| insert_auto_num | content |
| insert_body_after_heading | text_editing/insertions.py |
| insert_caption | content |
| insert_date_code | content |
| insert_endnote | content |
| insert_footnote | content |
| insert_heading | text_editing/insertions.py |
| insert_hyperlink | content |
| insert_line | content |
| insert_markdown | content |
| insert_memo | content |
| insert_page_break | content |
| insert_page_num | content |
| insert_picture | content |
| insert_row_at_cursor | table_editing/navigation.py |
| insert_text | text_editing/insertions.py |
| insert_textbox | content |
| list_controls | content |
| map_reference_to_sections | content_gen |
| map_table_cells | document |
| merge_current_selection | table_editing/navigation.py |
| navigate_cell | table_editing/navigation.py |
| open_document | document |
| outdent | content |
| ping | utility |
| privacy_scan | content |
| read_reference | table_editing/queries.py |
| save_as | document |
| save_document | document |
| set_background_picture | formatting/quick_actions.py |
| set_cell_color | formatting/quick_actions.py |
| set_cell_property | formatting/page_layout.py |
| set_column | formatting/page_layout.py |
| set_header_footer | formatting/page_layout.py |
| set_page_setup | formatting/page_layout.py |
| set_paragraph_style | formatting/char_para.py |
| set_table_border | formatting/quick_actions.py |
| smart_fill | table_editing/queries.py |
| snapshot_template_style | analysis/profile.py |
| table_add_column | table_editing/structure.py |
| table_add_row | table_editing/structure.py |
| table_create_from_data | table_editing/creation.py |
| table_delete_column | table_editing/structure.py |
| table_delete_row | table_editing/structure.py |
| table_distribute_width | table_editing/formulas_export.py |
| table_formula_avg | table_editing/formulas_export.py |
| table_formula_sum | table_editing/formulas_export.py |
| table_insert_from_csv | table_editing/creation.py |
| table_merge_cells | table_editing/structure.py |
| table_split_cell | table_editing/structure.py |
| table_swap_type | table_editing/formulas_export.py |
| table_to_csv | table_editing/formulas_export.py |
| table_to_json | table_editing/formulas_export.py |
| text_search | text_editing/search.py |
| toggle_checkbox | formatting/quick_actions.py |
| validate_consistency | analysis/verification.py |
| verify_5stage | analysis/verification.py |
| verify_after_fill | content |
| verify_layout | analysis/verification.py |
| word_count | content |

**총 112개** (v0.7.9 기준). `from hwp_core import REGISTRY; len(REGISTRY) == 112` 불변.

---

## 3. 모듈 안내 (언제 어느 모듈을 읽어야 하는가)

### hwp_core/ (MCP 핸들러 — REGISTRY 기반)

**언제**: MCP tool 호출 처리 로직을 볼 때. 파라미터/응답 스키마 확인. 핸들러 내부 디버깅.

**구조**:

#### `hwp_core/text_editing/` — 텍스트 편집 (9 handlers)
- `search.py` — text_search, find_replace, find_replace_multi, find_replace_nth
- `insertions.py` — insert_text, insert_heading, **insert_body_after_heading (354L 메가)**, extend_section, find_and_append
- `_internal.py` — 비공개 helpers: dialog auto-dismiss (Windows API ctypes), 3-tier fuzzy heading matcher, 마커/번호 감지

#### `hwp_core/analysis/` — 분석 + 감지 (13 handlers)
- `metadata.py` — get_page_setup, get_cursor_context
- `profile.py` — extract_style_profile, extract_full_profile, extract_template_structure, snapshot_template_style
- `verification.py` — verify_5stage, verify_layout, validate_consistency
- `detection.py` — detect_document_type, analyze_writing_patterns, estimate_workload, form_detect

#### `hwp_core/formatting/` — 서식 (15 handlers)
- `char_para.py` — get_char_shape, get_para_shape, **set_paragraph_style (245L 메가)**
- `page_layout.py` — set_page_setup, set_column, set_cell_property, set_header_footer
- `styles.py` — apply_style, apply_document_preset, apply_style_profile
- `quick_actions.py` — toggle_checkbox, set_background_picture, set_cell_color, set_table_border, auto_map_reference

#### `hwp_core/table_editing/` — 표 편집 (25 handlers)
- `queries.py` — get_table_dimensions, get_cell_format, get_table_format_summary, smart_fill, read_reference
- `navigation.py` — enter_table, exit_table, navigate_cell, insert_row_at_cursor, merge_current_selection
- `structure.py` — table_add/delete_row/column, table_merge_cells, table_split_cell
- `creation.py` — table_create_from_data, create_approval_box, table_insert_from_csv
- `formulas_export.py` — table_formula_sum/avg, table_to_csv/json, table_swap_type, table_distribute_width

#### `hwp_core/content.py` — 콘텐츠 삽입 (28 handlers, **분할 안 함**, 섹션 주석으로 구분)

섹션 anchor (ctrl+F 검색 가능):
- `§ BASIC INSERTIONS` — line 21~163 (13 handlers): insert_page_break, break_section, break_column, insert_date_code, insert_auto_num, insert_memo, insert_line, insert_caption, insert_hyperlink, insert_footnote, insert_endnote, insert_markdown, insert_picture
- `§ METRICS & SCAN` — line 183~227 (3 handlers): privacy_scan, list_controls, word_count
- `§ INDENT / OUTDENT` — line 230~248 + 817~ (2 handlers): indent, outdent
- `§ COMPLEX INSERTIONS` — line 259~472 (6 handlers): insert_textbox, draw_line, image_extract, insert_page_num, generate_toc, create_gantt_chart
- `§ GUIDE TEXT (v0.7.7)` — line 505~750 (3 handlers): extract_guide_text, delete_guide_text, verify_after_fill
- `§ PDF CLONE` — line 764~815 (1 handler): clone_pdf_to_hwp (pdf_clone 패키지 위임)

#### `hwp_core/document.py` — 문서 라이프사이클 (17 handlers)
- open_document, save_document, save_as, close_document, get_document_info, document_new, document_merge, document_split, export_format, batch_convert, compare_documents, analyze_document, fill_document, fill_by_tab, fill_by_label, get_selected_text, map_table_cells

#### `hwp_core/content_gen.py` — 참고자료 → 섹션 매핑 (2 handlers, v0.7.8)
- map_reference_to_sections, build_section_context

#### `hwp_core/utility.py` — 유틸리티 (3 handlers)
- ping, get_font_list, get_preset_list

#### `hwp_core/_helpers.py` — 공유 유틸 (핸들러 없음, 모든 sub-module 이 import)
- `validate_params(params, required_keys, method_name)` — 파라미터 검증
- `normalize_unicode(text)`, `normalize_for_match(text)`, `normalize_for_display(text)` — v0.7.7 정규화
- `_exit_table_safely(hwp)` — 표 안전 탈출 (MovePos(3))
- `_execute_all_replace(hwp, find, replace, ...)` — AllReplace + 검증
- `validate_file_path(file_path, must_exist)` — 심링크 거부 + 권한 체크

#### `hwp_core/_state.py` — module-level state
- `_current_doc_path` + getter/setter/clear (single source of truth)

#### `hwp_core/__init__.py` — REGISTRY + @register 데코레이터
```python
REGISTRY: dict = {}
def register(method_name: str):
    def decorator(func):
        REGISTRY[method_name] = func
        return func
    return decorator
# 모든 sub-module 부작용 import → @register 데코레이터가 REGISTRY 자동 채움
from . import utility, document, analysis, formatting, content, text_editing, table_editing, content_gen
```

### hwp_editor/ (저수준 HWP 조작 유틸리티)

**언제**: pyhwpx COM API 직접 조작이 필요할 때. hwp_core/ 핸들러들이 lazy import 로 호출.

**구조**:
- `__init__.py` — 모든 공개 심볼 re-export (backward compat — `from hwp_editor import X`)
- `char_para.py` — get_char_shape, get_para_shape, get_cell_format, get_table_format_summary
- `text_style.py` — insert_text_with_color, insert_text_with_style, set_paragraph_style
- `tables.py` — fill_table_cells_by_tab, smart_fill_table_cells, fill_table_cells_by_label, set_cell_background_color, set_table_border_style, _goto_cell, _navigate_to_tab, verify_after_fill, _hex_to_rgb
- `markdown_picture.py` — insert_markdown, insert_picture, auto_map_reference_to_table, extract_all_text
- `document.py` — fill_document (대량 채우기 오케스트레이터)

### hwp_analyzer/ (문서 분석 + 라벨 매칭 유틸)

**언제**: 표 라벨 매칭, 표 타입 분류, 문서 전체 구조 분석이 필요할 때.

**구조**:
- `__init__.py` — 공개 API re-export
- `_constants.py` — `_LABEL_ALIASES` (30+ 정규 라벨 → 200+ 별칭), `_TABLE_TYPE_KEYWORDS`, MAX_TABLES
- `label.py` — _normalize, _canonical_label, _match_label, classify_table_type
- `document.py` — analyze_document (전체 문서 트래버스)
- `tables.py` — map_table_cells, _group_cells_into_rows, _find_label_column, _find_label_row, _find_cell_position_in_rows, _find_cell_in_flat, resolve_labels_to_tabs

### pdf_clone/ (PDF → HWP 변환 파이프라인)

**언제**: PDF 를 HWP 로 복제할 때. `clone_pdf_to_hwp` 가 유일한 공개 진입점. `hwp_core/content.py` 에서 lazy import.

**구조**:
- `__init__.py` — re-export clone_pdf_to_hwp + dataclasses
- `_models.py` — TextBlock, Paragraph, TableModel, PageLayout dataclasses + 상수 (POINTS_TO_MM, MAX_IMG_*)
- `ocr.py` — PaddleOCR 엔진, scanned PDF 처리, _detect_pdf_type, _get_ocr_singleton, _preprocess_for_ocr, _extract_ocr_blocks
- `native.py` — PyMuPDF native 추출, _extract_native_blocks, _make_paragraph, _detect_list_markers, _detect_tables_native, _extract_images_native, _extract_headers_footers
- `layout.py` — _layout_analyze, _emit_layout_to_hwp, _compute_fidelity_score, **clone_pdf_to_hwp (358L 메가)**

### ref_reader/ (외부 참고자료 리더)

**언제**: Excel/CSV/PDF/DOCX/PPTX 등 외부 참고자료 파일 읽을 때.

**구조**:
- `__init__.py` — re-export read_reference
- `dispatcher.py` — _check_volume_warning (4단계 🟢🟡🟠🔴), read_reference (format 라우터)
- `readers.py` — _read_text, _read_csv, _read_excel, _read_json, _read_pdf, _read_html, _read_xml, _read_hwp_structured
- `conversion.py` — _read_via_pdf_conversion, _convert_to_pdf_libreoffice, _read_docx_direct, _read_pptx_direct

### 단일 파일 유틸 (분할 안 함)

- `hwp_service.py` (403L) — MCP dispatcher. REGISTRY O(1) lookup + JSON-RPC 응답. v0.7.6.0 에서 89% 축소.
- `hwp_constants.py` (162L) — MSG_BOX_MODE, SCAN_FLAG 등 상수 + `scan_context` 컨텍스트 매니저
- `presets.py` (342L) — 스타일 프리셋 사전 (DOCUMENT_PRESETS, get_korean_business_default)
- `hwp_traversal.py` (223L) — 문서 트래버설 유틸 (HeadCtrl 순회)
- `hwpx_reader.py` (170L) — HWPX XML 직접 읽기 (COM 우회, v0.7.2.13)
- `privacy_scanner.py` (115L) — PII 탐지 (SSN/전화/이메일/계좌)

---

## 4. Dispatch Flow (작업 이해)

```
Claude Code (사용자)
    ↓ JSON-RPC via stdin
Node MCP server (claude-code-hwp-plugin/servers/bundle.mjs)
    ↓ child_process → Python stdin
hwp_service.py: dispatch(hwp, method, params)
    ↓ REGISTRY.get(method)  [O(1) lookup]
hwp_core/{sub-package}/{module}.py: @register("{method}") handler(hwp, params)
    ↓ lazy import + pyhwpx call
hwp_editor/, hwp_analyzer/, pdf_clone/, ref_reader/
    ↓ pyhwpx COM bridge
한글 프로그램 Hwp.exe (사용자 세션)
    ↓ response JSON
    ↑ (같은 경로 역방향)
```

**핵심 결정**:
- REGISTRY 는 `hwp_core/__init__.py:24` 의 dict
- `@register(method_name)` 데코레이터는 import time 에 REGISTRY 자동 등록
- `hwp_core/__init__.py` 는 모든 sub-package 부작용 import → 전 handler 가 dispatch 가능
- Cross-module 의존은 **함수 본문 내 lazy import** → 순환 없음

---

## 5. 작업 시 주의 사항

### 편집 대상
- **항상 `mcp-server/python/` 이 canonical**. `claude-code-hwp-plugin/python/` 은 build.mjs 가 덮어쓰는 산출물 (편집 금지).
- Python 스크립트 실행 시 `mcp-server/python/` **절대 경로** 사용.

### @register 추가 시
- 새 handler 는 해당 sub-package 의 sub-module 에 추가
- **hwp_core 의 sub-package (analysis/, formatting/, table_editing/, text_editing/) 의 경우**: `from .. import register` (두 점!)
- **최상위 hwp_core 파일 (content.py, document.py, utility.py 등) 의 경우**: `from . import register` (한 점)
- `_helpers`, `_state` 임포트도 같은 규칙 적용

### Import 규칙
- hwp_core 의 cross-module (예: analysis.py → hwp_editor): **함수 본문 내 lazy import** 유지 (순환 방지)
- `hwp_core/__init__.py` 만 top-level import (REGISTRY 등록용)
- `from <package> import *` 사용 금지 — 명시적 이름 지정

### 검증 (편집 후 반드시)
```bash
py -3.13 -c "import sys; sys.path.insert(0, '.'); from hwp_core import REGISTRY; assert len(REGISTRY) == 112"
py -3.13 -c "import sys; sys.path.insert(0, '.'); import hwp_service; print('ok')"
py -3.13 mcp-server/scripts/verify_python_index.py
```

---

## 6. 메가 함수 (Phase 10 DEFERRED — 내부 리팩터링 예정)

아래 함수들은 현재 분할 완료된 sub-module 안에 있지만, 단일 함수가 200줄 이상으로 길어 AI 가 읽기 어렵습니다. Phase 10 에서 private helper 로 분해 예정:

| 함수 | 위치 | 줄수 | 목적 |
|---|---|---|---|
| `insert_body_after_heading` | `hwp_core/text_editing/insertions.py` | 354 | 제목 찾기 + 본문 삽입 + 스타일 상속 + 뎁스별 들여쓰기 |
| `set_paragraph_style` | `hwp_core/formatting/char_para.py` | 245 | ParaShape 전체 속성 + 4면 테두리 + indent/margin |
| `clone_pdf_to_hwp` | `pdf_clone/layout.py` | 358 | PDF → HWP 전체 파이프라인 (5 stages) |

Phase 10 전에는 회귀 fixture 준비 필수.

---

## 7. 인덱스 업데이트 규칙

새 MCP tool 추가 / 구조 변경 시:

1. 이 파일의 "1. 의도별 빠른 찾기" 에 행 추가 (자주 쓰는 tool 만)
2. "2. REGISTRY 메서드 색인" 에 A-Z 순서로 추가
3. "3. 모듈 안내" 의 해당 섹션에 one-liner 추가
4. 검증: `py -3.13 mcp-server/scripts/verify_python_index.py`
5. `CHANGELOG.md` 에 변경 기록
6. 큰 구조 변경 시: `ARCHITECTURE.md` 도 업데이트

---

## 8. 작업 이력 (인덱스 자체 버전)

- **2026-04-11 (Phase 0)**: 스켈레톤 생성 — 모듈 안내만 포함
- **2026-04-11 (Phase 8)**: 완성 — 의도별 찾기 40+ 행 + REGISTRY 112개 전체 + 모든 sub-package 구조 반영

---

## 참고

- **상위 문서**: `../../CLAUDE.md` (프로젝트 전체 아키텍처)
- **버전 이력**: `../../CHANGELOG.md` (v0.5.0 → v0.7.9)
- **시스템 다이어그램**: `../../ARCHITECTURE.md`
- **검증 스크립트**: `../scripts/verify_python_index.py`
