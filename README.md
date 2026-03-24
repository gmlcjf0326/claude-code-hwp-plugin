# claude-code-hwp-plugin

Claude Code에서 한글(HWP/HWPX) 문서를 AI로 자동 편집하는 플러그인입니다.

85개+ MCP 도구를 번들하고, 8개 워크플로우 커맨드와 HWPX 규칙 hooks로 편의성을 극대화합니다.

> Windows 전용 | 한글 2014 이상 | Python 3.8+

## MCP vs 플러그인 — 뭘 써야 하나요?

| | MCP (claude-code-hwp-mcp) | Plugin (이 저장소) |
|---|---|---|
| **대상** | Claude Desktop + Code | Claude Code 전용 |
| **설치** | npm install -g + JSON 설정 | claude plugins install |
| **85개+ 도구** | O | O (MCP 번들) |
| **워크플로우 커맨드** | X | O (8개) |
| **HWPX 규칙 자동 검증** | X | O (hooks) |
| **에러 복구 가이드** | 에러 메시지만 | O (자동 안내) |
| **자동 트리거** | X | O (agent) |

**Claude Desktop 사용자**: [MCP 패키지](https://github.com/gmlcjf0326/claude-code-hwp-mcp) 사용

**Claude Code 사용자**: 이 플러그인 사용 (MCP가 포함되어 있으므로 별도 설치 불필요)

---

## 사전 요구사항

| 요구사항 | 설치 방법 |
|----------|-----------|
| Windows 10/11 | macOS/Linux 미지원 |
| 한글(HWP) 프로그램 | 한컴오피스 2014 이상 |
| Python 3.8+ | [python.org](https://www.python.org/downloads/) (Store 버전 X) |
| pyhwpx | `pip install pyhwpx pywin32` |

## 설치

```bash
claude plugins install github:gmlcjf0326/claude-code-hwp-plugin
```

또는 로컬 설치:

```bash
git clone https://github.com/gmlcjf0326/claude-code-hwp-plugin.git
cd claude-code-hwp-plugin
npm install
claude plugins install .
```

## 시작하기

1. 한글(HWP) 프로그램을 실행합니다
2. Claude Code에서 `/hwp-setup`으로 환경을 확인합니다
3. `/hwp-help`로 사용 가능한 기능을 확인합니다

---

## 커맨드 (8개)

### /hwp-help — 기능 안내

모든 커맨드와 사용법을 안내합니다. 처음 사용할 때 실행하세요.

### /hwp-setup — 환경 진단

Python, pyhwpx, 한글 프로그램의 설치 상태를 자동 진단합니다. 문제가 있으면 단계별 해결 방법을 안내합니다.

```
/hwp-setup
→ "Python 3.13.12 설치됨, pyhwpx 1.7.1 설치됨, 한글 실행 중"
```

### /hwp-fill — 양식 자동 채우기

사업계획서, 과업지시서, 신청서 등의 빈칸을 AI가 자동으로 채웁니다.

```
/hwp-fill
→ "파일 경로를 알려주세요" → "참고자료가 있나요?" → 자동 채우기 → 저장
```

워크플로우: 사전질문 → 문서열기 → 분석 → 채우기 → 개인정보확인 → 저장

### /hwp-write — 문서 작성

AI가 한글 문서를 처음부터 작성합니다. 양식이 있으면 서식을 따릅니다.

```
/hwp-write
→ "어떤 문서를 작성할까요?" → "양식 파일이 있나요?" → 작성 → 저장
```

### /hwp-analyze — 문서 분석

문서의 구조, 내용, 완성도를 분석하여 요약합니다.

```
/hwp-analyze
→ 문서 종류, 페이지수, 표수, 빈 항목, 내용 요약
```

### /hwp-convert — 형식 변환

HWP 문서를 PDF, DOCX, HTML, HWPX로 변환합니다.

```
/hwp-convert
→ "어떤 형식으로 변환할까요?" → 변환 → 저장
```

### /hwp-batch — 다건 문서 생성

엑셀/CSV 데이터를 기반으로 여러 건의 문서를 일괄 생성합니다.

```
/hwp-batch
→ "템플릿 파일과 데이터 파일을 알려주세요" → 행마다 문서 생성
```

### /hwp-privacy — 개인정보 스캔

문서에서 주민번호, 전화번호, 이메일 등 개인정보를 자동 감지합니다.

```
/hwp-privacy
→ 감지된 개인정보 목록 → "마스킹 처리할까요?"
```

---

## Hooks — 자동 규칙 검증

### pre-tool-use (코드 작성 시 자동 검증)

| 규칙 | 차단 대상 | 안내 |
|------|-----------|------|
| CLAUDE.md #6 | fast-xml-parser 설치 | @xmldom/xmldom 사용 |
| CLAUDE.md #7 | .tagName 사용 | .localName 사용 |
| CLAUDE.md #9 | charPrIDRef 변경 | 문서 깨짐 방지 |
| CLAUDE.md #1 | raw win32com | pyhwpx Hwp() 사용 |

### post-tool-use (에러 자동 복구 가이드)

| 에러 패턴 | 자동 안내 |
|-----------|-----------|
| RPC/COM 에러 | 한글 재시작 안내 |
| EBUSY (파일 잠금) | 한글에서 파일 닫기 안내 |
| Python 미설치 | python.org 설치 안내 |
| 파일 경로 오류 | hwp_list_files 검색 안내 |
| 문서 미열기 | hwp_open_document 안내 |
| 타임아웃 | 재시도 안내 |
| HWP find_replace 실패 | HWPX 변환 권유 |

---

## Agent — 자동 감지

"한글", "hwp", "양식", "사업계획서", "보고서" 등의 키워드가 감지되면 자동으로 환경을 체크하고 적절한 커맨드를 추천합니다.

---

## MCP 도구 (85개+)

플러그인 설치 시 자동으로 번들된 MCP 서버가 등록됩니다.

### 환경/문서 관리 (6개)

hwp_check_setup, hwp_list_files, hwp_open_document, hwp_close_document, hwp_save_document, hwp_export_pdf

### 문서 분석 (16개)

hwp_analyze_document, hwp_get_document_text, hwp_get_document_info, hwp_get_tables, hwp_map_table_cells, hwp_get_cell_format, hwp_get_table_format_summary, hwp_get_fields, hwp_get_as_markdown, hwp_get_page_text, hwp_text_search, hwp_form_detect, hwp_extract_style_profile, hwp_image_extract, hwp_document_split, hwp_read_reference

### 텍스트 편집 (18개)

hwp_insert_text, hwp_insert_markdown, hwp_insert_heading, hwp_find_replace, hwp_find_replace_multi, hwp_find_replace_nth, hwp_find_and_append, hwp_set_paragraph_style, hwp_indent, hwp_outdent, hwp_insert_page_break, hwp_insert_page_num, hwp_insert_date_code, hwp_insert_footnote, hwp_insert_endnote, hwp_insert_hyperlink, hwp_insert_auto_num, hwp_insert_memo

### 표 편집 (18개)

hwp_fill_table_cells, hwp_fill_fields, hwp_table_create_from_data, hwp_table_insert_from_csv, hwp_table_add_row, hwp_table_add_column, hwp_table_delete_row, hwp_table_delete_column, hwp_table_merge_cells, hwp_table_split_cell, hwp_table_distribute_width, hwp_table_swap_type, hwp_table_formula_sum, hwp_table_formula_avg, hwp_table_to_csv, hwp_table_to_json, hwp_set_cell_color, hwp_set_table_border

### 스마트/복합 (16개)

hwp_smart_analyze, hwp_smart_fill, hwp_auto_fill_from_reference, hwp_auto_map_reference, hwp_generate_multi_documents, hwp_generate_toc, hwp_create_gantt_chart, hwp_document_merge, hwp_document_summary, hwp_privacy_scan, hwp_batch_convert, hwp_compare_documents, hwp_word_count, hwp_delete_guide_text, hwp_toggle_checkbox, hwp_inspect_com_object

### 이미지/레이아웃 (5개)

hwp_insert_picture, hwp_set_background_picture, hwp_insert_line, hwp_break_section, hwp_break_column

### HWPX (4개, 한글 없이 동작)

hwp_template_list, hwp_document_create, hwp_template_generate, hwp_xml_edit_text

### 내보내기 (2개)

hwp_export_docx, hwp_export_html

---

## HWP vs HWPX

| 구분 | HWP | HWPX |
|------|-----|------|
| 텍스트 검색 | COM (제한적) | XML 직접 (안정적) |
| 찾기/바꾸기 | COM (제한적) | XML 직접 (안정적) |
| 표 생성/편집 | COM | COM |
| 문서 열기/저장 | COM | COM |

**HWPX 사용 권장.** 한글에서 "다른 이름으로 저장" > HWPX 형식 선택.

## 알려진 제한사항

- HWP 바이너리 파일에서 text_search가 0건을 반환할 수 있음 (COM API 한계)
- HWPX 파일이 한글에서 열린 상태에서 XML 편집 시 COM 폴백
- Windows 전용 (COM API 기반)

## 문제 해결

### 환경 문제
`/hwp-setup`을 실행하면 자동으로 진단하고 해결 방법을 안내합니다.

### 텍스트 치환이 안 됩니다
HWP → HWPX로 변환하면 더 안정적으로 작동합니다.

### 한글이 응답하지 않습니다
작업 관리자(Ctrl+Shift+Esc)에서 Hwp.exe를 종료하고 다시 실행하세요.

---

## 라이선스

MIT License

## 관련 링크

- [MCP 패키지 (Claude Desktop용)](https://github.com/gmlcjf0326/claude-code-hwp-mcp)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [MCP Protocol](https://modelcontextprotocol.io/)
