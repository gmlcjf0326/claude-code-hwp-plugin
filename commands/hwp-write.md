---
description: AI가 한글 문서를 처음부터 작성합니다. 양식이 있으면 서식을 따르고, 없으면 표준 형식으로 작성합니다.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_close_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_close_document"
  - "mcp__hwp-studio__hwp_insert_text"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_insert_text"
  - "mcp__hwp-studio__hwp_insert_heading"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_insert_heading"
  - "mcp__hwp-studio__hwp_insert_markdown"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_insert_markdown"
  - "mcp__hwp-studio__hwp_set_paragraph_style"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_set_paragraph_style"
  - "mcp__hwp-studio__hwp_extract_style_profile"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_extract_style_profile"
  - "mcp__hwp-studio__hwp_delete_guide_text"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_delete_guide_text"
  - "mcp__hwp-studio__hwp_table_create_from_data"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_table_create_from_data"
  - "mcp__hwp-studio__hwp_create_gantt_chart"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_create_gantt_chart"
  - "mcp__hwp-studio__hwp_generate_toc"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_generate_toc"
  - "mcp__hwp-studio__hwp_insert_page_break"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_insert_page_break"
  - "mcp__hwp-studio__hwp_insert_page_num"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_insert_page_num"
  - "mcp__hwp-studio__hwp_indent"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_indent"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_save_document"
  - "mcp__hwp-studio__hwp_read_reference"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_read_reference"
---

# HWP 문서 작성

## 사전 질문
1. "어떤 문서를 작성할까요? (사업계획서/보고서/공문/기안/제안서)"
2. "양식 파일이 있나요? (서식을 따를 기존 문서)"
3. "참고자료가 있나요?"
4. "문체: 개괄식(~했음) / 격식체(~했습니다)?"
5. "예상 분량? (예: 5~6페이지)"

## 워크플로우
1. 양식 있으면 → hwp_open_document + hwp_extract_style_profile로 서식 파악
2. hwp_delete_guide_text → 작성요령/가이드 텍스트 삭제
3. hwp_insert_heading → 제목 (서식 프로파일의 글꼴/크기 적용)
4. hwp_insert_text / hwp_insert_markdown → 본문 내용
5. hwp_table_create_from_data → 필요 시 표 생성
6. hwp_create_gantt_chart → 추진일정표 (사업계획서)
7. hwp_generate_toc → 목차 (필요 시)
8. hwp_save_document → 저장

## 공문서(기안문/시행문) 작성 시

### 표 기반 레이아웃
공문서는 전체 페이지를 하나의 표로 구성합니다:
```
hwp_table_create_from_data(
  data=[기관명, 시행일/결재란, 수신, 제목, 본문, ...],
  col_widths=[18, 65, 23, 23, 23, 23],  // mm 단위
  row_heights=[15, 10, 10, 10, 12, 12, 40]
)
```

### 셀 병합 (하단→상단 순서)
```
hwp_table_merge_cells(start_row=6, end_row=6, start_col=0, end_col=5)  // 본문 전체
hwp_table_merge_cells(start_row=4, end_row=4, start_col=1, end_col=5)  // 수신 우측
hwp_table_merge_cells(start_row=0, end_row=0, start_col=0, end_col=5)  // 기관명 전체
hwp_table_merge_cells(start_row=1, end_row=3, start_col=0, end_col=1)  // 시행일 좌측
```

### 공문서 순번 체계 (8단계)
1. → 가. → 1) → 가) → (1) → (가) → ① → ㉮
hwp_insert_heading의 numbering 옵션: roman, decimal, korean, circle, paren_decimal, paren_korean

### 공문서 3단 구조
- **두문**: 기관명 + 수신자 + 문서번호
- **본문**: 제목 + 내용 + 붙임
- **결문**: 발신명의 + 기안자/검토자/결재자 + 연락처

## 규칙
- 양식이 제공되면 반드시 서식(글꼴/크기/자간/들여쓰기) 동일하게 적용
- 결재란/서명란은 비워둠 (AI가 채우면 안 됨)
- 날짜: 마침표 구분 (2026. 3. 19.)
- 금액: 한글+숫자 병기 (금일백만원정(1,000,000원))
- 문체: 개괄식(~했음) 또는 격식체(~했습니다) 통일
- HWPX 파일 권장 (텍스트 작업 더 안정적)
- 한글 프로그램이 실행 중이어야 함
