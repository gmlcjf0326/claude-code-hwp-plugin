---
description: AI가 한글 문서를 처음부터 작성합니다. 양식이 있으면 서식을 따르고, 없으면 표준 형식으로 작성합니다.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_close_document"
  - "mcp__hwp-studio__hwp_insert_text"
  - "mcp__hwp-studio__hwp_insert_heading"
  - "mcp__hwp-studio__hwp_insert_markdown"
  - "mcp__hwp-studio__hwp_set_paragraph_style"
  - "mcp__hwp-studio__hwp_extract_style_profile"
  - "mcp__hwp-studio__hwp_delete_guide_text"
  - "mcp__hwp-studio__hwp_table_create_from_data"
  - "mcp__hwp-studio__hwp_create_gantt_chart"
  - "mcp__hwp-studio__hwp_generate_toc"
  - "mcp__hwp-studio__hwp_insert_page_break"
  - "mcp__hwp-studio__hwp_insert_page_num"
  - "mcp__hwp-studio__hwp_indent"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__hwp-studio__hwp_read_reference"
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

## 규칙
- 양식이 제공되면 반드시 서식(글꼴/크기/자간/들여쓰기) 동일하게 적용
- 결재란/서명란은 비워둠
- 공문서: 날짜 마침표 구분 (2026. 3. 19.)
- 한글 프로그램이 실행 중이어야 함
