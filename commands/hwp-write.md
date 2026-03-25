---
description: 문서 처음부터 작성. 공문/사업계획서/보고서/동의서 등. 긴 문서는 섹션별 분할 작성. PDF 변환 포함.
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
  - "mcp__hwp-studio__hwp_verify_layout"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_verify_layout"
  - "mcp__hwp-studio__hwp_table_merge_cells"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_table_merge_cells"
---

# HWP 문서 작성

## 사전 질문 (반드시 순서대로 진행)
1. "어떤 문서를 작성할까요?" (사업계획서/보고서/공문/기안/제안서/동의서/위촉장/기타)
2. "양식 파일이 있나요?" (서식을 따를 기존 문서)
3. "참고자료가 있나요?" (Excel/PDF/DOCX/PPT 등 — 모든 형식 지원)
   → 있으면: "참조 데이터를 docs/references/에 .md로 저장해둘까요? (이후 재사용 가능)"
4. "문체: 개괄식(~했음) / 격식체(~했습니다)?"
5. "예상 분량?" (1~2쪽 / 5~10쪽 / 10쪽 이상)
   → 5쪽 이상: "섹션별로 나눠서 작성하겠습니다. 먼저 목차를 잡아볼까요?"
6. "완성 후 PDF로도 변환할까요?"

## 참고자료 처리
- 참고자료가 있으면 hwp_read_reference로 읽기 (DOCX/PPTX도 자동 변환)
- 추출된 데이터를 docs/references/파일명.md로 저장 (Claude Code가 이후 참조)
- 긴 참고자료는 자동 분리: 참고자료_1.md, 참고자료_2.md

## 워크플로우

### 짧은 문서 (1~4쪽: 공문/동의서/위촉장)
1. 양식 있으면 → hwp_open_document + hwp_extract_style_profile
2. hwp_delete_guide_text → 작성요령 삭제
3. hwp_set_page_setup → 여백/용지 설정
4. hwp_insert_text/heading → 본문 작성 (한번에 완성)
5. hwp_table_create_from_data → 표 생성 + set_cell_color + fill_table_cells
6. hwp_verify_layout → PDF 시각 확인
7. 문제 시 조정 → 재확인
8. hwp_save_document → 저장 (+ PDF 변환 요청 시 hwp_export_pdf)

### 긴 문서 (5쪽 이상: 사업계획서/제안서)
1. 목차 구성 → 사용자 확인
2. hwp_set_page_setup → 여백/용지/머리글/바닥글
3. **섹션 1 작성** (1~3쪽) → hwp_verify_layout 확인
4. **섹션 2 작성** → hwp_verify_layout
5. ... 반복 (컨텍스트 관리를 위해 섹션별 진행)
6. hwp_generate_toc → 목차 자동 생성
7. 최종 hwp_verify_layout → 전체 레이아웃 검증
8. hwp_save_document → 저장

## 공문서(기안문/시행문) 작성 시

### 표 기반 레이아웃
공문서는 전체 페이지를 하나의 표로 구성합니다.
**중요: col_widths 합계는 170mm 이내** (A4 가용 너비 = 210mm - 좌우 여백 40mm)
```
hwp_table_create_from_data(
  data=[기관명, 시행일/결재란, 수신, 제목, 본문, ...],
  col_widths=[17, 63, 22, 22, 22, 22],  // 합계 168mm (170mm 이내)
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
