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
3. "참고자료가 있나요?" (Excel/PDF/DOCX/PPT/HWP 등 — 모든 형식 지원)
   → 있으면: "참조 데이터를 docs/references/에 .md로 저장해둘까요? (이후 재사용 가능)"

   💡 **참고자료 권장 볼륨 (v0.7.4.8)**:
   - 🟢 **최적** 1-3개 파일, 총 60KB 이하 (가장 정확, LLM focus 최상)
   - 🟡 **적정** 최대 5개, 총 150KB 이하 (실무 기본)
   - 🟠 **주의** 5개 초과 또는 총 150KB 초과 → focus degradation 시작
   - 🔴 **최대** 10개 또는 500KB — 품질 저하 감수 (split 권장)
4. "문체: 개괄식(~했음) / 격식체(~했습니다)?"
5. "예상 분량?" (1~2쪽 / 5~10쪽 / 10쪽 이상)
   → 5쪽 이상: "섹션별로 나눠서 작성하겠습니다. 먼저 목차를 잡아볼까요?"
6. "완성 후 PDF로도 변환할까요?"
7. "**작업 우선순위를 선택해 주세요** (v0.7.4.8 신규):
   ⚡ **빠른 완성** — 단순 양식 채우기, 2-3쪽 표준 문서 (Haiku/Sonnet, 3-5분)
   ⚖️ **균형** (기본) — 대부분의 업무 문서 (Sonnet, 5-8분)
   ⭐ **최고 품질** — 복잡한 구조, 법적/공식 문서 (Opus, 8-15분)

   답변 없으면: hwp_estimate_workload 결과로 자동 판단 (simple_fill→Haiku, text_generation→Sonnet, structured_analysis→Opus)"

## 📊 Pre-flight 분석 (질문 후, 작성 전)

모든 답변 수집 후 **반드시** `hwp_estimate_workload` 실행해 아래 정보를 사용자에게 제시:

```
📊 예상 작업 분석:
- 분량: {estimated_pages}쪽 ({estimated_sections}섹션, {estimated_tables}표)
- 복잡도: {suggested_workflow}
- 예상 시간: ⚡{haiku}s / ⚖️{sonnet}s / ⭐{opus}s
- 토큰: {total_tokens} ({context_window_usage_percent}% of 200K)
- 참고자료: {context_efficiency.recommendation}
- 권장: {recommended_model_for_complexity.default} — {recommended_model_for_complexity.reason}

이대로 진행할까요? 다른 모델 원하시면 ⚡/⚖️/⭐ 중 선택.
```

선택 후 해당 sub-agent 로 delegation:
- ⚡ → `hwp-fast` (model: haiku 또는 sonnet)
- ⚖️ → `hwp-standard` (model: sonnet)
- ⭐ → `hwp-quality` (model: opus)

## 양식 정밀 분석 (양식 파일 있을 때 — 가장 중요)

**양식의 형태를 완벽히 재현하는 것이 핵심.** 단순 텍스트가 아닌 용지/여백/폰트/줄간격/표 치수까지 모두 분석.

### Step A: 양식 종합 프로파일 추출
```
hwp_open_document → hwp extract_full_profile
반환 정보:
├─ 용지: 크기(mm), 방향, 여백(위/아래/좌/우/머리말/꼬리말), 사용가능 영역
├─ 본문 글자: 폰트명, 크기(pt), 굵게, 기울임, 색상, 자간, 장평
├─ 본문 문단: 정렬, 줄간격(%), 들여쓰기/내어쓰기, 문단 위/아래 간격
├─ 표: 전체 너비(mm), 셀 여백, 행/열 수, 헤더 서식
└─ 작성요령 텍스트 위치
```

### Step B: 분석 결과 사용자 확인
"양식 분석 결과입니다:" 형태로 보여주고 확인 받기.
예시:
```
- 용지: A4 세로, 여백 위20/아래15/좌30/우30mm
- 본문: 바탕 11pt, 줄간격 180%, 양쪽정렬
- 제목: 맑은 고딕 16pt Bold 가운데
- 표 1: 너비 148mm, 4열×6행, 셀여백 좌우1.8mm
- 들여쓰기: 10pt, 문단아래 3pt
- 작성요령 5건 발견 (삭제 예정)
```

### Step C: 프로파일 기반 작성
**모든 insert_text/heading/table에 프로파일 서식을 그대로 적용.**
표 생성 시 원본의 col_widths/row_heights/셀여백을 정확히 재현.

## 참고자료 처리
- 참고자료가 있으면 hwp_read_reference로 읽기 (DOCX/PPTX도 자동 변환)
- 추출된 데이터를 docs/references/파일명.md로 저장 (Claude Code가 이후 참조)
- 긴 참고자료는 자동 분리: 참고자료_1.md, 참고자료_2.md

## 멀티모델 활용 가이드 (속도 최적화)
- **문서 구조 분석**: opus 모델 (복잡한 구조 이해)
- **텍스트 작성/생성**: sonnet 모델 (균형잡힌 품질+속도)
- **서식 검증/레이아웃 확인**: haiku 모델 (빠른 확인)
- **긴 문서(10쪽+)**: 섹션별 병렬 에이전트 활용

## 워크플로우

### 짧은 문서 (1~4쪽: 공문/동의서/위촉장)
1. 양식 있으면 → hwp_open_document + **extract_full_profile** (양식 정밀 분석)
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
