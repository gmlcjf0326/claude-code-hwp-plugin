---
description: HWP 양식 자동 채우기. Excel/PDF/DOCX 참고자료 지원. 서식 보존하며 빈칸을 AI가 채웁니다.
allowed-tools:
  - "mcp__hwp-studio__hwp_check_setup"
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_close_document"
  - "mcp__hwp-studio__hwp_smart_analyze"
  - "mcp__hwp-studio__hwp_analyze_document"
  - "mcp__hwp-studio__hwp_smart_fill"
  - "mcp__hwp-studio__hwp_auto_fill_from_reference"
  - "mcp__hwp-studio__hwp_read_reference"
  - "mcp__hwp-studio__hwp_fill_table_cells"
  - "mcp__hwp-studio__hwp_fill_fields"
  - "mcp__hwp-studio__hwp_privacy_scan"
  - "mcp__hwp-studio__hwp_get_tables"
  - "mcp__hwp-studio__hwp_get_document_text"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__hwp-studio__hwp_extract_style_profile"
  - "mcp__hwp-studio__hwp_map_table_cells"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_check_setup"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_open_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_close_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_smart_analyze"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_analyze_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_smart_fill"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_auto_fill_from_reference"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_read_reference"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_fill_table_cells"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_fill_fields"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_privacy_scan"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_get_tables"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_get_document_text"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_save_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_extract_style_profile"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_map_table_cells"
  - "mcp__hwp-studio__hwp_verify_layout"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_verify_layout"
---

# HWP 양식 자동 채우기

## 반드시 사전 질문부터 (바로 채우지 마세요)

사용자에게 먼저 확인하세요:
1. "채울 HWP/HWPX 파일 경로를 알려주세요"
2. "참고할 자료가 있나요?" (Excel/CSV/PDF/DOCX/PPT/HWP 등 모든 형식 지원)
   → 있으면: "기존 작업에 데이터를 추가할까요, 새로운 작업으로 진행할까요?"
   → 있으면: "참조 데이터를 docs/references/에 .md로 저장해둘까요? (이후 재사용 가능)"

   💡 **참고자료 권장 볼륨 (v0.7.4.8)**:
   - 🟢 **최적** 1-3개 파일, 총 60KB 이하 (가장 정확, LLM focus 최상)
   - 🟡 **적정** 최대 5개, 총 150KB 이하 (실무 기본)
   - 🟠 **주의** 5개 초과 또는 총 150KB 초과 → focus degradation 시작
   - 🔴 **최대** 10개 또는 500KB — 품질 저하 감수 (split 권장)
3. "분량은 간결하게 / 표준 / 상세하게 중 어느 수준으로?"
4. "확정된 수치(금액, 일정 등)가 있으면 알려주세요"
5. "완성 후 PDF로도 변환할까요?"
6. "**작업 우선순위를 선택해 주세요** (v0.7.4.8 신규):
   ⚡ **빠른 완성** — 단순 양식 채우기, 2-3쪽 표준 문서 (Haiku/Sonnet, 3-5분)
   ⚖️ **균형** (기본) — 대부분의 업무 문서 (Sonnet, 5-8분)
   ⭐ **최고 품질** — 복잡한 구조, 법적/공식 문서 (Opus, 8-15분)

   답변 없으면: hwp_estimate_workload 결과로 자동 판단 (simple_fill→Haiku, text_generation→Sonnet, structured_analysis→Opus)"

## 📊 Pre-flight 분석 (질문 후, 채우기 전)

모든 답변 수집 후 **반드시** `hwp_estimate_workload` 실행해 아래 정보를 사용자에게 제시:

```
📊 예상 작업 분석:
- 분량: {estimated_pages}쪽 ({estimated_sections}섹션, {estimated_tables}표)
- 복잡도: {suggested_workflow}  (simple_fill | text_generation | structured_analysis)
- 예상 시간:
  ⚡ Haiku  {duration_by_model.haiku}초
  ⚖️ Sonnet {duration_by_model.sonnet}초
  ⭐ Opus   {duration_by_model.opus}초
- 토큰: {tokens.total_tokens_estimate} ({tokens.context_window_usage_percent}% of 200K)
- 참고자료: {context_efficiency.recommendation}
- 권장: {recommended_model_for_complexity.default} — {recommended_model_for_complexity.reason}

이대로 진행할까요? 다른 모델 원하시면 ⚡/⚖️/⭐ 중 선택.
```

확인 후 사용자가 선택한 모델로 delegation:
- ⚡ → `/hwp-fast` 서브 에이전트 (haiku 또는 sonnet)
- ⚖️ → `/hwp-standard` 서브 에이전트 (sonnet, 기본)
- ⭐ → `/hwp-quality` 서브 에이전트 (opus)

## 워크플로우

### Step 1: 환경 확인
hwp_check_setup으로 Python/한글 상태 확인. 문제 있으면 `/hwp-setup` 안내.

### Step 2: 문서 열기 + 분석
- hwp_open_document로 파일 열기
- hwp_smart_analyze로 구조 파악 (표, 필드, 빈칸). 90초 타임아웃이므로 대용량 문서는 시간이 걸릴 수 있음.
- hwp_extract_style_profile로 서식 파악 (글꼴, 크기, 자간)
- hwp_map_table_cells로 표 셀 탭 인덱스 매핑 (병합 셀 대응). 이 단계를 반드시 거쳐야 정확한 셀 위치를 알 수 있음.

### Step 3: 미리보기
- "다음과 같이 채울 예정입니다:" → 내용 요약 제시
- 사용자 확인 후 진행

### Step 4: 채우기
- 참고자료 있으면: hwp_read_reference로 로드 → hwp_auto_fill_from_reference
- 없으면: hwp_smart_fill (서식 보존)
- 필요 시 hwp_fill_table_cells로 개별 셀 조정. label 기반 매칭 권장.

### Step 5: 검증 + 저장
- hwp_get_tables로 채운 결과 확인
- hwp_privacy_scan으로 개인정보 확인
- hwp_save_document로 저장
- "수정할 부분이 있으면 말씀해주세요"

## 공문서 양식 채우기 특화

### 결재란 보존
- 결재란(담당/검토/협조/결재) 셀은 **절대 채우지 않음**
- hwp_map_table_cells로 결재란 위치를 먼저 파악하고 해당 탭 인덱스 제외

### 표 구조 파악
- hwp_map_table_cells로 병합 셀 포함 탭 인덱스 확인 (필수)
- label 기반 매칭 권장: `{label: "사업명", text: "AI 자동화"}`
- tab 기반은 병합 셀에서 인덱스 예측 어려우므로 label 우선

### 공문서 순번 체계
채울 때 순번 형식 유지:
- 1. → 가. → 1) → 가) → (1) → (가) → ① → ㉮

## 규칙
- 결재란/서명란: AI가 채우면 안 됨
- 이미 내용이 있는 셀: 변경하지 않음
- 사용자가 제공하지 않은 수치는 임의로 만들지 않음
- 날짜 형식: 마침표 구분 (예: 2026. 3. 19.)
- 금액: 한글+숫자 병기 (예: 금일백만원정(1,000,000원))
- HWPX 파일 권장 (텍스트 작업 더 안정적)

## 막혔을 때
- **표가 비어있음** → hwp_map_table_cells로 셀 매핑 확인, label 기반 매칭 시도
- **HWPX 권장** → 한글에서 "다른 이름으로 저장" > HWPX 선택
- **한글 응답 없음** → 작업 관리자(Ctrl+Shift+Esc)에서 Hwp.exe 종료 후 재시작
- **타임아웃** → 대용량 문서는 smart_analyze 대신 analyze_document 사용 (더 빠름)
- **셀 위치를 모르겠음** → hwp_map_table_cells로 탭 인덱스 확인 후 tab 기반으로 채우기
