---
description: HWP/HWPX 양식의 빈 항목을 AI가 자동으로 채웁니다. 사업계획서, 과업지시서, 신청서 등.
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
2. "참고할 자료(엑셀, 텍스트, 기존 문서 등)가 있나요?"
3. "분량은 간결하게 / 표준 / 상세하게 중 어느 수준으로?"
4. "확정된 수치(금액, 일정 등)가 있으면 알려주세요"

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
