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
- hwp_smart_analyze로 구조 파악 (표, 필드, 빈칸)
- hwp_extract_style_profile로 서식 파악 (글꼴, 크기, 자간)

### Step 3: 미리보기
- "다음과 같이 채울 예정입니다:" → 내용 요약 제시
- 사용자 확인 후 진행

### Step 4: 채우기
- 참고자료 있으면: hwp_read_reference로 로드 → hwp_auto_fill_from_reference
- 없으면: hwp_smart_fill (서식 보존)
- 필요 시 hwp_fill_table_cells로 개별 셀 조정

### Step 5: 검증 + 저장
- hwp_get_tables로 채운 결과 확인
- hwp_privacy_scan으로 개인정보 확인
- hwp_save_document로 저장
- "수정할 부분이 있으면 말씀해주세요"

## 규칙
- 결재란/서명란: AI가 채우면 안 됨
- 이미 내용이 있는 셀: 변경하지 않음
- 사용자가 제공하지 않은 수치는 임의로 만들지 않음
- 날짜 형식: 마침표 구분 (예: 2026. 3. 19.)
- 금액: 한글+숫자 병기 (예: 금일백만원정(1,000,000원))

## 막혔을 때
- 표가 비어있음 → hwp_map_table_cells로 셀 매핑 확인, label 기반 매칭 시도
- HWPX 권장 → 한글에서 "다른 이름으로 저장" > HWPX 선택
- 한글 응답 없음 → 작업 관리자에서 Hwp.exe 종료 후 재시작
