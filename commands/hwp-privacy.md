---
description: 개인정보 자동 스캔. 주민번호/전화번호/이메일/계좌번호 등 감지. 제출 전 필수 확인.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_privacy_scan"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_privacy_scan"
  - "mcp__hwp-studio__hwp_find_replace"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_find_replace"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_save_document"
---

# HWP 개인정보 스캔

## 워크플로우
1. hwp_open_document → 파일 열기
2. hwp_privacy_scan → 개인정보 자동 감지
3. 발견된 항목 보고 (유형, 위치, 마스킹된 값)
4. 사용자에게 "마스킹 처리할까요?" 확인
5. 필요 시 hwp_find_replace로 마스킹 처리
6. hwp_save_document → 저장

## 감지 항목
- 주민등록번호 (000000-0000000)
- 전화번호 (010-0000-0000)
- 이메일 주소
- 계좌번호
- 여권번호
