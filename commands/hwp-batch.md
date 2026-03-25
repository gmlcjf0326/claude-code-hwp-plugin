---
description: Excel 데이터로 위촉장/증명서 등 다건 일괄 생성. 파일명 규칙 지정, PDF 일괄 변환 포함.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_close_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_close_document"
  - "mcp__hwp-studio__hwp_generate_multi_documents"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_generate_multi_documents"
  - "mcp__hwp-studio__hwp_read_reference"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_read_reference"
  - "mcp__hwp-studio__hwp_list_files"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_list_files"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_save_document"
  - "mcp__hwp-studio__hwp_batch_convert"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_batch_convert"
---

# HWP 다건 문서 일괄 생성

## 사전 질문
1. "양식(템플릿) HWP 파일 경로를 알려주세요"
2. "데이터 파일(엑셀/CSV) 경로를 알려주세요" (DOCX/PDF도 가능 — 자동 변환)
3. "생성된 파일을 저장할 폴더를 지정해주세요"
4. "총 몇 건 생성 예정인지 확인합니다" (데이터 행 수 자동 감지)
5. "파일명 규칙이 있나요?" (예: 위촉장_홍길동.hwp, 증명서_001.hwp)
6. "완성 후 PDF로도 변환할까요?"

## 워크플로우
1. hwp_read_reference → 데이터 읽기 (Excel/CSV/DOCX/PDF 자동 지원)
2. 데이터 미리보기 → "총 N건 생성 예정입니다. 진행할까요?"
3. hwp_generate_multi_documents → 행마다 별도 문서 생성
4. PDF 변환 요청 시 → hwp_batch_convert
5. 결과 보고: 생성된 파일 수, 경로, 소요 시간

## 참고
- 각 행의 데이터가 템플릿의 빈칸에 자동 매핑됩니다
- 대량 생성(100건+) 시 수 분 소요될 수 있습니다
- 중간에 에러 발생 시 이미 생성된 파일은 유지됩니다
