---
description: 엑셀/CSV 데이터를 기반으로 여러 건의 HWP 문서를 일괄 생성합니다. 위촉장, 증명서 등.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_close_document"
  - "mcp__hwp-studio__hwp_generate_multi_documents"
  - "mcp__hwp-studio__hwp_read_reference"
  - "mcp__hwp-studio__hwp_list_files"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__hwp-studio__hwp_batch_convert"
---

# HWP 다건 문서 일괄 생성

## 사전 질문
1. "템플릿 HWP 파일 경로를 알려주세요"
2. "데이터 파일(엑셀/CSV) 경로를 알려주세요"
3. "생성된 파일을 저장할 폴더를 지정해주세요"

## 워크플로우
1. hwp_read_reference → 엑셀/CSV 데이터 읽기
2. hwp_generate_multi_documents → 행마다 별도 문서 생성
3. 결과 보고: 생성된 파일 수, 경로

## 참고
- 각 행의 데이터가 템플릿의 빈칸에 매핑됩니다
- 대량 생성 시 시간이 걸릴 수 있습니다
