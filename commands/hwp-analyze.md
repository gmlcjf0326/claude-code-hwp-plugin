---
description: HWP 문서의 구조와 내용을 분석하여 요약합니다. 문서 종류, 완성도, 빈 항목을 파악합니다.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_smart_analyze"
  - "mcp__hwp-studio__hwp_get_document_text"
  - "mcp__hwp-studio__hwp_get_tables"
  - "mcp__hwp-studio__hwp_word_count"
  - "mcp__hwp-studio__hwp_form_detect"
  - "mcp__hwp-studio__hwp_text_search"
  - "mcp__hwp-studio__hwp_extract_style_profile"
  - "mcp__hwp-studio__hwp_get_fields"
---

# HWP 문서 분석

## 워크플로우
1. hwp_open_document → 파일 열기
2. hwp_smart_analyze → 심층 분석 (문서 유형 추론, 서식 프로파일)
3. hwp_get_tables → 표 구조/데이터 파악
4. hwp_form_detect → 양식 빈칸/체크박스 감지
5. hwp_word_count → 글자수/단어수/페이지수

## 보고 항목
- 문서 종류 및 목적
- 페이지수, 표수, 필드수
- 작성 완성도(%)
- 빈 항목이 있다면 목록
- 내용 요약 (3-5줄)
- 사용된 서식 (글꼴, 크기, 줄간격)

## HWP text_search 참고
HWP 바이너리 파일에서 text_search가 0건을 반환할 수 있습니다.
이 경우 hwp_get_document_text로 전체 텍스트를 가져와서 직접 검색하세요.
HWPX 파일에서는 XML 검색이 안정적으로 작동합니다.
