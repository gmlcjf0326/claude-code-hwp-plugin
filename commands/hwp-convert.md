---
description: HWP 문서를 PDF, DOCX, HTML, HWPX 형식으로 변환합니다.
allowed-tools:
  - "mcp__hwp-studio__hwp_open_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_open_document"
  - "mcp__hwp-studio__hwp_save_document"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_save_document"
  - "mcp__hwp-studio__hwp_export_pdf"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_export_pdf"
  - "mcp__hwp-studio__hwp_export_docx"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_export_docx"
  - "mcp__hwp-studio__hwp_export_html"
  - "mcp__plugin_hwp-studio_hwp-studio__hwp_export_html"
---

# HWP 문서 변환

## 사전 질문
1. "변환할 파일 경로를 알려주세요"
2. "어떤 형식으로 변환할까요? (PDF/DOCX/HTML/HWPX)"
3. "저장 경로를 지정하시겠어요? (기본: 같은 폴더)"

## 워크플로우
1. hwp_open_document → 원본 파일 열기
2. 형식에 따라:
   - PDF → hwp_export_pdf
   - DOCX → hwp_export_docx
   - HTML → hwp_export_html
   - HWPX → hwp_save_document(format: "hwpx")

## 참고
- DOCX 변환은 시간이 걸릴 수 있습니다 (최대 120초)
- HWPX 변환 권장: 이후 텍스트 작업이 더 안정적입니다
