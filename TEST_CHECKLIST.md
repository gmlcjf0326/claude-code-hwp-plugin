# HWP Studio 플러그인 전체 테스트 체크리스트

> 플러그인 설치 = MCP 서버 94개 도구 자동 포함
> 새 터미널에서 설치부터 전 기능 검증까지 진행

---

## Phase 1: 플러그인 설치 + 환경 확인

### 1-1. 플러그인 설치

- [ ] `claude plugins add /path/to/claude-code-hwp-plugin` 실행
- [ ] `claude plugins list`에서 `claude-code-hwp-plugin` enabled 확인
- [ ] `/mcp` 입력 → `hwp-studio` MCP 서버가 목록에 있는지 확인
- [ ] MCP 도구 94개가 Available로 표시되는지 확인

### 1-2. 환경 체크

- [ ] `/hwp-setup` 실행
- [ ] Python 감지 정상 (경로 표시)
- [ ] pyhwpx 감지 정상 (버전 표시)
- [ ] 한글 프로그램 설치 확인
- [ ] 한글 프로그램 실행 중 확인

### 1-3. 커맨드 자동완성

- [ ] `/hwp` 입력 → 8개 커맨드 자동완성 목록 표시
- [ ] `/hwp-fill` → description 힌트: "HWP 양식 자동 채우기..."
- [ ] `/hwp-write` → description 힌트: "문서 처음부터 작성..."
- [ ] `/hwp-help` → description 힌트: "모든 기능과 사용법..."

---

## Phase 2: 커맨드 사전 질문 동작

### 2-1. /hwp-help

- [ ] 전체 기능 안내 페이지 표시
- [ ] 94개 도구 카테고리별 목록 표시
- [ ] 참고자료 지원 형식 표 표시
- [ ] 각 커맨드별 사용 예시 표시

### 2-2. /hwp-fill (양식 채우기)

- [ ] "채울 HWP/HWPX 파일 경로를 알려주세요" 질문
- [ ] "참고할 자료가 있나요?" 질문
- [ ] "기존 작업에 추가? 새 작업?" 질문
- [ ] "참조 데이터를 .md로 저장?" 질문
- [ ] "분량 수준?" 질문
- [ ] "완성 후 PDF 변환?" 질문

### 2-3. /hwp-write (문서 작성)

- [ ] "어떤 문서를 작성할까요?" 질문 (사업계획서/보고서/공문 등)
- [ ] "양식 파일이 있나요?" 질문
- [ ] "참고자료가 있나요?" 질문
- [ ] "예상 분량?" 질문
- [ ] 5쪽 이상 시 → "섹션별 분할 작성" 안내
- [ ] "문체?" 질문 (개괄식/격식체)
- [ ] "PDF 변환?" 질문

### 2-4. /hwp-batch (일괄 생성)

- [ ] "양식 파일 경로?" 질문
- [ ] "데이터 파일 경로?" 질문
- [ ] "저장 폴더?" 질문
- [ ] "총 건수 확인" 메시지
- [ ] "파일명 규칙?" 질문
- [ ] "PDF 변환?" 질문

### 2-5. /hwp-analyze (분석)

- [ ] 파일 경로 질문
- [ ] 분석 결과 표시 (표 구조, 텍스트 요약 등)

### 2-6. /hwp-convert (변환)

- [ ] 입력 파일 질문
- [ ] 출력 형식 질문 (PDF/DOCX/HTML/HWPX)

### 2-7. /hwp-privacy (개인정보)

- [ ] 파일 경로 질문
- [ ] 스캔 결과 표시

---

## Phase 3: 문서 관리 도구 (6개)

### 3-1. 문서 생성/열기/닫기

- [ ] "새 hwpx 문서를 만들어줘" → hwp_document_create 정상
- [ ] "C:/문서/test.hwpx 열어줘" → hwp_open_document 정상
- [ ] "현재 문서 정보 알려줘" → hwp_get_document_info (페이지수, 경로)
- [ ] "문서 닫아줘" → hwp_close_document 정상

### 3-2. 저장/내보내기

- [ ] "현재 문서 저장해줘" → hwp_save_document (HWPX)
- [ ] "HWP 형식으로 저장해줘" → hwp_save_document (HWP)
- [ ] "PDF로 내보내줘" → hwp_export_pdf 정상
- [ ] "DOCX로 내보내줘" → hwp_export_docx 정상
- [ ] "HTML로 내보내줘" → hwp_export_html 정상

### 3-3. 파일 목록

- [ ] "C:/문서 폴더에 hwp 파일 뭐 있어?" → hwp_list_files 정상

---

## Phase 4: 텍스트 편집 도구 (18개)

### 4-1. 텍스트 삽입

- [ ] "텍스트 삽입해줘: 안녕하세요" → hwp_insert_text 정상
- [ ] 연속 insert_text 호출 시 각각 별도 줄로 분리되는지 확인 (자동 줄바꿈)
- [ ] "굵은 빨간색 14pt로 '제목' 삽입해줘" → style 적용 확인
  - [ ] bold 적용
  - [ ] color 적용 (빨간색)
  - [ ] font_size 적용 (14pt)
- [ ] "밑줄 + 기울임으로 삽입해줘" → underline + italic
- [ ] "맑은 고딕 폰트로 삽입해줘" → font_name 적용
- [ ] 마크다운 삽입: "마크다운으로 삽입해줘" → hwp_insert_markdown

### 4-2. 제목 삽입

- [ ] "1단계 제목: 서론" → hwp_insert_heading level=1
- [ ] "순번 포함 제목: 1. 개요" → numbering=decimal
- [ ] "가. 세부내용" → numbering=korean
- [ ] "Ⅰ. 대제목" → numbering=roman

### 4-3. 검색/치환

- [ ] "OOO을 XXX로 바꿔줘" → hwp_find_replace
- [ ] "모든 '갑'을 '을'로 바꿔줘" → hwp_find_replace (전체)
- [ ] "3번째 나오는 '항목'만 바꿔줘" → hwp_find_replace_nth
- [ ] "'계약금액' 다음에 '(VAT 포함)' 추가해줘" → hwp_find_and_append
- [ ] "여러 항목 동시 치환" → hwp_find_replace_multi
- [ ] "'사업명' 텍스트 검색해줘" → hwp_text_search

### 4-4. 문단 서식

- [ ] "가운데 정렬로 바꿔줘" → hwp_set_paragraph_style align=center
- [ ] "줄간격 180%로" → line_spacing=180
- [ ] "들여쓰기 20pt" → indent=20
- [ ] "문단 앞 간격 10pt" → space_before=10

### 4-5. 기타 텍스트

- [ ] "페이지 나눠줘" → hwp_insert_page_break
- [ ] "쪽 번호 넣어줘" → hwp_insert_page_num
- [ ] "오늘 날짜 삽입해줘" → hwp_insert_date_code
- [ ] "각주 넣어줘: 출처" → hwp_insert_footnote
- [ ] "미주 넣어줘" → hwp_insert_endnote
- [ ] "링크 삽입해줘" → hwp_insert_hyperlink
- [ ] "메모 남겨줘" → hwp_insert_memo
- [ ] "들여쓰기 한 단계" → hwp_indent
- [ ] "들여쓰기 취소" → hwp_outdent

---

## Phase 5: 표 편집 도구 (18개) — 가장 중요

### 5-1. 표 생성

- [ ] "3행 4열 표 만들어줘" → hwp_table_create_from_data
  - [ ] 데이터가 정확히 채워지는지
  - [ ] col_widths 적용 확인 (mm 단위)
  - [ ] row_heights 적용 확인
  - [ ] alignment=center 적용 확인
  - [ ] header_style=true → 헤더행 Bold
- [ ] "CSV 파일로 표 생성해줘" → hwp_table_insert_from_csv

### 5-2. 셀 배경색 (핵심)

- [ ] "헤더행을 진남색으로" → hwp_set_cell_color #003366
  - [ ] 실제 진남색이 PDF에 보이는지 verify_layout으로 확인
- [ ] "합계행을 연회색으로" → hwp_set_cell_color #F0F0F0
- [ ] 여러 셀 동시 색상 변경

### 5-3. 셀 텍스트 채우기 + 정렬 (핵심)

- [ ] "헤더 흰색 굵게 가운데" → hwp_fill_table_cells
  - [ ] color: [255,255,255] 적용 확인
  - [ ] bold: true 적용 확인
  - [ ] align: "center" → 가운데 정렬 확인
- [ ] label 기반 매칭: "계약금액 칸에 50,000,000원" → label:"계약금액"
- [ ] tab 기반 매칭: tab 인덱스로 특정 셀 채우기
- [ ] row/col 기반 매칭

### 5-4. 셀 병합 (핵심)

- [ ] "마지막 행의 첫 2열 병합해줘" → hwp_table_merge_cells
  - [ ] hwp_map_table_cells로 병합 후 셀 수 감소 확인
- [ ] 여러 행 병합 (세로 병합)
- [ ] 여러 열 병합 (가로 병합)

### 5-5. 표 구조 변경

- [ ] "행 추가해줘" → hwp_table_add_row
- [ ] "열 추가해줘" → hwp_table_add_column
- [ ] "2번째 행 삭제해줘" → hwp_table_delete_row
- [ ] "3번째 열 삭제해줘" → hwp_table_delete_column
- [ ] "셀 분할해줘" → hwp_table_split_cell
- [ ] "셀 너비 균등하게" → hwp_table_distribute_width

### 5-6. 표 테두리

- [ ] "전체 테두리 실선 0.5pt" → hwp_set_table_border
- [ ] "테두리 색상 #333333" → color 적용 확인
- [ ] 특정 셀만 테두리 변경

### 5-7. 표 데이터 조회

- [ ] "표 내용 보여줘" → hwp_get_tables
- [ ] "셀 탭 인덱스 매핑해줘" → hwp_map_table_cells
- [ ] "셀 서식 확인해줘" → hwp_get_cell_format
- [ ] "표 서식 요약해줘" → hwp_get_table_format_summary

### 5-8. 표 수식/내보내기

- [ ] "3열 합계 구해줘" → hwp_table_formula_sum
- [ ] "3열 평균 구해줘" → hwp_table_formula_avg
- [ ] "표를 CSV로 내보내줘" → hwp_table_to_csv
- [ ] "표를 JSON으로 내보내줘" → hwp_table_to_json

---

## Phase 6: 페이지/레이아웃 도구 (9개)

### 6-1. 페이지 설정

- [ ] "여백 20mm로 설정해줘" → hwp_set_page_setup
- [ ] "용지 가로 방향으로" → orientation=landscape
- [ ] "B4 용지로 변경해줘" → paper_width/height

### 6-2. 머리글/바닥글

- [ ] "머리글에 '대외비' 넣어줘" → hwp_set_header_footer type=header
  - [ ] 타임아웃 없이 정상 완료 확인
- [ ] "바닥글에 기관 정보 넣어줘" → type=footer
- [ ] verify_layout에서 머리글/바닥글 표시 확인

### 6-3. 다단/섹션

- [ ] "2단으로 설정해줘" → hwp_set_column count=2
- [ ] "구분선 포함 2단" → line_type=1
- [ ] "섹션 나눠줘" → hwp_break_section
- [ ] "다단 나눠줘" → hwp_break_column

### 6-4. 이미지

- [ ] "이미지 삽입해줘" → hwp_insert_picture
- [ ] "배경 이미지 설정해줘" → hwp_set_background_picture
- [ ] "문서 내 이미지 추출해줘" → hwp_image_extract

### 6-5. 시각 검증

- [ ] "현재 레이아웃 확인해줘" → hwp_verify_layout
  - [ ] PDF→PNG 변환 정상
  - [ ] Claude가 PNG 이미지를 읽고 분석하는지
  - [ ] 특정 페이지만 확인 (pages="1")

---

## Phase 7: 서식/그리기 도구 (5개)

### 7-1. 스타일

- [ ] "본문 스타일 적용해줘" → hwp_apply_style style_name="본문"
- [ ] "제목1 스타일 적용해줘" → style_name="제목1"

### 7-2. 셀 속성

- [ ] "셀 수직 가운데 정렬" → hwp_set_cell_property vert_align=middle
- [ ] "셀 여백 3mm" → margin_left/right/top/bottom

### 7-3. 그리기

- [ ] "글상자 만들어줘" → hwp_insert_textbox
- [ ] "선 그려줘" → hwp_draw_line
  - [ ] 두께/색상/스타일 적용
- [ ] "선 삽입해줘" → hwp_insert_line

### 7-4. 캡션

- [ ] "표 아래에 캡션 넣어줘" → hwp_insert_caption
  - [ ] "[표 1] 제목" 형태

---

## Phase 8: 문서 분석 도구 (16개)

### 8-1. 분석

- [ ] "문서 전체 분석해줘" → hwp_analyze_document
- [ ] "AI로 심층 분석해줘" → hwp_smart_analyze
- [ ] "문서 요약해줘" → hwp_document_summary

### 8-2. 텍스트 추출

- [ ] "전체 텍스트 보여줘" → hwp_get_document_text
- [ ] "2페이지 텍스트만" → hwp_get_page_text
- [ ] "마크다운으로 변환해줘" → hwp_get_as_markdown
- [ ] "글자수 세어줘" → hwp_word_count

### 8-3. 필드/양식

- [ ] "필드 목록 보여줘" → hwp_get_fields
- [ ] "양식 빈칸 감지해줘" → hwp_form_detect
- [ ] "서식 프로파일 추출해줘" → hwp_extract_style_profile

### 8-4. 기타

- [ ] "두 문서 비교해줘" → hwp_compare_documents
- [ ] "문서 병합해줘" → hwp_document_merge
- [ ] "문서 분할해줘" → hwp_document_split

---

## Phase 9: 스마트/복합 도구 (16개)

### 9-1. AI 자동화

- [ ] "서식 보존하며 채워줘" → hwp_smart_fill
- [ ] "엑셀 읽어서 자동 채워줘" → hwp_auto_fill_from_reference
- [ ] "참고자료 헤더 매핑해줘" → hwp_auto_map_reference
- [ ] "작성요령 텍스트 삭제해줘" → hwp_delete_guide_text

### 9-2. 일괄/복합

- [ ] "다건 문서 생성해줘" → hwp_generate_multi_documents
- [ ] "목차 만들어줘" → hwp_generate_toc
- [ ] "간트차트 만들어줘" → hwp_create_gantt_chart
- [ ] "일괄 변환해줘" → hwp_batch_convert

### 9-3. 기타

- [ ] "개인정보 스캔해줘" → hwp_privacy_scan
- [ ] "체크박스 체크해줘" → hwp_toggle_checkbox
- [ ] "필드에 값 채워줘" → hwp_fill_fields
- [ ] "[개발용] COM 속성 확인" → hwp_inspect_com_object

---

## Phase 10: 참고자료 읽기

### 10-1. 기본 형식

- [ ] "엑셀 파일 읽어줘" → hwp_read_reference (.xlsx)
- [ ] "CSV 파일 읽어줘" → hwp_read_reference (.csv)
- [ ] "텍스트 파일 읽어줘" → hwp_read_reference (.txt)
- [ ] "JSON 파일 읽어줘" → hwp_read_reference (.json)

### 10-2. PDF/DOCX/PPTX (신규)

- [ ] "PDF 파일 읽어줘" → hwp_read_reference (.pdf) — PyMuPDF 텍스트 추출
- [ ] "Word 파일 읽어줘" → hwp_read_reference (.docx) — PDF 변환 또는 직접 추출
- [ ] "PPT 파일 읽어줘" → hwp_read_reference (.pptx) — PDF 변환 또는 직접 추출
- [ ] 비지원 확장자 시 에러 메시지 + 안내 확인

---

## Phase 11: HWPX 전용 도구 (4개, 한글 없이 동작)

- [ ] "템플릿 목록 보여줘" → hwp_template_list
- [ ] "빈 HWPX 문서 생성해줘" → hwp_document_create
- [ ] "템플릿으로 문서 생성해줘" → hwp_template_generate
- [ ] "HWPX XML 텍스트 편집해줘" → hwp_xml_edit_text

---

## Phase 12: Hooks 동작

### 12-1. pre-tool-use (코드 작성 시 규칙 검증)

- [ ] fast-xml-parser 사용 시도 → 차단 메시지
- [ ] element.tagName 사용 시도 → 차단 (localName 권장)
- [ ] raw win32com 사용 시도 → 차단 (pyhwpx 권장)

### 12-2. post-tool-use (에러 후 복구 안내)

- [ ] COM 에러 발생 시 → 복구 가이드 표시
- [ ] HWPX 편집 후 → linesegarray 삭제 확인 메시지

---

## Phase 13: Agent 자동 트리거

### 13-1. 키워드 감지

- [ ] "한글 문서 작성해줘" → hwp-assistant 자동 작동
- [ ] "hwp 파일 열어줘" → 환경 체크 실행
- [ ] "양식 채워줘" → /hwp-fill 추천
- [ ] "사업계획서" → /hwp-write 추천
- [ ] "동의서 만들어줘" → /hwp-write 추천
- [ ] "위촉장 100건" → /hwp-batch 추천
- [ ] "개인정보 확인해줘" → /hwp-privacy 추천

---

## Phase 14: 종합 시나리오 테스트

### 14-1. 공문서(인사발령) 생성

1. [ ] /hwp-write → "공문서" 선택
2. [ ] 새 HWPX 문서 생성
3. [ ] 페이지 설정 (여백 20mm)
4. [ ] 머리글 "대외비"
5. [ ] 기관명 삽입 (가운데, 큰 글씨, 진남색)
6. [ ] 문서번호 + 시행일
7. [ ] 수신 + 제목
8. [ ] 본문 (순번체계: 1. 가. 1))
9. [ ] 인사발령 표 생성
   - [ ] col_widths 적용
   - [ ] 헤더행 진남색 배경 + 흰색 글자 + 가운데 정렬
   - [ ] 테두리 적용
10. [ ] 합계행 셀 병합
11. [ ] 발령 일자/조건 (가. 나. 다.)
12. [ ] 붙임
13. [ ] 발신명의 (가운데)
14. [ ] 결재란 표
15. [ ] 바닥글 (기관 연락처)
16. [ ] verify_layout → PDF 시각 확인
17. [ ] 저장 (HWPX + PDF)

### 14-2. 양식 채우기 (Excel 참조)

1. [ ] /hwp-fill → 양식 파일 경로 제공
2. [ ] Excel 참고자료 경로 제공
3. [ ] "참조 데이터 .md 저장" 선택
4. [ ] 자동 분석 (smart_analyze)
5. [ ] 미리보기 → 사용자 확인
6. [ ] auto_fill_from_reference로 자동 채우기
7. [ ] verify_layout 확인
8. [ ] 저장 + PDF 변환

### 14-3. 위촉장 100건 일괄 생성

1. [ ] /hwp-batch → 양식 파일 제공
2. [ ] Excel 데이터 파일 제공 (이름/소속/직급 등)
3. [ ] 출력 폴더 지정
4. [ ] 파일명 규칙: "위촉장\_이름.hwp"
5. [ ] 일괄 생성 실행
6. [ ] 결과 파일 수 확인
7. [ ] PDF 일괄 변환 (선택)

### 14-4. 사업계획서 (긴 문서, 10쪽+)

1. [ ] /hwp-write → "사업계획서" 선택
2. [ ] "분량 10쪽" → 섹션별 분할 작성 안내
3. [ ] 목차 구성 → 사용자 확인
4. [ ] 섹션 1: 사업 개요 (2쪽) → verify_layout
5. [ ] 섹션 2: 추진 전략 (2쪽) → verify_layout
6. [ ] 섹션 3: 추진일정 (간트차트 포함)
7. [ ] 섹션 4: 소요예산 (표 + 합계)
8. [ ] 섹션 5: 기대효과
9. [ ] 목차 자동 생성 (generate_toc)
10. [ ] 최종 verify_layout → 전체 확인
11. [ ] 저장

### 14-5. 자연어 요청 시나리오 모음

- [ ] "C:/문서/사업계획서.hwp 파일 열어서 분석해줘"
- [ ] "표의 계약금액 칸에 50,000,000원을 채워줘"
- [ ] "문서를 PDF로 변환해줘"
- [ ] "작성요령 텍스트 삭제해줘"
- [ ] "참고자료.xlsx를 읽어서 양식.hwp의 표를 자동으로 채워줘"
- [ ] "직원\_명단.xlsx의 각 행으로 위촉장.hwp를 개별 생성해줘"
- [ ] "'갑'을 '을'로 전체 바꿔줘"
- [ ] "3페이지 텍스트만 보여줘"
- [ ] "개인정보 있는지 스캔해줘"
- [ ] "현재 문서 레이아웃 확인해줘"
- [ ] "표에 행 하나 추가해줘"
- [ ] "셀 배경색 파란색으로 바꿔줘"
- [ ] "머리글에 '비밀' 넣어줘"
- [ ] "2단 레이아웃으로 바꿔줘"
- [ ] "글자수 세어줘"
- [ ] "마크다운으로 변환해줘"
- [ ] "표를 CSV로 내보내줘"
- [ ] "이미지 삽입해줘"
- [ ] "각주 넣어줘: 출처 표시"
- [ ] "여백 15mm로 줄여줘"

---

## Phase 15: 엣지 케이스 + 안정성 테스트

### 15-1. 에러 처리

- [ ] 한글 프로그램 미실행 상태에서 도구 호출 → 에러 메시지 + 안내
- [ ] 존재하지 않는 파일 경로 → FileNotFoundError 메시지
- [ ] 표가 없는 문서에서 fill_table_cells 호출 → 에러 처리
- [ ] 잘못된 table_index (범위 초과) → 에러 메시지
- [ ] 빈 텍스트 insert_text("") → 빈 줄 생성 확인

### 15-2. HWPX vs HWP 동작 차이

- [ ] HWPX 파일에서 find_replace → XML 라우팅 동작 확인
- [ ] HWP 파일에서 find_replace → COM 경로 동작 확인
- [ ] HWPX에서 text_search → XML 검색 확인
- [ ] HWP에서 text_search → COM 검색 (제한적)

### 15-3. 대용량/특수 케이스

- [ ] 50행 이상 큰 표 생성 → 정상 동작
- [ ] 10페이지 이상 문서 → 페이지별 텍스트 추출
- [ ] 한글+영문+숫자 혼합 텍스트 → 인코딩 정상
- [ ] 특수문자 포함 텍스트 (따옴표, 괄호 등) → 정상 삽입
- [ ] 경로에 한글 포함 (C:/문서/사업계획서.hwp) → 정상 열기
- [ ] 동시에 여러 표 있는 문서 → table_index로 구분 정상

### 15-4. 연속 작업 안정성

- [ ] 문서 열기 → 편집 → 저장 → 닫기 → 다른 문서 열기 → 연속 동작
- [ ] 같은 도구 10회 연속 호출 → 안정성 확인
- [ ] insert_text 20회 연속 → 모두 별도 줄로 분리
- [ ] 표 생성 → 배경색 → 정렬 → 병합 → 테두리 → 순서대로 정상

### 15-5. verify_layout 검증 항목

- [ ] 1페이지 문서 → PNG 1장 생성
- [ ] 3페이지 문서 → pages="1-3" 지정 → PNG 3장
- [ ] 표 배경색이 PDF/PNG에서 보이는지 확인
- [ ] 머리글/바닥글이 PDF에서 보이는지 확인
- [ ] 셀 병합이 PDF에서 올바르게 표시되는지 확인
- [ ] 텍스트 정렬(가운데)이 PDF에서 반영되는지 확인

---

## Phase 16: v0.5.3 신규 기능 테스트

### 16-1. 구조화된 에러 응답
- [ ] 한글 미실행 상태에서 도구 호출 → error_type: "com_disconnected" + guide 확인
- [ ] 없는 파일 경로 → error_type: "file_not_found" + guide 확인
- [ ] 문서 안 열고 편집 → error_type: "no_document" + guide 확인
- [ ] 잠긴 파일 접근 → error_type: "file_locked" + guide 확인

### 16-2. 머리글/바닥글 스타일
- [ ] 머리글 bold + center: `/hwp-studio:hwp-help` 후 "머리글을 굵게 가운데로 넣어줘"
- [ ] 바닥글 font_size=8 + right: "바닥글 8pt 오른쪽 정렬로"
- [ ] verify_layout으로 머리글/바닥글 스타일 시각 확인

### 16-3. 검색 대소문자 무시
- [ ] "ABC" 입력 후 find_replace("abc", "XYZ", case_sensitive=false) → 치환 성공
- [ ] case_sensitive=true (기본) → "abc"는 못 찾음 확인

### 16-4. col_widths 자동 축소
- [ ] table_create_from_data col_widths=[100, 100, 100] → 경고 + 자동 축소 확인
- [ ] col_widths=[30, 40, 50, 40] (합계 160mm) → 경고 없이 정상

### 16-5. find_and_append (이전 FAIL → 수정)
- [ ] 텍스트 삽입 후 find_and_append → "found" 반환 확인

### 16-6. get_tables (이전 FAIL → 수정)
- [ ] 표 생성 후 get_tables → total_count > 0 확인
- [ ] analyze_document → 표 감지 확인

### 16-7. 임시 파일 정리
- [ ] verify_layout 실행 후 %TEMP%에 hwp_verify_layout.pdf 없는지 확인
- [ ] PNG 파일만 남아있는지 확인

---

## Phase 17: v0.5.4 근본 수정 검증

### 17-1. get_tables (4회 연속 FAIL → v0.5.4 수정)
- [ ] 새 문서에서 표 생성 → hwp_get_tables → total_count > 0 확인
- [ ] hwp_analyze_document → tables 배열에 표 데이터 있는지 확인
- [ ] hwp_map_table_cells → 표 셀 매핑 정상 (기존 PASS 유지)
- [ ] 표 2개 생성 → get_tables → total_count = 2 확인

### 17-2. find_and_append (4회 연속 FAIL → v0.5.4 수정)
- [ ] 텍스트 "테스트문장" 삽입 → hwp_find_and_append("테스트문장", " 추가됨") → status: "ok"
- [ ] 결과 확인: "테스트문장 추가됨"으로 변경되었는지
- [ ] HWPX 파일에서 find_and_append → XML 경로 정상 동작
- [ ] HWP 파일에서 find_and_append → COM 경로 정상 동작

### 17-3. 머리글 성능 (20.4초 퇴보 → v0.5.4 최적화)
- [ ] hwp_set_header_footer(type="header", text="테스트") → 15초 이내 완료
- [ ] style 없이 호출 시 10초대 복귀 확인
- [ ] style={bold:true, align:"center"} 포함 시에도 20초 미만

### 17-4. insert_textbox 위치/크기 (4회 연속 WARN → v0.5.4 수정)
- [ ] hwp_insert_textbox(x=50, y=50, width=80, height=30, text="글상자") → method 확인
- [ ] "fallback" 아닌 정상 경로로 생성되는지
- [ ] verify_layout에서 글상자 위치 시각 확인

### 17-5. 표 카운트 정확도 (50개 중복 → v0.5.4 수정)
- [ ] 표 1개 문서 → analyze_document → total_count = 1 (50이 아닌지)
- [ ] 표 3개 문서 → total_count = 3

---

## 발견된 이슈 기록

| #   | Phase | 이슈 | 심각도 | 상태 | 비고 |
| --- | ----- | ---- | ------ | ---- | ---- |
|     |       |      |        |      |      |

---

## 테스트 환경

- Windows 11
- Python 3.8+ / pyhwpx 1.7+
- 한글 2014 이상
- Claude Code (최신)
- 플러그인: claude-code-hwp-plugin v0.5.3
- MCP: claude-code-hwp-mcp v0.5.3 (플러그인에 포함)

## 테스트 이력

| 회차 | 버전 | 통과율 | 핵심 개선 |
|------|------|--------|----------|
| 1차 | v0.5.0 | 75.6% | 기본 기능 검증 |
| 2차 | v0.3.1 | 87.1% | SetMessageBoxMode, 스타일, 바닥글 |
| 3차 | v0.5.2 | 88.7% | PDF 정상화, 표 배경색, DOCX/HTML |
| 4차 | v0.5.3 | 82.6% (캐시 미갱신) | Phase 16 미구현으로 하락 (기존 90%) |
| 5차 | v0.5.4 | 목표 95%+ | get_tables/find_and_append 근본수정, textbox 좌표 |
