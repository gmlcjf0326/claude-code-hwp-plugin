# claude-code-hwp-plugin

Claude Code 전용 한글(HWP/HWPX) 문서 자동화 플러그인입니다.

이 플러그인만 설치하면 MCP 서버(94개 도구)가 자동으로 포함됩니다. MCP를 별도로 설치할 필요가 없습니다.

> Windows 전용 | 한글 2014 이상 | Python 3.8+

---

## 이 플러그인이 제공하는 것

| 구분 | 내용 |
|------|------|
| MCP 도구 94개 | 문서 열기/저장, 표 편집, 텍스트 검색/치환, 서식, 분석 등 |
| 워크플로우 커맨드 8개 | /hwp-fill, /hwp-write, /hwp-analyze 등 |
| HWPX 규칙 hooks 2개 | fast-xml-parser 차단, tagName 차단, linesegarray 확인 |
| 에러 복구 가이드 | COM 에러, 파일 잠금, Python 미설치 시 자동 안내 |
| 자동 트리거 agent | "한글", "hwp", "양식" 등 감지 시 자동 환경 체크 |

---

## 사전 요구사항

| 요구사항 | 설치 방법 |
|----------|-----------|
| Windows 10/11 | macOS/Linux 미지원 (COM API 기반) |
| 한글(HWP) 프로그램 | 한컴오피스 한글 2014, 2018, 2020, 2022, 2024 |
| Python 3.8+ | [python.org](https://www.python.org/downloads/)에서 설치. Microsoft Store 버전 X |
| pyhwpx | `pip install pyhwpx pywin32` |

## 설치

```bash
claude plugins marketplace add gmlcjf0326/claude-code-hwp-plugin
claude plugins install claude-code-hwp-plugin@hwp-marketplace
```

설치 확인:

```bash
claude plugins list
```

`claude-code-hwp-plugin@hwp-marketplace`가 `enabled` 상태이면 성공입니다.

## 시작하기

1. **한글(HWP) 프로그램을 실행**합니다 (빈 문서 열기)
2. Claude Code에서 `/hwp-setup`으로 환경을 확인합니다
3. `/hwp-help`로 사용 가능한 기능을 확인합니다

---

## 사용 예시

### 양식 채우기

```
나: "C:\문서\과업지시서.hwp 양식에 내용을 채워줘"

Claude:
  1. 환경 확인 (Python, 한글 실행 상태)
  2. "참고할 자료가 있나요?" 질문
  3. 문서 열기 → 구조 분석 (표 3개, 빈칸 12개 발견)
  4. "다음과 같이 채울 예정입니다:" → 미리보기 제시
  5. 사용자 확인 후 채우기 실행
  6. 개인정보 확인 → 저장
```

### 문서 작성

```
나: "사업계획서를 작성해줘. 양식은 C:\양식\사업계획서.hwp"

Claude:
  1. 양식 열기 → 서식 추출 (글꼴: 맑은고딕, 크기: 11pt)
  2. "문체는 개괄식/격식체 중?" 질문
  3. 작성요령 텍스트 삭제
  4. 제목, 본문, 표 순서대로 작성 (서식 보존)
  5. 추진일정표(간트차트) 자동 생성
  6. 저장
```

### 텍스트 치환

```
나: "이 문서에서 '대전도시공사'를 '대전광역시도시공사'로 전부 바꿔줘"

Claude:
  → HWPX 파일이면 XML 직접 치환 (안정적)
  → HWP 파일이면 COM 치환 + 결과 검증
  → "5건 치환 완료"
```

### 엑셀 데이터로 다건 생성

```
나: "직원명단.xlsx를 읽어서 위촉장.hwp 양식으로 각각 생성해줘"

Claude:
  → 엑셀 읽기 (30명 데이터)
  → 위촉장 템플릿으로 30건 개별 문서 생성
  → "30개 파일 생성 완료: 위촉장_홍길동.hwp ..."
```

### 개인정보 스캔

```
나: "이 문서에 개인정보가 있는지 확인해줘"

Claude:
  → "주민등록번호 2건, 전화번호 3건 발견"
  → "마스킹 처리할까요?"
```

### PDF 변환

```
나: "이 문서를 PDF로 변환해줘"

Claude:
  → "C:\문서\과업지시서.pdf 저장 완료 (261KB)"
```

### 문서 분석

```
나: "이 문서 내용을 요약해줘"

Claude:
  → "6페이지, 표 5개, 필드 0개"
  → "과업지시서: 대전도시공사 챗봇 솔루션 임차 용역..."
  → "작성 완성도: 85%"
```

---

## 커맨드 (8개)

### /hwp-help — 기능 안내

**언제 사용하나요?** 플러그인을 처음 설치했을 때, 또는 어떤 기능이 있는지 모를 때.

모든 커맨드 목록과 각각의 용도를 한눈에 보여줍니다. 한글 프로그램 실행 여부, HWP/HWPX 차이점도 안내합니다.

```
/hwp-help
→ "사용 가능한 커맨드 8개를 안내합니다..."
```

---

### /hwp-setup — 환경 진단

**언제 사용하나요?**
- 플러그인을 처음 설치한 직후
- 도구 실행 시 에러가 발생할 때
- Python이나 한글 프로그램 설치를 변경한 후

Python, pyhwpx, 한글 프로그램의 설치 상태를 자동 진단합니다. 문제가 있으면 단계별 해결 방법을 안내합니다.

| 진단 항목 | 정상 시 | 문제 시 안내 |
|-----------|---------|-------------|
| Python | 버전 + 경로 표시 | python.org 설치 안내 |
| Microsoft Store Python | 경고 표시 | python.org 재설치 권장 |
| pyhwpx | 버전 표시 | `pip install pyhwpx pywin32` 안내 |
| 한글 프로그램 | 설치 확인 | 한컴오피스 설치 안내 |
| 한글 실행 여부 | 실행 중 확인 | "한글을 먼저 실행하세요" |

```
/hwp-setup
→ "✅ Python 3.13.12, ✅ pyhwpx 1.7.1, ✅ 한글 실행 중 — 준비 완료!"
```

---

### /hwp-fill — 양식 자동 채우기

**언제 사용하나요?**
- 사업계획서, 과업지시서, 신청서 등의 빈칸을 채울 때
- 엑셀 데이터를 한글 양식에 자동으로 매핑할 때
- 기존 양식의 서식(글꼴, 크기)을 유지하면서 내용을 넣을 때

AI가 사전 질문 후 문서를 분석하고, 빈칸을 자동으로 채웁니다. 채우기 전에 미리보기를 보여주고, 개인정보도 자동 확인합니다.

**워크플로우**: 사전질문 → 문서열기 → 분석 → 서식파악 → 셀매핑 → 미리보기 → 채우기 → 개인정보확인 → 저장

**규칙**: 결재란 안 채움, 기존 내용 유지, 임의 수치 생성 안 함

```
/hwp-fill
→ "파일 경로를 알려주세요"
→ "참고자료가 있나요?"
→ 분석 → 미리보기 → 채우기 → 저장
```

---

### /hwp-write — 문서 작성

**언제 사용하나요?**
- 빈 문서에 처음부터 내용을 작성할 때
- 양식은 있는데 내용이 전혀 없을 때
- 사업계획서, 보고서, 공문, 기안서, 제안서를 새로 써야 할 때

AI가 문서 유형과 양식을 확인한 후, 서식을 따라서 제목/본문/표를 작성합니다.

**워크플로우**: 문서유형확인 → 양식열기 → 서식추출 → 작성요령삭제 → 내용작성 → 간트차트/목차 → 저장

```
/hwp-write
→ "어떤 문서를 작성할까요? (사업계획서/보고서/공문/기안/제안서)"
→ "양식 파일이 있나요?"
→ 서식 추출 → 작성 → 저장
```

---

### /hwp-analyze — 문서 분석

**언제 사용하나요?**
- 받은 문서의 내용을 빠르게 파악하고 싶을 때
- 문서의 완성도(빈칸이 얼마나 있는지)를 확인할 때
- 표 구조와 데이터를 확인할 때
- 문서 작업 전에 구조를 먼저 파악할 때

문서를 심층 분석하여 종류, 페이지수, 표수, 완성도, 빈 항목 목록, 내용 요약을 보고합니다.

```
/hwp-analyze
→ "6페이지, 표 5개, 필드 0개"
→ "과업지시서: 대전도시공사 챗봇 솔루션 임차 용역"
→ "작성 완성도: 85%, 빈 항목: 계약금액, 착수일"
```

---

### /hwp-convert — 형식 변환

**언제 사용하나요?**
- 한글 문서를 PDF로 변환하여 배포/제출할 때
- Word(DOCX)로 변환하여 다른 사람에게 보낼 때
- HWP를 HWPX로 변환하여 텍스트 작업 안정성을 높일 때

| 형식 | 용도 | 비고 |
|------|------|------|
| PDF | 배포/인쇄/제출 | 가장 많이 사용 |
| DOCX | Word 호환 | 시간이 걸릴 수 있음 |
| HTML | 웹 게시 | |
| HWPX | XML 기반 | 텍스트 작업이 더 안정적 |

```
/hwp-convert
→ "어떤 형식으로 변환할까요?"
→ "PDF로 변환 완료: C:\문서\과업지시서.pdf (261KB)"
```

---

### /hwp-batch — 다건 문서 생성

**언제 사용하나요?**
- 위촉장, 증명서, 안내문 등을 여러 명에게 개별 발급할 때
- 엑셀 명단의 각 행으로 별도 HWP 파일을 만들어야 할 때
- 같은 양식을 데이터만 바꿔서 반복 생성할 때

```
/hwp-batch
→ "템플릿 파일: 위촉장.hwp"
→ "데이터 파일: 직원명단.xlsx (30명)"
→ "30개 파일 생성 완료: 위촉장_홍길동.hwp, 위촉장_김철수.hwp ..."
```

---

### /hwp-privacy — 개인정보 스캔

**언제 사용하나요?**
- 문서를 외부에 제출하기 전에 개인정보가 포함되었는지 확인할 때
- 개인정보보호법 준수를 위해 문서를 점검할 때
- 주민번호, 전화번호 등을 자동으로 마스킹 처리하고 싶을 때

| 감지 유형 | 패턴 예시 |
|-----------|----------|
| 주민등록번호 | 000000-0000000 |
| 전화번호 | 010-0000-0000 |
| 이메일 | name@example.com |
| 계좌번호 | 은행 계좌 패턴 |
| 여권번호 | 여권 패턴 |

```
/hwp-privacy
→ "주민등록번호 2건, 전화번호 3건 발견"
→ "마스킹 처리할까요?"
→ "마스킹 완료 → 저장"
```

---

## MCP 도구 전체 목록 (94개)

플러그인 설치 시 자동으로 사용 가능합니다.

### 환경/문서 관리 (6개)

| 도구 | 설명 |
|------|------|
| hwp_check_setup | Python/pyhwpx/한글 설치 상태 자동 진단 |
| hwp_list_files | 디렉토리 내 HWP/HWPX 파일 목록 |
| hwp_open_document | 문서 열기 (자동 백업) |
| hwp_close_document | 문서 닫기 |
| hwp_save_document | 저장 (HWP/HWPX/PDF/DOCX) |
| hwp_export_pdf | PDF 내보내기 |

### 문서 분석 (16개)

| 도구 | 설명 |
|------|------|
| hwp_analyze_document | 전체 구조 분석 |
| hwp_smart_analyze | AI용 심층 분석 (문서 유형 추론) |
| hwp_get_document_text | 전문 텍스트 추출 |
| hwp_get_document_info | 메타데이터 (페이지수 등) |
| hwp_get_tables | 표 데이터 조회 |
| hwp_map_table_cells | 셀 탭 인덱스 매핑 (병합 셀 대응) |
| hwp_get_cell_format | 특정 셀 서식 정보 |
| hwp_get_table_format_summary | 표 서식 요약 |
| hwp_get_fields | 양식 필드 목록 |
| hwp_get_as_markdown | 마크다운 변환 |
| hwp_get_page_text | 특정 페이지 텍스트 |
| hwp_text_search | 텍스트 검색 |
| hwp_form_detect | 양식 빈칸/체크박스 감지 |
| hwp_extract_style_profile | 서식 프로파일 추출 |
| hwp_image_extract | 이미지 추출 |
| hwp_read_reference | 참고자료 읽기 (Excel/CSV/TXT/JSON) |

### 텍스트 편집 (18개)

| 도구 | 설명 |
|------|------|
| hwp_insert_text | 텍스트 삽입 (색상/볼드/서식) |
| hwp_insert_markdown | 마크다운을 HWP 서식으로 변환 삽입 |
| hwp_insert_heading | 제목 삽입 (H1~H6 + 자동 순번) |
| hwp_find_replace | 찾기/바꾸기 (HWPX: XML 직접) |
| hwp_find_replace_multi | 다건 찾기/바꾸기 |
| hwp_find_replace_nth | N번째 항목만 치환 |
| hwp_find_and_append | 텍스트 뒤에 추가 |
| hwp_set_paragraph_style | 문단 서식 (정렬, 줄간격) |
| hwp_indent | 들여쓰기 |
| hwp_outdent | 내어쓰기 |
| hwp_insert_page_break | 페이지 나누기 |
| hwp_insert_page_num | 쪽 번호 삽입 |
| hwp_insert_date_code | 날짜 자동 삽입 |
| hwp_insert_footnote | 각주 |
| hwp_insert_endnote | 미주 |
| hwp_insert_hyperlink | 하이퍼링크 |
| hwp_insert_auto_num | 자동 번호매기기 |
| hwp_insert_memo | 메모 |

### 표 편집 (18개)

| 도구 | 설명 |
|------|------|
| hwp_fill_table_cells | 표 셀 채우기 (탭/라벨/좌표) |
| hwp_fill_fields | 양식 필드 채우기 |
| hwp_smart_fill | 서식 감지 후 보존하며 채우기 |
| hwp_table_create_from_data | 2D 배열로 표 생성 |
| hwp_table_insert_from_csv | CSV/Excel에서 표 생성 |
| hwp_table_add_row | 행 추가 |
| hwp_table_add_column | 열 추가 |
| hwp_table_delete_row | 행 삭제 |
| hwp_table_delete_column | 열 삭제 |
| hwp_table_merge_cells | 셀 병합 |
| hwp_table_split_cell | 셀 분할 |
| hwp_table_distribute_width | 너비 균등 분배 |
| hwp_table_swap_type | 행/열 교환 |
| hwp_table_formula_sum | 합계 수식 |
| hwp_table_formula_avg | 평균 수식 |
| hwp_table_to_csv | CSV 내보내기 |
| hwp_table_to_json | JSON 내보내기 |
| hwp_set_cell_color | 셀 배경색 |

### 스마트/복합 (12개)

| 도구 | 설명 |
|------|------|
| hwp_auto_fill_from_reference | Excel에서 자동 매핑 후 채우기 |
| hwp_auto_map_reference | 참고자료 헤더와 표 라벨 매핑 |
| hwp_generate_multi_documents | 다건 문서 일괄 생성 |
| hwp_generate_toc | 목차 자동 생성 |
| hwp_create_gantt_chart | 간트차트 추진일정 표 |
| hwp_document_merge | 여러 문서 병합 |
| hwp_document_split | 문서 분할 |
| hwp_document_summary | 문서 요약 정보 |
| hwp_privacy_scan | 개인정보 자동 감지 |
| hwp_batch_convert | HWP 일괄 변환 |
| hwp_compare_documents | 두 문서 비교 |
| hwp_word_count | 글자수/단어수/페이지수 |

### 페이지/레이아웃 (9개)

| 도구 | 설명 |
|------|------|
| hwp_set_page_setup | 여백, 용지 크기, 방향(가로/세로) 설정 |
| hwp_set_header_footer | 머리글/바닥글 삽입 |
| hwp_set_column | 다단 설정 (2단/3단, 구분선) |
| hwp_verify_layout | PDF→PNG 시각 검증 (PyMuPDF) |
| hwp_insert_picture | 이미지 삽입 |
| hwp_set_background_picture | 배경 이미지 |
| hwp_insert_line | 선 삽입 |
| hwp_break_section | 섹션 나누기 |
| hwp_break_column | 다단 나누기 |

### 서식/그리기 (5개)

| 도구 | 설명 |
|------|------|
| hwp_apply_style | 문단 스타일 적용 ("제목1", "본문" 등) |
| hwp_set_cell_property | 셀 여백/수직정렬/텍스트방향/보호 |
| hwp_insert_textbox | 글상자 생성 (위치/크기 지정) |
| hwp_draw_line | 선 그리기 (두께/색상/스타일) |
| hwp_insert_caption | 표/그림 캡션 삽입 |

### 기타 (7개)

| 도구 | 설명 |
|------|------|
| hwp_set_table_border | 표 테두리 스타일 |
| hwp_delete_guide_text | 작성요령 삭제 |
| hwp_toggle_checkbox | 체크박스 전환 |
| hwp_export_docx | DOCX 내보내기 |
| hwp_export_html | HTML 내보내기 |
| hwp_inspect_com_object | [개발용] COM 속성 조회 |
| hwp_word_count | 글자수/단어수/페이지수 |

### HWPX 전용 (4개, 한글 프로그램 없이 동작)

| 도구 | 설명 |
|------|------|
| hwp_template_list | 문서 템플릿 목록 (22종) |
| hwp_document_create | 빈 HWPX 문서 생성 |
| hwp_template_generate | 템플릿 기반 문서 생성 |
| hwp_xml_edit_text | HWPX XML 직접 텍스트 편집 |

---

## Hooks — 자동 규칙 검증

### 코드 작성 시 자동 차단 (pre-tool-use)

| 규칙 | 차단 대상 | 올바른 사용법 |
|------|-----------|-------------|
| fast-xml-parser 금지 | npm install fast-xml-parser | @xmldom/xmldom 사용 |
| .tagName 금지 | element.tagName | element.localName 사용 |
| charPrIDRef 변경 금지 | charPrIDRef 수정 | 문서 깨짐 방지 |
| raw win32com 금지 | import win32com | from pyhwpx import Hwp |

### 에러 발생 시 자동 안내 (post-tool-use)

| 에러 | 자동 안내 |
|------|-----------|
| RPC/COM 에러 | "한글을 종료하고 다시 실행하세요" |
| 파일 잠금 (EBUSY) | "한글에서 파일을 닫고 다시 시도하세요" |
| Python 미설치 | "python.org에서 설치하세요" |
| 파일 경로 오류 | "hwp_list_files로 검색하세요" |
| 문서 미열기 | "hwp_open_document로 열어주세요" |
| 타임아웃 | "대용량 문서는 시간이 걸립니다" |
| HWP find_replace 실패 | "HWPX로 변환하면 안정적입니다" |

---

## HWP vs HWPX

| 구분 | HWP (바이너리) | HWPX (XML 기반) |
|------|--------------|-----------------|
| 텍스트 검색 | COM API (제한적) | XML 직접 검색 (안정적) |
| 찾기/바꾸기 | COM API (제한적) | XML 직접 치환 (안정적) |
| 표 편집 | COM API | COM API |
| 문서 열기/저장 | COM API | COM API |

**HWPX 사용을 권장합니다.** 한글에서 "다른 이름으로 저장" > HWPX 형식 선택.

---

## MCP 별도 설치가 필요 없는 이유

플러그인 안에 MCP 서버 코드와 Python 브릿지가 물리적으로 포함되어 있습니다. 설치 시 자동으로 MCP 서버가 등록됩니다.

## 주의: MCP와 중복 설치하지 마세요

이미 `claude-code-hwp-mcp`를 settings.json에 등록한 상태에서 이 플러그인을 설치하면 같은 도구가 2개씩 등록됩니다. **플러그인을 설치했다면 settings.json에서 hwp-studio MCP 설정을 삭제하세요.**

---

## 알려진 제한사항

- HWP 바이너리 파일에서 text_search가 0건 반환 가능 (COM API 한계)
- HWPX 파일이 한글에서 열린 상태에서 XML 편집 시 COM 자동 폴백
- Windows 전용 (한글 COM API 기반)

## 문제 해결

| 문제 | 해결 |
|------|------|
| 환경 문제 | `/hwp-setup` 실행 |
| 텍스트 치환 안 됨 | HWPX로 변환 후 재시도 |
| 한글 응답 없음 | Ctrl+Shift+Esc → Hwp.exe 종료 → 재실행 |
| 표 셀 위치 모름 | hwp_map_table_cells로 탭 인덱스 확인 |
| Microsoft Store Python | python.org 재설치 권장 |

---

## Claude Desktop을 사용하는 경우

이 플러그인은 Claude Code 전용입니다. Claude Desktop에서는 MCP 패키지를 사용하세요:

- npm: `npm install -g claude-code-hwp-mcp`
- 설정: [claude-code-hwp-mcp README](https://github.com/gmlcjf0326/claude-code-hwp-mcp)

---

## 변경 이력

### v0.5.0 (2026-03-25) — 94개 도구, 표 서식 대폭 강화

**신규 도구 9개:**

| 도구 | 설명 | 용도 |
|------|------|------|
| hwp_set_page_setup | 여백, 용지 크기, 방향 설정 | 공문서 페이지 설정 |
| hwp_set_header_footer | 머리글/바닥글 삽입 | 기관명, 페이지번호 |
| hwp_set_column | 다단 설정 (2단/3단) | 신문/뉴스레터 레이아웃 |
| hwp_verify_layout | PDF→PNG 시각 검증 | 레이아웃 확인 |
| hwp_apply_style | 문단 스타일 적용 | "제목1", "본문" 등 |
| hwp_set_cell_property | 셀 여백/수직정렬/보호 | 표 정밀 제어 |
| hwp_insert_textbox | 글상자 생성 | 결재란, 위치 지정 요소 |
| hwp_draw_line | 선 그리기 | 구분선, 장식선 |
| hwp_insert_caption | 표/그림 캡션 | "[표 1] 제목" 형식 |

**텍스트 서식 25+ 속성 추가:**
- 밑줄 7종, 취소선 4종, 위/아래첨자, 외곽선, 그림자, 양각/음각, 작은대문자
- 라틴 전용 글꼴, 그림자 색상/오프셋, 커닝

**표 핵심 개선:**
- 셀 배경색: `cell_fill()` pyhwpx 내장 메서드 (안정적 적용)
- 셀 정렬: `TableCellAlignCenterCenter` 액션 (텍스트 삽입 후 적용)
- 셀 병합: `TableCellBlockExtend` 정확한 블록 선택
- 표 생성: col_widths/row_heights(mm), alignment, header_style

**버그 수정:**
- insert_text 자동 줄바꿈 (각 호출이 독립 문단)
- header_footer CreateAction 방식 (대화상자 타임아웃 해결)
- apply_style Execute 방식 (대화상자 방지)
- draw_line 대화형 fallback 제거

### v0.3.0 — HWPX XML 라우팅 + 버그 수정
- HWPX 파일의 텍스트 검색/치환을 XML 엔진으로 직접 처리
- 10건 버그 수정 (SelectAll 파괴, find_replace 검증 등)

---

## 라이선스

MIT License

## 관련 링크

- [MCP 패키지 (Claude Desktop용)](https://github.com/gmlcjf0326/claude-code-hwp-mcp)
- [npm 패키지](https://www.npmjs.com/package/claude-code-hwp-mcp)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [MCP Protocol](https://modelcontextprotocol.io/)
- [한컴오피스](https://www.hancom.com/)
