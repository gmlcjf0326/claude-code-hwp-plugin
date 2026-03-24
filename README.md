# claude-code-hwp-plugin

Claude Code 전용 한글(HWP/HWPX) 문서 자동화 플러그인입니다.

이 플러그인만 설치하면 MCP 서버(85개+ 도구)가 자동으로 포함됩니다. MCP를 별도로 설치할 필요가 없습니다.

> Windows 전용 | 한글 2014 이상 | Python 3.8+

---

## 이 플러그인이 제공하는 것

| 구분 | 내용 |
|------|------|
| MCP 도구 85개+ | 문서 열기/저장, 표 편집, 텍스트 검색/치환, 서식, 분석 등 (번들 포함) |
| 워크플로우 커맨드 8개 | /hwp-fill, /hwp-write, /hwp-analyze 등 고수준 자동화 |
| HWPX 규칙 hooks 2개 | fast-xml-parser 차단, tagName 차단, linesegarray 삭제 확인 |
| 에러 복구 가이드 | COM 에러, 파일 잠금, Python 미설치 시 자동 안내 |
| 자동 트리거 agent | "한글", "hwp", "양식" 등 키워드 감지 시 자동 환경 체크 |

---

## 사전 요구사항

| 요구사항 | 설치 방법 |
|----------|-----------|
| Windows 10/11 | macOS/Linux 미지원 (COM API 기반) |
| 한글(HWP) 프로그램 | 한컴오피스 한글 2014, 2018, 2020, 2022, 2024 |
| Python 3.8+ | [python.org](https://www.python.org/downloads/)에서 설치. Microsoft Store 버전 X |
| pyhwpx | `pip install pyhwpx pywin32` |

## 설치

### Step 1: 마켓플레이스 등록 (최초 1회)

```bash
claude plugins marketplace add gmlcjf0326/claude-code-hwp-plugin
```

### Step 2: 플러그인 설치

```bash
claude plugins install claude-code-hwp-plugin@hwp-marketplace
```

### Step 3: 확인

```bash
claude plugins list
```

`claude-code-hwp-plugin@hwp-marketplace`가 `enabled` 상태이면 성공입니다.

### 로컬 설치 (개발자용)

```bash
git clone https://github.com/gmlcjf0326/claude-code-hwp-plugin.git
cd claude-code-hwp-plugin
npm install
claude plugins marketplace add .
claude plugins install claude-code-hwp-plugin@hwp-marketplace
```

---

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
  → hwp_find_replace 실행
  → HWPX 파일이면 XML 직접 치환 (안정적)
  → HWP 파일이면 COM 치환 + 결과 검증
  → "5건 치환 완료"
```

### 엑셀 데이터로 자동 채우기

```
나: "직원명단.xlsx를 읽어서 위촉장.hwp 양식으로 각각 생성해줘"

Claude:
  → 엑셀 읽기 (30명 데이터)
  → 위촉장 템플릿으로 30건 개별 문서 생성
  → "30개 파일 생성 완료: C:\output\위촉장_홍길동.hwp ..."
```

### 개인정보 스캔

```
나: "이 문서에 개인정보가 있는지 확인해줘"

Claude:
  → hwp_privacy_scan 실행
  → "주민등록번호 2건, 전화번호 3건 발견"
  → "마스킹 처리할까요?"
```

### PDF 변환

```
나: "이 문서를 PDF로 변환해줘"

Claude:
  → hwp_export_pdf 실행
  → "C:\문서\과업지시서.pdf 저장 완료 (261KB)"
```

### 문서 분석

```
나: "이 문서 내용을 요약해줘"

Claude:
  → hwp_smart_analyze 실행
  → "6페이지, 표 5개, 필드 0개"
  → "과업지시서: 대전도시공사 챗봇 솔루션 임차 용역..."
  → "작성 완성도: 85%"
```

---

## 커맨드 (8개)

### /hwp-help — 기능 안내

모든 커맨드와 사용법을 안내합니다. 처음 사용할 때 실행하세요.

### /hwp-setup — 환경 진단

Python, pyhwpx, 한글 프로그램의 설치 상태를 자동 진단합니다.
문제가 있으면 단계별 해결 방법을 안내합니다.

| 진단 항목 | 정상 | 문제 시 안내 |
|-----------|------|-------------|
| Python | 버전 + 경로 표시 | python.org 설치 안내 |
| Microsoft Store Python | 경고 표시 | python.org 재설치 권장 |
| pyhwpx | 버전 표시 | pip install 명령 안내 |
| 한글 프로그램 | 설치 확인 | 한컴오피스 설치 안내 |
| 한글 실행 여부 | 실행 중 확인 | 한글 실행 요청 |

### /hwp-fill — 양식 자동 채우기

사업계획서, 과업지시서, 신청서 등의 빈칸을 AI가 자동으로 채웁니다.

**사전 질문**: 파일 경로, 참고자료, 분량, 확정 수치

**워크플로우**:
1. 환경 확인
2. 문서 열기 + 구조 분석 + 서식 파악
3. 표 셀 매핑 (병합 셀 대응)
4. 미리보기 → 사용자 확인
5. 채우기 (참고자료 있으면 자동 매핑)
6. 개인정보 확인 → 저장

**규칙**: 결재란 비움, 기존 내용 유지, 임의 수치 생성 금지

### /hwp-write — 문서 작성

AI가 한글 문서를 처음부터 작성합니다.

**사전 질문**: 문서 유형, 양식 파일, 참고자료, 문체, 분량

**워크플로우**:
1. 양식 열기 + 서식 추출
2. 작성요령 삭제
3. 제목/본문/표 작성 (서식 보존)
4. 간트차트/목차 생성 (필요 시)
5. 저장

### /hwp-analyze — 문서 분석

문서의 구조, 내용, 완성도를 분석하여 요약합니다.

**보고 항목**: 문서 종류, 페이지수, 표수, 필드수, 완성도(%), 빈 항목, 내용 요약

### /hwp-convert — 형식 변환

HWP 문서를 PDF, DOCX, HTML, HWPX로 변환합니다.

| 형식 | 용도 |
|------|------|
| PDF | 배포/인쇄용 |
| DOCX | Word 호환 |
| HTML | 웹 게시 |
| HWPX | XML 기반 (텍스트 작업 안정적) |

### /hwp-batch — 다건 문서 생성

엑셀/CSV 데이터의 각 행으로 개별 HWP 문서를 생성합니다.
위촉장, 증명서, 안내문 등 대량 생성에 사용합니다.

### /hwp-privacy — 개인정보 스캔

문서에서 자동 감지하는 개인정보:

| 유형 | 패턴 |
|------|------|
| 주민등록번호 | 000000-0000000 |
| 전화번호 | 010-0000-0000 |
| 이메일 | name@example.com |
| 계좌번호 | 은행 계좌 패턴 |
| 여권번호 | 여권 패턴 |

---

## MCP 도구 전체 목록 (85개+)

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

### 기타 (11개)

| 도구 | 설명 |
|------|------|
| hwp_set_table_border | 표 테두리 스타일 |
| hwp_insert_picture | 이미지 삽입 |
| hwp_set_background_picture | 배경 이미지 |
| hwp_insert_line | 선 삽입 |
| hwp_break_section | 섹션 나누기 |
| hwp_break_column | 다단 나누기 |
| hwp_delete_guide_text | 작성요령 삭제 |
| hwp_toggle_checkbox | 체크박스 전환 |
| hwp_export_docx | DOCX 내보내기 |
| hwp_export_html | HTML 내보내기 |
| hwp_inspect_com_object | [개발용] COM 속성 조회 |

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
| Python 미설치 | "python.org에서 설치하세요" + 링크 |
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
| 표 생성/편집 | COM API | COM API |
| 문서 열기/저장 | COM API | COM API |

**HWPX 사용을 권장합니다.** 텍스트 검색/치환이 XML을 직접 조작하여 더 안정적입니다.

변환 방법: 한글에서 "다른 이름으로 저장" > HWPX 형식 선택

---

## MCP 별도 설치가 필요 없는 이유

플러그인 안에 MCP 서버 코드와 Python 브릿지가 물리적으로 포함되어 있습니다.

```
claude-code-hwp-plugin/
├── .mcp.json           ← MCP 서버 자동 등록 설정
├── servers/            ← MCP 서버 코드 (85개+ 도구)
├── python/             ← Python COM 브릿지 (pyhwpx)
├── commands/           ← 8개 워크플로우 커맨드
├── hooks/              ← HWPX 규칙 + 에러 복구
└── agents/             ← 자동 트리거
```

플러그인 설치 시 `.mcp.json`이 MCP 서버를 자동 등록하여 85개+ 도구를 바로 사용할 수 있습니다.

## 주의: MCP와 중복 설치하지 마세요

이미 `claude-code-hwp-mcp` MCP를 settings.json에 등록한 상태에서 이 플러그인을 설치하면, 같은 도구가 2개씩 등록됩니다.

**플러그인을 설치했다면 settings.json에서 hwp-studio MCP 설정을 삭제하세요.**

---

## 알려진 제한사항

- HWP 바이너리 파일에서 text_search가 0건 반환 가능 (COM API 한계)
- HWPX 파일이 한글에서 열린 상태에서 XML 편집 시 COM 자동 폴백
- Windows 전용 (한글 COM API 기반)

## 문제 해결

| 문제 | 해결 |
|------|------|
| 환경 문제 | `/hwp-setup` 실행 → 자동 진단 + 안내 |
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

## 라이선스

MIT License

## 관련 링크

- [MCP 패키지 (Claude Desktop용)](https://github.com/gmlcjf0326/claude-code-hwp-mcp)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [MCP Protocol](https://modelcontextprotocol.io/)
- [한컴오피스](https://www.hancom.com/)
