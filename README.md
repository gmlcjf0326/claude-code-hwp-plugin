# claude-code-hwp-plugin

Claude Code 전용 한글(HWP/HWPX) 문서 자동화 플러그인입니다.

이 플러그인만 설치하면 MCP 서버(85개+ 도구)가 자동으로 포함됩니다. MCP를 별도로 설치할 필요가 없습니다.

> Windows 전용 | 한글 2014 이상 | Python 3.8+

## 이 플러그인이 제공하는 것

| 구분 | 내용 |
|------|------|
| MCP 도구 85개+ | 문서 열기/저장, 표 편집, 텍스트 검색/치환, 서식, 분석 등 (번들 포함) |
| 워크플로우 커맨드 8개 | /hwp-fill, /hwp-write, /hwp-analyze 등 고수준 자동화 |
| HWPX 규칙 hooks | fast-xml-parser 차단, tagName 차단, linesegarray 삭제 확인 |
| 에러 복구 가이드 | COM 에러, 파일 잠금, Python 미설치 시 자동 안내 |
| 자동 트리거 agent | "한글", "hwp", "양식" 등 키워드 감지 시 자동 환경 체크 |

## MCP 별도 설치가 필요 없는 이유

플러그인 안에 MCP 서버 코드(servers/)와 Python 브릿지(python/)가 **물리적으로 포함**되어 있습니다. 플러그인 설치 시 `.mcp.json` 설정으로 MCP 서버가 자동 등록되어, 85개+ 도구를 바로 사용할 수 있습니다.

```
claude-code-hwp-plugin/
├── .mcp.json           ← MCP 서버 자동 등록
├── servers/            ← MCP 서버 코드 (번들)
├── python/             ← Python COM 브릿지 (번들)
├── commands/           ← 8개 워크플로우 커맨드
├── hooks/              ← HWPX 규칙 + 에러 복구
└── agents/             ← 자동 트리거
```

## 주의: MCP와 중복 설치하지 마세요

이미 `claude-code-hwp-mcp` MCP를 settings.json에 등록한 상태에서 이 플러그인을 설치하면, 같은 도구가 2개씩 등록되어 토큰이 낭비됩니다.

**플러그인을 설치했다면 settings.json에서 hwp-studio MCP 설정을 삭제하세요.**

---

## 사전 요구사항

| 요구사항 | 설치 방법 |
|----------|-----------|
| Windows 10/11 | macOS/Linux 미지원 |
| 한글(HWP) 프로그램 | 한컴오피스 2014 이상 |
| Python 3.8+ | [python.org](https://www.python.org/downloads/) (Store 버전 X) |
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

`claude-code-hwp-plugin@hwp-marketplace` 가 `enabled` 상태이면 성공입니다.

### 로컬 설치 (개발자용)

```bash
git clone https://github.com/gmlcjf0326/claude-code-hwp-plugin.git
cd claude-code-hwp-plugin
npm install
claude plugins marketplace add .
claude plugins install claude-code-hwp-plugin@hwp-marketplace
```

## 시작하기

1. 한글(HWP) 프로그램을 실행합니다
2. `/hwp-setup`으로 환경을 확인합니다
3. `/hwp-help`로 사용 가능한 기능을 확인합니다

---

## 사용 방법

### 자연어로 요청 (가장 쉬움)

별도 커맨드 없이 자연어로 요청하면 agent가 자동으로 감지합니다:

```
"이 한글 문서 채워줘"
→ agent가 자동으로 /hwp-fill 워크플로우 실행

"사업계획서 작성해줘"
→ agent가 자동으로 /hwp-write 워크플로우 실행
```

### 커맨드로 요청 (명시적)

```
/hwp-fill     → 양식 자동 채우기
/hwp-write    → 문서 작성
/hwp-analyze  → 문서 분석
/hwp-convert  → 형식 변환
/hwp-batch    → 다건 생성
/hwp-privacy  → 개인정보 스캔
/hwp-setup    → 환경 진단
/hwp-help     → 기능 안내
```

### MCP 도구 직접 호출 (숙련자)

85개+ MCP 도구를 개별적으로 호출할 수도 있습니다:

```
"hwp_open_document로 C:\문서\양식.hwp 열어줘"
"hwp_find_replace로 '테스트'를 'TEST'로 바꿔줘"
"hwp_export_pdf로 PDF 변환해줘"
```

---

## 커맨드 상세

### /hwp-fill — 양식 자동 채우기

사업계획서, 과업지시서, 신청서 등의 빈칸을 AI가 자동으로 채웁니다.

**흐름**: 사전질문 → 문서열기 → 분석 → 서식파악 → 셀매핑 → 채우기 → 개인정보확인 → 저장

### /hwp-write — 문서 작성

AI가 한글 문서를 처음부터 작성합니다. 양식이 있으면 서식을 따릅니다.

**흐름**: 문서유형확인 → 양식열기 → 서식추출 → 작성요령삭제 → 내용작성 → 저장

### /hwp-analyze — 문서 분석

문서의 구조, 내용, 완성도를 분석하여 요약합니다.

**흐름**: 문서열기 → 심층분석 → 표구조 → 양식감지 → 통계 → 보고서

### /hwp-convert — 형식 변환

HWP 문서를 PDF, DOCX, HTML, HWPX로 변환합니다.

### /hwp-batch — 다건 문서 생성

엑셀/CSV 데이터를 기반으로 여러 건의 문서를 일괄 생성합니다.

### /hwp-privacy — 개인정보 스캔

문서에서 주민번호, 전화번호, 이메일 등 개인정보를 자동 감지합니다.

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
| RPC/COM 에러 | 한글 재시작 안내 |
| 파일 잠금 (EBUSY) | 한글에서 파일 닫기 안내 |
| Python 미설치 | python.org 설치 안내 |
| 문서 미열기 | hwp_open_document 안내 |
| HWP find_replace 실패 | HWPX 변환 권유 |

---

## HWP vs HWPX

| 구분 | HWP | HWPX |
|------|-----|------|
| 텍스트 검색 | COM (제한적) | XML 직접 (안정적) |
| 찾기/바꾸기 | COM (제한적) | XML 직접 (안정적) |
| 표 편집 | COM | COM |
| 문서 열기/저장 | COM | COM |

**HWPX 사용 권장.** 한글에서 "다른 이름으로 저장" > HWPX 형식 선택.

## 알려진 제한사항

- HWP 바이너리 파일에서 text_search가 0건 반환 가능 (COM API 한계)
- HWPX 파일이 한글에서 열린 상태에서 XML 편집 시 COM 폴백
- Windows 전용

---

## Claude Desktop을 사용하는 경우

이 플러그인은 Claude Code 전용입니다. Claude Desktop에서는 MCP 패키지를 사용하세요:

```bash
npm install -g claude-code-hwp-mcp
```

설정: [claude-code-hwp-mcp README](https://github.com/gmlcjf0326/claude-code-hwp-mcp)

---

## 라이선스

MIT License

## 관련 링크

- [MCP 패키지 (Claude Desktop용)](https://github.com/gmlcjf0326/claude-code-hwp-mcp)
- [pyhwpx](https://pypi.org/project/pyhwpx/)
- [MCP Protocol](https://modelcontextprotocol.io/)
