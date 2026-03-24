/**
 * MCP Prompts: fill_document, edit_document, analyze_document, batch_process
 */
import { z } from 'zod';
export function registerPrompts(server) {
    server.prompt('fill_document', 'HWP 문서의 빈 필드와 표 셀을 사전질문→분석→미리보기→채우기→검증 흐름으로 채웁니다.', {
        file_path: z.string().describe('HWP 파일 경로'),
        context: z.string().optional().describe('추가 지시사항 (선택)'),
    }, ({ file_path, context }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `다음 HWP 문서의 빈 항목을 채워주세요.

## ⚠️ 바로 채우지 마세요. 아래 순서를 따르세요.

### Step 1: 사전 확인
사용자에게 먼저 확인하세요:
- "참고할 자료(엑셀, 텍스트 등)가 있으신가요?"
- "분량은 간결하게 / 표준 / 상세하게 중 어느 수준으로?"
- "확정된 수치(금액, 일정 등)가 있으면 알려주세요"

### Step 2: 문서 분석
- hwp_smart_analyze로 구조와 빈 항목 파악
- 참고자료가 있으면 hwp_read_reference로 로드

### Step 3: 미리보기
- "다음과 같이 채울 예정입니다:" → 내용 요약 제시
- 사용자 확인 후 진행

### Step 4: 채우기 + 저장
- hwp_smart_fill 또는 hwp_fill_table_cells로 채우기
- hwp_save_document로 저장

### Step 5: 검증
- 빈 항목 잔여 확인 (hwp_smart_analyze 재실행)
- "수정할 부분이 있으면 말씀해주세요"

파일: ${file_path}
${context ? `추가 지시: ${context}` : ''}

규칙:
- 결재란/서명란: AI가 채우면 안 됨
- 이미 내용이 있는 셀: 변경하지 않음
- 사용자가 제공하지 않은 수치는 임의로 만들지 않음`,
                },
            }],
    }));
    server.prompt('edit_document', '현재 열린 HWP 문서의 특정 부분을 수정합니다.', {
        instructions: z.string().describe('수정 지시사항'),
    }, ({ instructions }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `현재 열린 HWP 문서를 다음 지시에 따라 수정해주세요.

작업 순서:
1. hwp_analyze_document 또는 hwp_get_document_text로 현재 내용 확인
2. 지시사항에 맞게 hwp_find_replace, hwp_fill_fields, hwp_fill_table_cells 등 적절한 도구 사용
3. 수정 후 hwp_save_document로 저장

지시사항: ${instructions}

규칙:
- 지시하지 않은 부분은 변경하지 않음
- 직급/부서명은 원본 그대로 유지 (변경 금지)
- 결재란은 비워둠 (AI가 채우면 안 됨)
- 날짜 형식: 마침표 구분 (예: 2026. 3. 19.)
- 금액: 한글+숫자 병기 (예: 금일백만원정(₩1,000,000))
- 수정 전후 내용을 요약해서 보고`,
                },
            }],
    }));
    server.prompt('analyze_document', '문서를 분석하고 구조, 내용, 완성도를 요약합니다.', {
        file_path: z.string().describe('HWP 파일 경로'),
    }, ({ file_path }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `다음 HWP 문서를 분석하고 요약해주세요.

작업 순서:
1. hwp_document_summary로 문서 완성도 확인
2. hwp_get_document_text로 본문 내용 확인
3. hwp_get_tables로 표 데이터 확인

파일: ${file_path}

다음 항목을 보고해주세요:
- 문서 종류 및 목적
- 페이지 수, 표 수, 필드 수
- 작성 완성도(%)
- 빈 항목이 있다면 목록
- 문서 내용 요약 (3-5줄)`,
                },
            }],
    }));
    server.prompt('batch_process', '디렉토리 내 모든 HWP 파일을 일괄 처리합니다.', {
        directory: z.string().describe('처리할 디렉토리 경로'),
        operation: z.enum(['analyze', 'fill', 'convert']).describe('작업 유형: analyze, fill, convert'),
    }, ({ directory, operation }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `디렉토리 내의 모든 HWP 파일을 일괄 처리합니다.

작업 순서:
1. hwp_list_files로 '${directory}' 내 파일 목록 확인
2. 각 파일에 대해 순차적으로 ${operation} 작업 수행
3. 작업 완료 후 결과 요약 보고

작업 유형: ${operation}
- analyze: 각 파일 분석 및 요약
- fill: 각 파일의 빈 항목 채우기
- convert: 각 파일을 지정 형식으로 변환

주의: 각 파일 처리 후 반드시 hwp_close_document로 닫은 후 다음 파일 진행`,
                },
            }],
    }));
    server.prompt('fill_public_document', '공공기관 사업계획서/신청서 작성. 기본: 빠른 모드(2개 질문). context에 "--상세" 추가 시 상세 모드(5개 질문+미리보기).', {
        file_path: z.string().describe('HWP 파일 경로'),
        reference: z.string().optional().describe('참고자료 파일 경로 또는 내용'),
        context: z.string().optional().describe('추가 지시. "--상세" 포함 시 상세 모드'),
    }, ({ file_path, reference, context }) => {
        const detailMode = context?.includes('--상세') || context?.includes('--detail');
        return {
            messages: [{
                    role: 'user',
                    content: {
                        type: 'text',
                        text: `한글 문서를 작성해주세요.

## ⚠️ 사전 확인 (반드시 먼저 물어보세요. 바로 채우지 마세요.)

다음 질문을 사용자에게 확인한 후 작업을 시작하세요:

1. **참고자료**: "채울 내용의 참고자료(엑셀, 기존 문서, 회사소개서 등)가 있으신가요?"
   → 있으면 hwp_read_reference로 로드
2. **양식**: "참고할 양식 파일이 있으신가요? (서식/구조를 따를 문서)"
   → 있으면 hwp_extract_style_profile로 서식(글꼴/크기/자간/들여쓰기) 추출하여 동일하게 적용
   → 없으면 일반 공공기관 표준 적용 (document_format_guide 참조)
3. **말투**: "개괄식(~했음, ~임)과 격식체(~했습니다, ~입니다) 중 어떤 문체로?"
4. **분량**: "예상 페이지 수가 있으신가요? (예: 5~6페이지)"
5. **확정 수치**: "금액, 일정, 목표 등 확정된 수치가 있으면 알려주세요" (없으면 임의 생성하지 않음)
6. **추가 참조**: "웹 검색이나 다른 소스를 참조할까요?"

## 작업 순서
1. hwp_smart_analyze → 문서 구조/타입/완성도 파악
2. 참고자료 있으면 → hwp_auto_fill_from_reference로 자동 매핑+채우기
3. 없으면 → hwp_smart_fill로 서식 보존하며 채우기
${detailMode ? `4. 채우기 전 미리보기 → 사용자 확인 후 진행` : ''}
5. hwp_save_document → 저장
6. 검증: hwp_privacy_scan(개인정보) + 빈 셀 확인 → 결과 보고

파일: ${file_path}
${reference ? `참고자료: ${reference}` : ''}
${context ? `지시: ${context.replace('--상세', '').replace('--detail', '').trim()}` : ''}

## 규칙 (공문서 표준 — document_format_guide 참조)
- 서식: style 미지정 → 기존 셀 서식 자동 상속
- 날짜: "2025. 3. 22.", 금액: "금10,269,000원", 순번: 1.→가.→1)→가)
- 결재란/서명란/이미 채워진 셀: 변경 금지

## 양식 동적 대응 원칙
- 양식이 제공되면 반드시 hwp_extract_style_profile로 서식 추출 → 동일 적용
- 글꼴/크기/자간/줄간격을 고정 규칙이 아닌 양식에서 추출한 값으로 적용
- 순번 체계는 양식의 패턴을 따름 (Ⅰ/□/1./ㅇ 등 양식마다 다름)
- 작성요령(< 작성요령 >, ※ 표시) → 작성 완료 후 hwp_delete_guide_text로 삭제 여부 확인
- 체크박스(□, ☐) → hwp_toggle_checkbox로 체크 전환
- 공문서 양식은 수신/참조/결재란 구조를 인식하고 보존`,
                    },
                }],
        };
    });
    server.prompt('document_format_guide', '한국 공공기관 공문서 서식 표준 가이드를 제공합니다.', {}, () => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `한국 공공기관 공문서 서식 표준을 알려주세요.

## 핵심 규격

### 1. 문서 구조: 두문·본문·결문 3단 구성
- 두문: 행정기관명, 수신자
- 본문: 제목, 내용, 붙임 (붙임 뒤 쌍점 없음)
- 결문: 발신명의, 기안자/검토자/결재자, 기관 정보

### 2. 글꼴/크기
| 용도 | 보고서 | 기안문 |
|------|--------|--------|
| 본문 | 휴먼명조(HTF) 15pt | 맑은 고딕 11.5pt |
| 대제목 | HY헤드라인M 22pt | - |
| 소제목 | 16pt | - |
| 참조 | 중고딕 13pt | - |
| 표 안 | 10pt | 10pt |

### 3. 줄간격/자간/장평
| 항목 | 보고서 | 기안문(맑은고딕) | 기안문(굴림, 구) |
|------|--------|----------------|----------------|
| 줄간격 | 160% | 103% | 123% |
| 장평 | 100 | 100 | 100 |
| 자간 | 0 (-10~0) | 0 | 0 |
| 표 안 줄간격 | 100~130% | - | 123% |

### 4. 순번 체계 (시행규칙 제2조 제1항)
1. → 가. → 1) → 가) → (1) → (가) → ① → ㉮
들여쓰기: 2단계부터 2타(한글 1자)씩

### 5. 여백 (보고서/공문서)
| 항목 | 보고서(mm) | 공문서(mm) |
|------|-----------|-----------|
| 위 | 15 | 30 |
| 아래 | 10~15 | 15 |
| 좌 | 20 | 20 |
| 우 | 20 | 15 |

### 6. 날짜/금액/기타
- 날짜: "2025. 3. 21." (연월일 글자 생략, 마침표 사이 1타)
- 금액: 아라비아 숫자 + 괄호 한글 병기
- "끝" 표시: 본문 마지막 2타 띄운 후
- 법령 인용: 낫표 「 」
- 관인: 발신명의 마지막 글자가 인영 가운데`,
                },
            }],
    }));
    server.prompt('review_document', '작성 완료된 문서를 검토하고 수정점을 제안합니다. 공문서 표준 준수, 개인정보, 빈 항목, 내용 일관성을 점검합니다.', {
        file_path: z.string().describe('HWP 파일 경로'),
    }, ({ file_path }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `다음 문서를 검토하고 수정이 필요한 부분을 알려주세요.

## 검토 순서
1. hwp_smart_analyze로 문서 구조 + 완성도 확인
2. hwp_get_as_markdown으로 전체 내용 확인
3. hwp_privacy_scan으로 개인정보 포함 여부 확인

## 검토 항목 (체크리스트)
- [ ] 빈 셀/필드가 남아있는지
- [ ] 날짜 형식: "2025. 3. 22." (마침표 구분, 연월일 글자 생략)
- [ ] 금액 형식: "금OOO원" (아라비아 숫자)
- [ ] 순번 체계: 1. → 가. → 1) → 가) → (1) → (가) → ① → ㉮
- [ ] 결재란/서명란이 비워져 있는지 (AI가 채우면 안 됨)
- [ ] 개인정보(주민번호, 전화번호, 이메일) 포함 시 마스킹 필요 여부
- [ ] 내용의 논리적 일관성 (목표 ↔ 추진전략 ↔ 기대효과 연결)
- [ ] 문체 일관성 (격식체/일반체 혼용 여부)

파일: ${file_path}

## 결과 형식
검토 결과를 다음 형식의 표로 정리하세요:
| 항목 | 상태 | 내용 |
|------|------|------|
| 완성도 | ✅/⚠️ | 예: 100%, 빈 셀 3개 |
| 날짜 형식 | ✅/❌ | 예: 모두 정상 / "2025년 3월" → "2025. 3." 수정 필요 |
| ... | | |

수정이 필요한 부분은 구체적인 위치(표 번호, 셀)와 수정 방법을 안내하세요.`,
                },
            }],
    }));
    server.prompt('auto_fill_from_excel', '엑셀/CSV 데이터로 HWP 양식을 자동으로 채웁니다. "엑셀로 양식 채워줘" 요청에 사용.', {
        file_path: z.string().describe('HWP 파일 경로'),
        excel_path: z.string().describe('엑셀 또는 CSV 파일 경로'),
    }, ({ file_path, excel_path }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `엑셀 데이터로 한글 양식을 자동 채워주세요.

## 작업 순서
1. hwp_smart_analyze로 HWP 문서 구조 파악
2. hwp_read_reference로 엑셀 데이터 로드
3. hwp_auto_fill_from_reference로 자동 매핑 + 채우기
4. 매핑 안 된 항목은 사용자에게 확인
5. hwp_save_document로 저장
6. hwp_privacy_scan으로 개인정보 확인

HWP 파일: ${file_path}
엑셀 파일: ${excel_path}

규칙:
- 서식: 기존 셀 서식 자동 상속
- 결재란/서명란: 채우지 않음
- 매핑 불확실한 항목: 사용자에게 확인 후 진행`,
                },
            }],
    }));
    server.prompt('create_report', '빈 문서에서 보고서를 새로 작성합니다.', {
        topic: z.string().describe('보고서 주제'),
        context: z.string().optional().describe('추가 지시'),
    }, ({ topic, context }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `새 보고서를 작성해주세요.

## 사전 확인
1. "참고자료가 있으신가요?"
2. "보고서 분량은? (간결/표준/상세)"
3. "특별히 포함할 내용이 있으면 알려주세요"

## 작업 순서
1. 현재 열린 문서에 hwp_insert_heading으로 제목 삽입
2. hwp_insert_markdown으로 본문 구조 작성
3. 필요시 hwp_table_create_from_data로 표 삽입
4. hwp_save_document로 저장

주제: ${topic}
${context ? `추가 지시: ${context}` : ''}

규칙:
- 공문서 표준: 순번 1.→가.→1)→가), 날짜 "2025. 3. 22."
- 서식: 제목 Bold 22pt, 본문 11pt`,
                },
            }],
    }));
    server.prompt('write_document', '양식을 참고하여 새 문서를 처음부터 작성합니다. 양식 서식 추출 → 사전 질문 → 내용 작성 → 저장.', {
        topic: z.string().describe('문서 주제/과업명'),
        template_path: z.string().optional().describe('참고 양식 파일 경로 (서식/구조를 따를 문서)'),
        reference_path: z.string().optional().describe('참고 내용 파일 경로 (내용을 가져올 문서)'),
    }, ({ topic, template_path, reference_path }) => ({
        messages: [{
                role: 'user',
                content: {
                    type: 'text',
                    text: `다음 주제로 문서를 작성해주세요.

## ⚠️ 사전 확인 (반드시 먼저 물어보세요)

1. **말투**: "개괄식(~했음, ~임)과 격식체(~했습니다, ~입니다) 중 어떤 문체로?"
2. **분량**: "예상 페이지 수가 있으신가요?"
3. **확정 수치**: "금액, 일정 등 확정된 수치가 있으면 알려주세요"
4. **추가 참조**: "웹 검색이나 다른 소스를 참조할까요?"
5. **구분 표시**: "AI가 추가한 내용을 파란색으로 표시할까요?"

## 작업 순서
${template_path ? `1. 양식 분석: hwp_open_document("${template_path}") → hwp_extract_style_profile로 서식 추출` : '1. 일반 공공기관 표준 서식 적용'}
${reference_path ? `2. 내용 참고: hwp_read_reference("${reference_path}")로 내용 로드` : '2. AI가 주제에 맞는 내용 생성'}
3. 새 문서에 양식 서식 + 내용 조합하여 작성
4. 들여쓰기 적용 (left_margin + indent로 첫줄/나머지줄 설정)
5. hwp_save_document로 저장 (한글에서 열린 상태 유지)
6. 검증: hwp_privacy_scan + 빈 항목 확인

주제: ${topic}
${template_path ? `양식: ${template_path}` : ''}
${reference_path ? `참고 내용: ${reference_path}` : ''}

## 서식 규칙
- 양식이 제공되면: hwp_extract_style_profile 결과에 따라 동적 적용
- 양식이 없으면: 공공기관 표준 (document_format_guide 참조)
- 들여쓰기: left_margin(나머지 줄) + indent(첫 줄 추가) 조합
- 색상 구분: 원본 내용=검정, AI 추가=파란색(0,0,255)
- 작업 완료 후 한글 프로그램과 문서를 열린 상태로 유지`,
                },
            }],
    }));
}
