/**
 * HWPX XML Engine — 한글 프로그램 없이 HWPX 파일 직접 생성/편집
 *
 * HWPX = ZIP(application/hwp+zip) + XML 형식
 * Python execSync 제거 — Node.js jszip + @xmldom/xmldom만 사용
 *
 * CLAUDE.md 규칙 준수:
 * - @xmldom/xmldom 사용 (fast-xml-parser 금지)
 * - element.localName 사용 (tagName 금지)
 * - 텍스트 수정 후 linesegarray 삭제 필수
 * - charPrIDRef 변경 금지
 * - 표 셀 경로: tc → subList → p → run → t
 */
import fs from 'node:fs';
import path from 'node:path';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
export const TEMPLATES = [
    // 공문서
    { id: 'gov_official_letter', name: '공문서 (기안문/시행문)', category: '공문서', fields: ['수신자', '발신부서', '발신자', '제목', '본문', '시행일자', '문서번호'] },
    { id: 'gov_report', name: '보고서', category: '공문서', fields: ['보고제목', '보고자', '부서', '보고일', '요약', '현황', '문제점', '개선방안', '기대효과'] },
    { id: 'gov_draft', name: '기안문', category: '공문서', fields: ['기안자', '검토자', '결재자', '제목', '내용', '시행일', '근거'] },
    { id: 'gov_minutes', name: '회의록', category: '공문서', fields: ['회의명', '일시', '장소', '참석자', '안건', '토의내용', '결정사항', '향후계획'] },
    { id: 'gov_plan', name: '사업계획서', category: '공문서', fields: ['사업명', '기관명', '사업기간', '사업목적', '추진내용', '기대효과', '예산'] },
    { id: 'gov_notice', name: '공고문', category: '공문서', fields: ['공고제목', '공고내용', '공고일', '기관명'] },
    { id: 'gov_budget', name: '예산서', category: '공문서', fields: ['사업명', '항목', '금액', '산출근거'] },
    // 기업
    { id: 'biz_proposal', name: '사업제안서', category: '기업', fields: ['회사명', '제안제목', '제안배경', '제안내용', '기대효과', '일정', '예산'] },
    { id: 'biz_contract', name: '계약서', category: '기업', fields: ['갑', '을', '계약명', '계약금액', '계약기간'] },
    { id: 'biz_invoice', name: '견적서', category: '기업', fields: ['발행처', '수신처', '품목', '합계', '부가세', '총액'] },
    { id: 'biz_meeting', name: '기업 회의록', category: '기업', fields: ['회의명', '일시', '참석자', '안건', '결정사항'] },
    { id: 'biz_memo', name: '업무 메모', category: '기업', fields: ['수신', '발신', '제목', '내용'] },
    { id: 'biz_mou', name: '양해각서(MOU)', category: '기업', fields: ['기관1', '기관2', '목적', '협력내용', '기간'] },
    { id: 'biz_nda', name: '비밀유지계약서(NDA)', category: '기업', fields: ['갑', '을', '비밀정보범위', '기간'] },
    // 학술
    { id: 'academic_paper', name: '학술 논문', category: '학술', fields: ['제목', '저자', '초록', '키워드', '서론', '본론', '결론', '참고문헌'] },
    { id: 'academic_report', name: '학술 보고서', category: '학술', fields: ['제목', '작성자', '과목', '내용'] },
    // 개인
    { id: 'personal_resume', name: '이력서', category: '개인', fields: ['이름', '생년월일', '연락처', '이메일', '학력', '경력', '자격증'] },
    { id: 'personal_letter', name: '자기소개서', category: '개인', fields: ['이름', '지원분야', '성장배경', '지원동기', '입사후포부'] },
    { id: 'personal_certificate', name: '증명서', category: '개인', fields: ['성명', '생년월일', '발급사유', '발급일'] },
    // ── v0.6.0 추가 13개 (총 35종) ──
    { id: 'gov_appointment', name: '인사발령', category: '공문서', fields: ['대상자', '현직급', '현부서', '발령직급', '발령부서', '발령사유', '발령일'] },
    { id: 'gov_audit', name: '감사보고서', category: '공문서', fields: ['감사기간', '감사대상', '감사결과', '지적사항', '개선권고', '기관명'] },
    { id: 'legal_power_of_attorney', name: '위임장', category: '법무', fields: ['위임인', '위임인주소', '수임인', '수임인주소', '위임사항', '위임기간', '작성일'] },
    { id: 'legal_content_cert', name: '내용증명', category: '법무', fields: ['발신인', '발신주소', '수신인', '수신주소', '내용', '요청사항', '발송일'] },
    { id: 'hr_leave_request', name: '휴가신청서', category: 'HR', fields: ['신청인', '부서', '직급', '휴가종류', '시작일', '종료일', '사유'] },
    { id: 'hr_employment_contract', name: '고용계약서', category: 'HR', fields: ['갑', '을', '직위', '근무장소', '급여', '근무시간', '계약기간'] },
    { id: 'hr_resignation', name: '퇴사서', category: 'HR', fields: ['성명', '부서', '직급', '입사일', '퇴사일', '퇴사사유'] },
    { id: 'hr_performance_review', name: '성과평가서', category: 'HR', fields: ['평가대상', '평가기간', '업무성과', '역량평가', '종합등급', '평가자'] },
    { id: 'biz_purchase_order', name: '구매발주서', category: '기업', fields: ['발주처', '납품처', '품목', '수량', '단가', '금액', '납기일'] },
    { id: 'biz_expense_claim', name: '경비청구서', category: '기업', fields: ['청구인', '부서', '항목', '금액', '사용일', '증빙', '승인자'] },
    { id: 'biz_work_report', name: '업무보고서', category: '기업', fields: ['보고자', '부서', '보고기간', '주요실적', '이슈', '다음계획'] },
    { id: 'biz_training_plan', name: '교육훈련계획', category: '기업', fields: ['교육명', '교육대상', '교육기간', '교육장소', '교육내용', '예산'] },
    { id: 'finance_payslip', name: '급여명세서', category: '재무', fields: ['성명', '부서', '직급', '기본급', '수당', '공제', '실수령액', '지급일'] },
];
// ── HWPX 네임스페이스 ──
const NS_HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph';
// ── HWPX ZIP 유틸 (Node.js jszip — Python execSync 제거) ──
/**
 * HWPX(ZIP) 파일에서 특정 XML 파일을 읽어서 DOM으로 파싱.
 */
export async function readHwpxXml(hwpxPath, xmlName) {
    const data = fs.readFileSync(hwpxPath);
    const zip = await JSZip.loadAsync(data);
    const xmlFile = zip.file(xmlName);
    if (!xmlFile) {
        throw new Error(`HWPX 내에 ${xmlName}을 찾을 수 없습니다.`);
    }
    const xmlStr = await xmlFile.async('string');
    const parser = new DOMParser();
    return parser.parseFromString(xmlStr, 'text/xml');
}
/**
 * HWPX(ZIP) 파일의 특정 XML을 수정 후 저장.
 * 기존 ZIP의 다른 파일은 그대로 유지.
 */
export async function writeHwpxXml(sourcePath, outputPath, xmlName, doc) {
    const serializer = new XMLSerializer();
    const xmlStr = serializer.serializeToString(doc);
    const data = fs.readFileSync(sourcePath);
    const zip = await JSZip.loadAsync(data);
    zip.file(xmlName, xmlStr);
    const newData = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(outputPath, newData);
}
// ── 텍스트 추출 ──
/**
 * HWPX section XML에서 모든 텍스트를 추출.
 * 경로: hp:p > hp:run > hp:t
 */
export function extractTextFromSection(doc) {
    const texts = [];
    const paragraphs = doc.getElementsByTagNameNS(NS_HP, 'p');
    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const runs = p.getElementsByTagNameNS(NS_HP, 'run');
        let paraText = '';
        for (let j = 0; j < runs.length; j++) {
            const tNodes = runs[j].getElementsByTagNameNS(NS_HP, 't');
            for (let k = 0; k < tNodes.length; k++) {
                paraText += tNodes[k].textContent || '';
            }
        }
        texts.push(paraText);
    }
    return texts;
}
// ── 텍스트 치환 ──
/**
 * HWPX section XML에서 텍스트 찾아 바꾸기.
 * CLAUDE.md 규칙: 수정 후 linesegarray 삭제 필수.
 */
export function replaceTextInSection(doc, find, replace) {
    let count = 0;
    const tNodes = doc.getElementsByTagNameNS(NS_HP, 't');
    for (let i = 0; i < tNodes.length; i++) {
        const t = tNodes[i];
        const text = t.textContent || '';
        if (text.includes(find)) {
            t.textContent = text.replaceAll(find, replace);
            count++;
        }
    }
    // CLAUDE.md 규칙 8: 텍스트 수정 후 linesegarray 삭제 필수
    if (count > 0) {
        removeLinesegarray(doc);
    }
    return count;
}
// ── HWPX XML 검색 (COM 우회) ──
/**
 * HWPX section XML에서 텍스트 검색. COM FindReplace 우회용.
 */
export function searchTextInSection(doc, searchText) {
    const results = [];
    const paragraphs = doc.getElementsByTagNameNS(NS_HP, 'p');
    let matchCount = 0;
    for (let i = 0; i < paragraphs.length; i++) {
        const runs = paragraphs[i].getElementsByTagNameNS(NS_HP, 'run');
        let paraText = '';
        for (let j = 0; j < runs.length; j++) {
            const tNodes = runs[j].getElementsByTagNameNS(NS_HP, 't');
            for (let k = 0; k < tNodes.length; k++) {
                paraText += tNodes[k].textContent || '';
            }
        }
        let pos = 0;
        while ((pos = paraText.indexOf(searchText, pos)) !== -1) {
            matchCount++;
            const start = Math.max(0, pos - 20);
            const end = Math.min(paraText.length, pos + searchText.length + 20);
            results.push({
                index: matchCount,
                paragraph: i,
                context: paraText.slice(start, end),
            });
            pos += searchText.length;
        }
    }
    return { total: matchCount, results };
}
/**
 * HWPX section XML에서 N번째 텍스트만 치환.
 */
export function replaceTextNthInSection(doc, find, replace, nth) {
    const tNodes = doc.getElementsByTagNameNS(NS_HP, 't');
    let matchCount = 0;
    for (let i = 0; i < tNodes.length; i++) {
        const t = tNodes[i];
        const text = t.textContent || '';
        let pos = 0;
        while ((pos = text.indexOf(find, pos)) !== -1) {
            matchCount++;
            if (matchCount === nth) {
                t.textContent = text.slice(0, pos) + replace + text.slice(pos + find.length);
                removeLinesegarray(doc);
                return true;
            }
            pos += find.length;
        }
    }
    return false;
}
/**
 * HWPX section XML에서 텍스트를 찾아 그 뒤에 추가.
 */
export function findAndAppendInSection(doc, find, appendText) {
    // 문단(p) 단위로 run/t 텍스트를 연결하여 검색 (searchTextInSection과 동일 패턴)
    const paragraphs = doc.getElementsByTagNameNS(NS_HP, 'p');
    for (let i = 0; i < paragraphs.length; i++) {
        const runs = paragraphs[i].getElementsByTagNameNS(NS_HP, 'run');
        let paraText = '';
        const tNodeList = [];
        for (let j = 0; j < runs.length; j++) {
            const tNodes = runs[j].getElementsByTagNameNS(NS_HP, 't');
            for (let k = 0; k < tNodes.length; k++) {
                const nodeText = tNodes[k].textContent || '';
                tNodeList.push({ node: tNodes[k], offset: paraText.length });
                paraText += nodeText;
            }
        }
        const pos = paraText.indexOf(find);
        if (pos !== -1) {
            const endPos = pos + find.length;
            for (let n = tNodeList.length - 1; n >= 0; n--) {
                const entry = tNodeList[n];
                const nodeText = entry.node.textContent || '';
                if (entry.offset <= endPos && endPos <= entry.offset + nodeText.length) {
                    const localOffset = endPos - entry.offset;
                    entry.node.textContent = nodeText.slice(0, localOffset) + appendText + nodeText.slice(localOffset);
                    removeLinesegarray(doc);
                    return true;
                }
            }
            // fallback: 마지막 t 노드에 추가
            if (tNodeList.length > 0) {
                const last = tNodeList[tNodeList.length - 1];
                last.node.textContent = (last.node.textContent || '') + appendText;
                removeLinesegarray(doc);
                return true;
            }
        }
    }
    return false;
}
/**
 * linesegarray 요소 삭제 (CLAUDE.md 규칙 8).
 */
function removeLinesegarray(doc) {
    const linesegArrays = doc.getElementsByTagNameNS(NS_HP, 'linesegarray');
    const toRemove = [];
    for (let i = 0; i < linesegArrays.length; i++) {
        toRemove.push(linesegArrays[i]);
    }
    for (const el of toRemove) {
        el.parentNode?.removeChild(el);
    }
}
/**
 * 특정 element 내부의 linesegarray만 삭제 (전역 X).
 * v0.7.0: 표 셀 단위 텍스트 변경 시 사용.
 */
function removeLinesegarrayInElement(el) {
    const linesegArrays = el.getElementsByTagNameNS(NS_HP, 'linesegarray');
    const toRemove = [];
    for (let i = 0; i < linesegArrays.length; i++) {
        toRemove.push(linesegArrays[i]);
    }
    for (const tgt of toRemove) {
        tgt.parentNode?.removeChild(tgt);
    }
}
/**
 * v0.7.0 신규: 표의 (row, col) 셀에서 텍스트 치환.
 * 경로: tc → subList → p → run → t (CLAUDE.md 규칙 10).
 * - linesegarray 셀 내부만 삭제
 * - charPrIDRef 보존 (CLAUDE.md 규칙 9)
 * - 평탄화 tableIndex (재귀 X — 중첩 표는 replaceInNestedTable 사용)
 *
 * @throws Error 'TableNotFound' / 'RowOutOfRange' / 'ColOutOfRange'
 */
export function replaceInTableCell(doc, opts) {
    const warnings = [];
    const occurrence = opts.occurrence ?? 0;
    // 1. doc 전체에서 tbl flat list (재귀 X — direct children만 골라야 하지만
    //    HWPX는 모든 tbl이 sec/p/run 내부에 있으므로 getElementsByTagNameNS로 평탄화)
    const tbls = doc.getElementsByTagNameNS(NS_HP, 'tbl');
    if (opts.tableIndex < 0 || opts.tableIndex >= tbls.length) {
        throw new Error(`TableNotFound: index=${opts.tableIndex}, total=${tbls.length}`);
    }
    const tbl = tbls[opts.tableIndex];
    // 2. tbl의 direct tr children
    const allTrs = tbl.getElementsByTagNameNS(NS_HP, 'tr');
    // direct children만 (parent === tbl)
    const trs = [];
    for (let i = 0; i < allTrs.length; i++) {
        const tr = allTrs[i];
        if (tr.parentNode === tbl)
            trs.push(tr);
    }
    if (opts.rowIndex < 0 || opts.rowIndex >= trs.length) {
        throw new Error(`RowOutOfRange: row=${opts.rowIndex}, total=${trs.length}`);
    }
    const tr = trs[opts.rowIndex];
    // 3. tr의 direct tc children
    const allTcs = tr.getElementsByTagNameNS(NS_HP, 'tc');
    const tcs = [];
    for (let i = 0; i < allTcs.length; i++) {
        const tc = allTcs[i];
        if (tc.parentNode === tr)
            tcs.push(tc);
    }
    if (opts.colIndex < 0 || opts.colIndex >= tcs.length) {
        throw new Error(`ColOutOfRange: col=${opts.colIndex}, total=${tcs.length}`);
    }
    const tc = tcs[opts.colIndex];
    // 4. tc → 첫 subList → p → run → t 노드 수집
    const allSubLists = tc.getElementsByTagNameNS(NS_HP, 'subList');
    let subList = null;
    for (let i = 0; i < allSubLists.length; i++) {
        const sl = allSubLists[i];
        if (sl.parentNode === tc) {
            subList = sl;
            break;
        }
    }
    if (!subList) {
        return { matched: 0, cellText: '', charPrIDRef: null, warnings: ['NoSubList'] };
    }
    // 5. cell의 모든 t 노드 수집 + paraText 재구성
    const tNodes = subList.getElementsByTagNameNS(NS_HP, 't');
    const tElements = [];
    let cellText = '';
    let charPrIDRef = null;
    for (let i = 0; i < tNodes.length; i++) {
        const tEl = tNodes[i];
        tElements.push(tEl);
        cellText += tEl.textContent || '';
        // 첫 run의 charPrIDRef 추출
        if (charPrIDRef === null) {
            const runEl = tEl.parentNode;
            if (runEl) {
                const ref = runEl.getAttribute('charPrIDRef');
                if (ref)
                    charPrIDRef = ref;
            }
        }
    }
    // 6. find 위치 탐색 + 치환
    if (!opts.find || cellText.indexOf(opts.find) === -1) {
        return { matched: 0, cellText, charPrIDRef, warnings };
    }
    let matched = 0;
    // 단순 전략: 첫 t 노드에 치환된 텍스트 통합, 나머지는 빈 문자열로.
    // (run 경계 가로지르는 매치도 안전 처리. charPrIDRef는 첫 run 보존)
    let newCellText;
    if (occurrence === 0) {
        // 전체 치환
        const before = cellText;
        newCellText = before.split(opts.find).join(opts.replace);
        matched = (before.length - newCellText.replace(new RegExp(escapeRegex(opts.replace), 'g'), '').length) /
            Math.max(opts.find.length, 1);
        // 더 정확한 카운트: split 결과 - 1
        matched = before.split(opts.find).length - 1;
    }
    else {
        // N번째 치환
        let idx = -1;
        for (let n = 0; n < occurrence; n++) {
            idx = cellText.indexOf(opts.find, idx + 1);
            if (idx === -1)
                break;
        }
        if (idx === -1) {
            return { matched: 0, cellText, charPrIDRef, warnings: ['OccurrenceNotFound'] };
        }
        newCellText =
            cellText.substring(0, idx) +
                opts.replace +
                cellText.substring(idx + opts.find.length);
        matched = 1;
    }
    // 7. 첫 t에 newCellText 통합, 나머지 t는 빈 문자열
    if (tElements.length > 0) {
        tElements[0].textContent = newCellText;
        for (let i = 1; i < tElements.length; i++) {
            tElements[i].textContent = '';
        }
    }
    // 8. 셀 내부 linesegarray만 삭제
    removeLinesegarrayInElement(tc);
    return { matched, cellText: newCellText, charPrIDRef, warnings };
}
function escapeRegex(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
/**
 * v0.7.0 신규: 문서 내 모든 자동 계산 필드에 dirty="1" attribute 추가.
 * 한글이 파일을 다음 열 때 자동 재계산.
 *
 * 처리 대상: PageNum, TotalPage, Date, Time, TOC, Index, CrossRef, FieldFormula
 */
export function markFieldsForRecalc(doc, fieldTypes) {
    const types = fieldTypes && fieldTypes.length > 0 ? fieldTypes : ['all'];
    const wantAll = types.includes('all');
    const result = { marked: 0, byType: {}, unsupported: [] };
    // 1. 모든 ctrl 요소 수집
    const ctrls = doc.getElementsByTagNameNS(NS_HP, 'ctrl');
    // 2. 각 ctrl의 자식 첫 element의 localName으로 종류 식별
    const typeMap = {
        pageNum: 'PageNum',
        pageNumCtrl: 'PageNum',
        pageInfo: 'PageNum',
        totalPage: 'TotalPage',
        totalPageCtrl: 'TotalPage',
        date: 'Date',
        dateCtrl: 'Date',
        time: 'Time',
        timeCtrl: 'Time',
        tocCtrl: 'TOC',
        toc: 'TOC',
        indexCtrl: 'Index',
        index: 'Index',
        crossRefCtrl: 'CrossRef',
        crossRef: 'CrossRef',
        fieldBegin: 'FieldFormula', // generic field
    };
    for (let i = 0; i < ctrls.length; i++) {
        const ctrl = ctrls[i];
        // 자식 element 중 첫 번째
        let firstChildEl = null;
        const children = ctrl.childNodes;
        for (let j = 0; j < children.length; j++) {
            const child = children[j];
            if (child && child.nodeType === 1 /* ELEMENT_NODE */) {
                firstChildEl = child;
                break;
            }
        }
        if (!firstChildEl) {
            result.unsupported.push('ctrl-no-child');
            continue;
        }
        const localName = firstChildEl.localName || '';
        const fieldType = typeMap[localName];
        if (!fieldType) {
            result.unsupported.push(localName);
            continue;
        }
        if (!wantAll && !types.includes(fieldType)) {
            continue;
        }
        // dirty attribute 설정
        try {
            firstChildEl.setAttribute('dirty', '1');
            ctrl.setAttribute('dirty', '1');
            result.marked++;
            result.byType[fieldType] = (result.byType[fieldType] || 0) + 1;
        }
        catch (e) {
            result.unsupported.push(`${localName}-setattr-failed`);
        }
    }
    return result;
}
/**
 * v0.7.0 신규: HWPX XML과 COM 통계 비교.
 * Python COM에서 사전 수집한 stats를 입력으로 받아 XML 파싱 결과와 diff.
 * 임계: |Δtables| ≤ 0, |Δparagraphs| ≤ 1, |Δchars| ≤ 5
 */
export async function compareCOMAndXML(filePath, comStats) {
    const warnings = [];
    // section0.xml만 우선 처리 (다중 섹션은 향후 확장)
    let doc;
    try {
        doc = await readHwpxXml(filePath, 'Contents/section0.xml');
    }
    catch (e) {
        return {
            com: comStats,
            xml: { tables: 0, paragraphs: 0, fields: 0, runs: 0, chars: 0 },
            diff: { tables: 0, paragraphs: 0, fields: 0, runs: 0, chars: 0 },
            ok: false,
            warnings: [`readHwpxXml failed: ${e.message}`],
        };
    }
    const tables = doc.getElementsByTagNameNS(NS_HP, 'tbl').length;
    const paragraphs = doc.getElementsByTagNameNS(NS_HP, 'p').length;
    const ctrls = doc.getElementsByTagNameNS(NS_HP, 'ctrl').length;
    const runs = doc.getElementsByTagNameNS(NS_HP, 'run').length;
    let chars = 0;
    const tNodes = doc.getElementsByTagNameNS(NS_HP, 't');
    for (let i = 0; i < tNodes.length; i++) {
        chars += (tNodes[i].textContent || '').length;
    }
    const xml = { tables, paragraphs, fields: ctrls, runs, chars };
    const diff = {
        tables: xml.tables - comStats.tables,
        paragraphs: xml.paragraphs - comStats.paragraphs,
        fields: xml.fields - comStats.fields,
        runs: xml.runs - comStats.runs,
        chars: xml.chars - comStats.chars,
    };
    let ok = true;
    if (Math.abs(diff.tables) > 0) {
        warnings.push(`table count mismatch: Δ=${diff.tables}`);
        ok = false;
    }
    if (Math.abs(diff.paragraphs) > 1) {
        warnings.push(`paragraph count mismatch: Δ=${diff.paragraphs} (threshold ±1)`);
        ok = false;
    }
    if (Math.abs(diff.chars) > 5) {
        warnings.push(`char count mismatch: Δ=${diff.chars} (threshold ±5)`);
        ok = false;
    }
    return { com: comStats, xml, diff, ok, warnings };
}
/**
 * v0.7.0 신규 (인터페이스만): 중첩 표 셀 텍스트 치환.
 * - path.length === 1: replaceInTableCell로 위임 (정식 지원)
 * - path.length >= 2: 'nested-table-experimental' warning + 첫 단계만 처리
 *
 * 정식 다단계 지원은 v0.7.2.1에서 구현 (재귀 처리).
 */
export function replaceInNestedTable(doc, path, find, replace) {
    if (path.length === 0) {
        throw new Error('NestedPathError: empty path');
    }
    if (path.length === 1) {
        const step = path[0];
        return replaceInTableCell(doc, {
            tableIndex: step.tableIndex,
            rowIndex: step.row,
            colIndex: step.col,
            find,
            replace,
        });
    }
    // path.length >= 2: 첫 단계만 처리, 경고 추가
    const step = path[0];
    const result = replaceInTableCell(doc, {
        tableIndex: step.tableIndex,
        rowIndex: step.row,
        colIndex: step.col,
        find,
        replace,
    });
    result.warnings.push(`nested-table-experimental: depth=${path.length}, only first step processed (full support in v0.7.2.1)`);
    return result;
}
// ── 빈 HWPX 생성 ──
/**
 * BUG-9 fix: blank_template.hwpx 파일 의존 제거.
 * 최소 유효 HWPX를 프로그래밍적으로 생성.
 */
async function createMinimalHwpx(outputPath, title) {
    const zip = new JSZip();
    // mimetype
    zip.file('mimetype', 'application/hwp+zip');
    // META-INF/manifest.xml
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/hwp+zip"/>
  <manifest:file-entry manifest:full-path="Contents/section0.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Contents/content.hpf" manifest:media-type="text/xml"/>
</manifest:manifest>`);
    // Contents/content.hpf
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hp:HWPMLPackageFormat xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:BodyText>
    <hp:SectionRef hp:IDRef="0"/>
  </hp:BodyText>
</hp:HWPMLPackageFormat>`);
    // Contents/section0.xml
    const titleText = title || '';
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p>
    <hp:run>
      <hp:t>${titleText}</hp:t>
    </hp:run>
  </hp:p>
</hp:sec>`);
    const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(outputPath, buffer);
}
function getTemplatePath() {
    const thisFile = new URL(import.meta.url).pathname.replace(/^\/([A-Z]:)/, '$1');
    return path.join(path.dirname(thisFile), '../../blank_template.hwpx');
}
/**
 * 빈 HWPX 파일 생성.
 * BUG-9 fix: blank_template.hwpx가 없으면 프로그래밍적으로 생성.
 */
export async function createBlankHwpx(outputPath, title) {
    const templatePath = getTemplatePath();
    if (!fs.existsSync(templatePath)) {
        // 템플릿 파일 없음 → 프로그래밍적 생성
        await createMinimalHwpx(outputPath, title);
        return;
    }
    fs.copyFileSync(templatePath, outputPath);
    if (title) {
        const doc = await readHwpxXml(outputPath, 'Contents/section0.xml');
        const tNodes = doc.getElementsByTagNameNS(NS_HP, 't');
        if (tNodes.length > 0) {
            tNodes[0].textContent = title;
            removeLinesegarray(doc);
        }
        await writeHwpxXml(outputPath, outputPath, 'Contents/section0.xml', doc);
    }
}
// ── 템플릿 생성 ──
/**
 * 템플릿 기반 문서 생성.
 * blank_template.hwpx를 복사 → 변수 치환.
 */
export async function generateFromTemplate(templateId, variables, outputPath) {
    const template = TEMPLATES.find(t => t.id === templateId);
    if (!template) {
        throw new Error(`템플릿을 찾을 수 없습니다: ${templateId}. hwp_template_list로 사용 가능한 템플릿을 확인하세요.`);
    }
    // 빈 HWPX 복사
    await createBlankHwpx(outputPath);
    // 변수로 텍스트 생성
    const lines = [];
    lines.push(template.name);
    lines.push('');
    for (const field of template.fields) {
        const value = variables[field] || `{{${field}}}`;
        lines.push(`${field}: ${value}`);
    }
    // section0.xml에 텍스트 삽입
    const doc = await readHwpxXml(outputPath, 'Contents/section0.xml');
    const tNodes = doc.getElementsByTagNameNS(NS_HP, 't');
    if (tNodes.length > 0) {
        tNodes[0].textContent = lines.join('\n');
        removeLinesegarray(doc);
    }
    await writeHwpxXml(outputPath, outputPath, 'Contents/section0.xml', doc);
    const filledFields = template.fields.filter(f => variables[f]).length;
    const emptyFields = template.fields.filter(f => !variables[f]);
    return { filledFields, emptyFields };
}
