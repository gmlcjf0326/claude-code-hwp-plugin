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
