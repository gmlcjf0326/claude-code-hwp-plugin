export interface HwpxTemplate {
    id: string;
    name: string;
    category: string;
    fields: string[];
}
export declare const TEMPLATES: HwpxTemplate[];
/**
 * HWPX(ZIP) 파일에서 특정 XML 파일을 읽어서 DOM으로 파싱.
 */
export declare function readHwpxXml(hwpxPath: string, xmlName: string): Promise<Document>;
/**
 * HWPX(ZIP) 파일의 특정 XML을 수정 후 저장.
 * 기존 ZIP의 다른 파일은 그대로 유지.
 */
export declare function writeHwpxXml(sourcePath: string, outputPath: string, xmlName: string, doc: Document): Promise<void>;
/**
 * HWPX section XML에서 모든 텍스트를 추출.
 * 경로: hp:p > hp:run > hp:t
 */
export declare function extractTextFromSection(doc: Document): string[];
/**
 * HWPX section XML에서 텍스트 찾아 바꾸기.
 * CLAUDE.md 규칙: 수정 후 linesegarray 삭제 필수.
 */
export declare function replaceTextInSection(doc: Document, find: string, replace: string): number;
/**
 * HWPX section XML에서 텍스트 검색. COM FindReplace 우회용.
 */
export declare function searchTextInSection(doc: Document, searchText: string): {
    total: number;
    results: Array<{
        index: number;
        paragraph: number;
        context: string;
    }>;
};
/**
 * HWPX section XML에서 N번째 텍스트만 치환.
 */
export declare function replaceTextNthInSection(doc: Document, find: string, replace: string, nth: number): boolean;
/**
 * HWPX section XML에서 텍스트를 찾아 그 뒤에 추가.
 */
export declare function findAndAppendInSection(doc: Document, find: string, appendText: string): boolean;
export interface CellReplaceOptions {
    tableIndex: number;
    rowIndex: number;
    colIndex: number;
    find: string;
    replace: string;
    preserveCharPr?: boolean;
    occurrence?: number;
}
export interface CellReplaceResult {
    matched: number;
    cellText: string;
    charPrIDRef: string | null;
    warnings: string[];
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
export declare function replaceInTableCell(doc: Document, opts: CellReplaceOptions): CellReplaceResult;
export type FieldType = 'PageNum' | 'TotalPage' | 'Date' | 'Time' | 'TOC' | 'Index' | 'CrossRef' | 'FieldFormula' | 'all';
export interface FieldRecalcResult {
    marked: number;
    byType: Record<string, number>;
    unsupported: string[];
}
/**
 * v0.7.0 신규: 문서 내 모든 자동 계산 필드에 dirty="1" attribute 추가.
 * 한글이 파일을 다음 열 때 자동 재계산.
 *
 * 처리 대상: PageNum, TotalPage, Date, Time, TOC, Index, CrossRef, FieldFormula
 */
export declare function markFieldsForRecalc(doc: Document, fieldTypes?: FieldType[]): FieldRecalcResult;
export interface ComXmlStats {
    tables: number;
    paragraphs: number;
    fields: number;
    runs: number;
    chars: number;
}
export interface ComXmlDiff {
    com: ComXmlStats;
    xml: ComXmlStats;
    diff: ComXmlStats;
    ok: boolean;
    warnings: string[];
}
/**
 * v0.7.0 신규: HWPX XML과 COM 통계 비교.
 * Python COM에서 사전 수집한 stats를 입력으로 받아 XML 파싱 결과와 diff.
 * 임계: |Δtables| ≤ 0, |Δparagraphs| ≤ 1, |Δchars| ≤ 5
 */
export declare function compareCOMAndXML(filePath: string, comStats: ComXmlStats): Promise<ComXmlDiff>;
export interface NestedPathStep {
    tableIndex: number;
    row: number;
    col: number;
}
/**
 * v0.7.0 신규 (인터페이스만): 중첩 표 셀 텍스트 치환.
 * - path.length === 1: replaceInTableCell로 위임 (정식 지원)
 * - path.length >= 2: 'nested-table-experimental' warning + 첫 단계만 처리
 *
 * 정식 다단계 지원은 v0.7.2.1에서 구현 (재귀 처리).
 */
export declare function replaceInNestedTable(doc: Document, path: NestedPathStep[], find: string, replace: string): CellReplaceResult;
/**
 * 빈 HWPX 파일 생성.
 * BUG-9 fix: blank_template.hwpx가 없으면 프로그래밍적으로 생성.
 */
export declare function createBlankHwpx(outputPath: string, title?: string): Promise<void>;
/**
 * 템플릿 기반 문서 생성.
 * blank_template.hwpx를 복사 → 변수 치환.
 */
export declare function generateFromTemplate(templateId: string, variables: Record<string, string>, outputPath: string): Promise<{
    filledFields: number;
    emptyFields: string[];
}>;
