/**
 * Composite tools: smart analysis with completion rate
 */
import { z } from 'zod';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
const HWP_EXTENSIONS = new Set(['.hwp', '.hwpx']);
const ANALYSIS_TIMEOUT = 60000;
export function registerCompositeTools(server, bridge) {
    // ===== v0.7.5.4 P0-2: runAutoFixLoop NO-OP 전환 =====
    // v0.7.4.0 의 auto_fix loop 은 apply_style_profile({target: 'all'}) 로 전체 문서 서식을
    // 덮어쓰는 부작용이 있었음 (공무원 양식의 셀별 서식 손실). v0.7.5.4 부터는 validate 만
    // 수행하고 개별 수정은 Claude host 가 직접 set_paragraph_style 호출로 처리하도록 위임.
    //
    // 기존 호출자 (form_workflow / autopilot) 는 동일한 시그니처를 기대하므로 wrapper 유지.
    // iterations: 0 을 반환해서 "아무것도 안 했음" 명시.
    async function runAutoFixLoop(opts) {
        // validate 만 수행 — 전체 override 는 하지 않음
        let score = 100;
        const recommendedFixes = [];
        try {
            const params = { file_path: opts.outputPath };
            if (opts.styleProfile)
                params.expected_profile = opts.styleProfile;
            const r = await bridge.send('validate_consistency', params, ANALYSIS_TIMEOUT);
            if (r.success && r.data) {
                const d = r.data;
                if (typeof d.consistency_score === 'number')
                    score = d.consistency_score;
                // 발견된 deviations 를 recommended_fixes 로 반환 (Claude host 가 개별 수정)
                if (Array.isArray(d.format_deviations)) {
                    for (const dev of d.format_deviations) {
                        recommendedFixes.push({
                            field: dev.field,
                            expected: dev.expected,
                            actual: dev.actual,
                            suggestion: 'set_paragraph_style 또는 set_char_shape 로 개별 수정',
                        });
                    }
                }
            }
        }
        catch { }
        return {
            iterations: 0,
            score_before: score,
            score_after: score,
            log: [],
            stopped_reason: score >= opts.threshold ? 'already_passed' : 'no_auto_fix_v0754',
            recommended_fixes: recommendedFixes,
        };
    }
    // ── 진단 도구 (개발용) ──
    server.tool('hwp_inspect_com_object', '[개발용] pyhwpx COM 객체의 실제 속성 목록을 덤프합니다. HCharShape/HParaShape 등의 정확한 속성명을 확인할 때 사용.', {
        object: z.enum(['HCharShape', 'HParaShape', 'HFindReplace', 'HSecDef', 'HPageDef']).optional().describe('조사할 COM 객체 (기본: HCharShape)'),
    }, async ({ object: objName }) => {
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('inspect_com_object', { object: objName ?? 'HCharShape' }, 30000);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_generate_multi_documents', '하나의 템플릿으로 여러 건의 문서를 생성합니다. 각 데이터마다 템플릿을 별도 파일로 복사 후 채우기/치환하므로 AllReplace 범위 문제가 없습니다. 같은 양식에 여러 사람/기업 데이터를 채울 때 사용하세요.', {
        template_path: z.string().describe('템플릿 HWP/HWPX 파일 경로'),
        data_list: z.array(z.object({
            name: z.string().describe('출력 파일명 접미사 (예: "이준혁_(주)딥러닝코리아")'),
            table_cells: z.record(z.string(), z.array(z.object({
                tab: z.number().int().min(0).describe('Tab 인덱스'),
                text: z.string().describe('채울 텍스트'),
            }))).optional().describe('표 채우기 데이터 { "표인덱스": [{tab, text}, ...] }'),
            replacements: z.array(z.object({
                find: z.string().describe('찾을 텍스트'),
                replace: z.string().describe('바꿀 텍스트'),
            })).optional().describe('텍스트 치환 목록'),
            verify_tables: z.array(z.number().int().min(0)).optional().describe('채우기 후 검증할 표 인덱스 목록'),
        })).describe('각 문서별 데이터'),
        output_dir: z.string().optional().describe('출력 디렉토리 (생략 시 템플릿과 같은 폴더)'),
    }, async ({ template_path, data_list, output_dir }) => {
        const resolved = path.resolve(template_path);
        if (!fs.existsSync(resolved)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `템플릿 파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const params = {
                template_path: resolved,
                data_list,
            };
            if (output_dir)
                params.output_dir = path.resolve(output_dir);
            const response = await bridge.send('generate_multi_documents', params, ANALYSIS_TIMEOUT * 2);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_smart_fill', '표 셀 채우기 + 서식 자동 감지/보존. hwp_fill_table_cells와 달리 각 셀의 글꼴/크기/자간/장평을 자동 감지하고 유지합니다. 공공기관 문서처럼 서식이 중요한 경우 이 도구를 사용하세요. 적용된 서식 정보도 반환합니다.', {
        table_index: z.number().int().min(0).describe('표 인덱스'),
        cells: z.array(z.object({
            tab: z.number().int().min(0).describe('Tab 인덱스'),
            text: z.string().describe('채울 텍스트'),
        })).describe('채울 셀 목록'),
    }, async ({ table_index, cells }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.' }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('smart_fill', { table_index, cells }, ANALYSIS_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_auto_map_reference', '참고자료(Excel/CSV)의 헤더와 표의 라벨을 자동 매칭하여 채울 데이터를 생성합니다. 매핑 결과를 확인한 후 hwp_fill_table_cells로 실제 채우기를 진행하세요. hwp_read_reference로 데이터를 읽은 뒤 이 도구로 매핑하면 편리합니다.', {
        table_index: z.number().int().min(0).describe('표 인덱스'),
        ref_headers: z.array(z.string()).describe('참고자료 헤더 목록 (예: ["기업명", "대표자", "전화번호"])'),
        ref_row: z.array(z.string()).describe('참고자료 데이터 행 (헤더 순서에 맞춤)'),
    }, async ({ table_index, ref_headers, ref_row }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.' }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('auto_map_reference', { table_index, ref_headers, ref_row }, ANALYSIS_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_table_insert_from_csv', 'CSV 또는 Excel 파일을 읽어서 표로 자동 생성합니다. 현재 커서 위치에 헤더+데이터가 포함된 표가 삽입됩니다.', {
        file_path: z.string().describe('CSV 또는 Excel 파일 경로'),
    }, async ({ file_path }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        }
        const resolved = path.resolve(file_path);
        if (!fs.existsSync(resolved)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('table_insert_from_csv', { file_path: resolved }, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            bridge.setCachedAnalysis(null);
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_document_merge', '현재 열린 문서에 다른 HWP 문서의 내용을 합칩니다. 여러 문서를 하나로 합칠 때 사용하세요.', {
        file_path: z.string().describe('합칠 HWP 파일 경로'),
    }, async ({ file_path }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        }
        const resolved = path.resolve(file_path);
        if (!fs.existsSync(resolved)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('document_merge', { file_path: resolved }, ANALYSIS_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            bridge.setCachedAnalysis(null);
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_document_summary', '문서를 분석하고 빈 필드/셀을 강조하며 작성 완성도(%)를 계산합니다. 문서 상태를 한눈에 파악하고 다음 작업을 결정할 때 사용하세요.', {
        file_path: z.string().describe('HWP/HWPX 파일 경로'),
    }, async ({ file_path }) => {
        const resolved = path.resolve(file_path);
        if (!fs.existsSync(resolved)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
        }
        const ext = path.extname(resolved).toLowerCase();
        if (!HWP_EXTENSIONS.has(ext)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: 'HWP 또는 HWPX 파일만 지원합니다.' }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('analyze_document', { file_path: resolved }, ANALYSIS_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            const analysis = response.data;
            bridge.setCachedAnalysis(analysis);
            bridge.setCurrentDocument(resolved);
            const fields = analysis.fields || [];
            const emptyFields = fields.filter(f => !f.value || f.value.trim() === '');
            const tables = analysis.tables || [];
            let totalCells = 0;
            let emptyCells = 0;
            for (const table of tables) {
                for (const row of table.data) {
                    for (const cell of row) {
                        totalCells++;
                        if (!cell || cell.trim() === '') {
                            emptyCells++;
                        }
                    }
                }
            }
            const totalItems = fields.length + totalCells;
            const filledItems = (fields.length - emptyFields.length) + (totalCells - emptyCells);
            const completionRate = totalItems > 0
                ? Math.round((filledItems / totalItems) * 100)
                : 100;
            const parts = [];
            if (emptyFields.length > 0) {
                parts.push(`${emptyFields.length}개 빈 필드`);
            }
            if (emptyCells > 0) {
                parts.push(`${emptyCells}개 빈 셀`);
            }
            let recommendation;
            const nextActions = [];
            if (parts.length > 0) {
                recommendation = `${parts.join('과 ')}이 있습니다.`;
                if (emptyFields.length > 0) {
                    recommendation += ' hwp_fill_fields로 필드를 채울 수 있습니다.';
                    nextActions.push({ tool: 'hwp_fill_fields', reason: `${emptyFields.length}개 빈 필드 채우기` });
                }
                if (emptyCells > 0) {
                    recommendation += ' hwp_fill_table_cells로 표 셀을 채울 수 있습니다.';
                    nextActions.push({ tool: 'hwp_fill_table_cells', reason: `${emptyCells}개 빈 표 셀 채우기` });
                }
                nextActions.push({ tool: 'hwp_save_document', reason: '변경사항 저장' });
            }
            else {
                recommendation = '문서가 완전히 작성되었습니다.';
            }
            const summary = {
                file_name: analysis.file_name,
                file_format: analysis.file_format,
                pages: analysis.pages,
                table_count: tables.length,
                field_count: fields.length,
                empty_fields: emptyFields.map(f => ({ name: f.name })),
                empty_cell_count: emptyCells,
                total_cell_count: totalCells,
                completion_rate: `${completionRate}%`,
                text_preview: analysis.text_preview,
                recommendation,
                next_actions: nextActions,
            };
            return { content: [{ type: 'text', text: JSON.stringify(summary) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── 보안/개인정보 도구 ──
    server.tool('hwp_privacy_scan', '문서 텍스트에서 개인정보(주민번호, 전화번호, 이메일, 계좌번호 등)를 자동 감지합니다. 공공기관 문서 제출 전 개인정보 포함 여부를 확인할 때 사용하세요.', {
        file_path: z.string().optional().describe('HWP 파일 경로 (생략 시 현재 문서의 텍스트 스캔)'),
    }, async ({ file_path }) => {
        try {
            await bridge.ensureRunning();
            // 문서 텍스트 추출
            let text;
            if (file_path) {
                const resolved = path.resolve(file_path);
                const resp = await bridge.send('analyze_document', { file_path: resolved }, ANALYSIS_TIMEOUT);
                if (!resp.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: resp.error }) }], isError: true };
                }
                text = resp.data.full_text || '';
            }
            else {
                const current = bridge.getCurrentDocument();
                if (!current) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
                }
                const resp = await bridge.send('analyze_document', { file_path: current }, ANALYSIS_TIMEOUT);
                if (!resp.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: resp.error }) }], isError: true };
                }
                text = resp.data.full_text || '';
            }
            // Python에서 개인정보 스캔
            const scanResp = await bridge.send('privacy_scan', { text }, ANALYSIS_TIMEOUT);
            if (!scanResp.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: scanResp.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(scanResp.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── 차별화 복합 도구 ──
    server.tool('hwp_smart_analyze', '문서를 열고, 구조 분석 + 문서 타입 추론 + 서식 프로파일 + 완성도 + 추천 작업을 한번에 수행합니다. 문서를 처음 다룰 때 이 도구 하나면 충분합니다. analyze_document + document_summary + get_table_format_summary를 통합한 원스톱 분석 도구입니다.', {
        file_path: z.string().describe('HWP/HWPX 파일 경로'),
    }, async ({ file_path }) => {
        const resolved = path.resolve(file_path);
        if (!fs.existsSync(resolved)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            // 1. 문서 분석 (smart_analyze는 복합 도구이므로 90초 타임아웃)
            const analysisResp = await bridge.send('analyze_document', { file_path: resolved }, 90000);
            if (!analysisResp.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: analysisResp.error }) }], isError: true };
            }
            const analysis = analysisResp.data;
            bridge.setCachedAnalysis(analysis);
            bridge.setCurrentDocument(resolved);
            // 2. 완성도 계산
            const fields = analysis.fields || [];
            const emptyFields = fields.filter(f => !f.value || f.value.trim() === '');
            const tables = analysis.tables || [];
            let totalCells = 0, emptyCells = 0;
            for (const table of tables) {
                for (const row of table.data) {
                    for (const cell of row) {
                        totalCells++;
                        if (!cell || cell.trim() === '')
                            emptyCells++;
                    }
                }
            }
            const totalItems = fields.length + totalCells;
            const filledItems = (fields.length - emptyFields.length) + (totalCells - emptyCells);
            const completionRate = totalItems > 0 ? Math.round((filledItems / totalItems) * 100) : 100;
            // 3. 문서 타입 추론
            const fullText = (analysis.full_text || '').toLowerCase();
            let documentType = '일반 문서';
            const typePatterns = [
                ['사업계획서/신청서', ['사업계획', '신청서', '지원사업', '보조금', '참여기업']],
                ['공문서 (기안문/시행문)', ['수신자', '발신명의', '시행일자', '기안자', '결재']],
                ['보고서', ['보고서', '보고일', '현황', '문제점', '개선방안', '기대효과']],
                ['계약서', ['계약서', '갑', '을', '계약금', '계약기간']],
                ['이력서', ['이력서', '학력', '경력', '자격증', '자기소개']],
                ['회의록', ['회의록', '참석자', '안건', '결정사항', '향후계획']],
                ['견적서/인보이스', ['견적서', '인보이스', '품목', '단가', '합계']],
            ];
            let maxScore = 0;
            for (const [type, keywords] of typePatterns) {
                const score = keywords.filter(k => fullText.includes(k)).length;
                if (score > maxScore) {
                    maxScore = score;
                    documentType = type;
                }
            }
            // 4. 서식 프로파일 (첫 번째 데이터 표의 서식 샘플)
            let formatProfile = null;
            const dataTables = tables.filter(t => t.data.length > 0);
            const targetTable = dataTables.length > 0 ? dataTables[0] : (tables.length > 0 ? tables[0] : null);
            if (targetTable) {
                try {
                    const fmtResp = await bridge.send('get_table_format_summary', {
                        table_index: targetTable.index,
                    }, ANALYSIS_TIMEOUT);
                    if (fmtResp.success) {
                        formatProfile = fmtResp.data;
                    }
                }
                catch { /* 서식 조회 실패해도 계속 */ }
            }
            // 5. 추천 작업
            const recommendations = [];
            if (emptyCells > 0)
                recommendations.push(`hwp_fill_table_cells 또는 hwp_smart_fill로 ${emptyCells}개 빈 셀 채우기`);
            if (emptyFields.length > 0)
                recommendations.push(`hwp_fill_fields로 ${emptyFields.length}개 빈 필드 채우기`);
            if (completionRate >= 90)
                recommendations.push('hwp_save_document로 저장');
            if (documentType.includes('사업계획서'))
                recommendations.push('fill_public_document 프롬프트로 공문서 표준 작성');
            return {
                content: [{
                        type: 'text',
                        text: JSON.stringify({
                            file_name: analysis.file_name,
                            file_format: analysis.file_format,
                            document_type: documentType,
                            pages: analysis.pages,
                            table_count: tables.length,
                            field_count: fields.length,
                            completion_rate: `${completionRate}%`,
                            empty_cells: emptyCells,
                            empty_fields: emptyFields.length,
                            text_preview: analysis.text_preview,
                            format_profile: formatProfile,
                            recommendations,
                        }),
                    }],
            };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_auto_fill_from_reference', '엑셀/CSV → 자동 매핑 → 서식 보존 채우기를 일괄 수행하는 원스톱 도구. hwp_smart_fill + hwp_read_reference + hwp_auto_map_reference를 통합합니다. "이 엑셀 데이터로 신청서를 채워줘" 같은 요청에 사용하세요.', {
        file_path: z.string().describe('참고자료 파일 경로 (xlsx, csv, json)'),
        table_index: z.number().int().min(0).describe('채울 표 인덱스'),
        row_index: z.number().int().min(0).optional().describe('참고자료에서 사용할 행 번호 (0부터, 생략 시 첫 번째 행)'),
    }, async ({ file_path: refPath, table_index, row_index }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다. hwp_open_document 또는 hwp_smart_analyze로 문서를 먼저 열어주세요.' }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const resolvedRef = path.resolve(refPath);
            // 1. 참고자료 읽기
            const refResp = await bridge.send('read_reference', { file_path: resolvedRef }, ANALYSIS_TIMEOUT);
            if (!refResp.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `참고자료 읽기 실패: ${refResp.error}` }) }], isError: true };
            }
            const refData = refResp.data;
            // 헤더와 데이터 행 추출
            let headers = [];
            let dataRow = [];
            const format = refData.format;
            const ri = row_index ?? 0;
            if (format === 'csv') {
                headers = refData.headers || [];
                const data = refData.data || [];
                dataRow = data[ri] || [];
            }
            else if (format === 'excel') {
                const sheets = refData.sheets || [];
                if (sheets.length > 0) {
                    headers = sheets[0].headers || [];
                    dataRow = sheets[0].data?.[ri] || [];
                }
            }
            else if (format === 'json') {
                const jsonData = refData.data;
                if (Array.isArray(jsonData) && jsonData.length > 0) {
                    const obj = jsonData[ri] || jsonData[0];
                    headers = Object.keys(obj);
                    dataRow = Object.values(obj).map(v => String(v ?? ''));
                }
            }
            else {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `자동 매핑은 csv, xlsx, json만 지원합니다. (현재: ${format})` }) }], isError: true };
            }
            if (headers.length === 0) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '참고자료에서 헤더를 찾을 수 없습니다.' }) }], isError: true };
            }
            // 2. 자동 매핑
            const mapResp = await bridge.send('auto_map_reference', {
                table_index, ref_headers: headers, ref_row: dataRow,
            }, ANALYSIS_TIMEOUT);
            if (!mapResp.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `매핑 실패: ${mapResp.error}` }) }], isError: true };
            }
            const mapData = mapResp.data;
            if (mapData.mappings.length === 0) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                status: 'no_mapping',
                                message: '자동 매핑된 항목이 없습니다. 표의 라벨과 참고자료 헤더가 일치하지 않습니다.',
                                unmapped: mapData.unmapped,
                                ref_headers: headers,
                            }) }] };
            }
            // 3. 서식 보존 채우기 (smart_fill)
            const cells = mapData.mappings.map(m => ({ tab: m.tab, text: m.text }));
            const fillResp = await bridge.send('smart_fill', { table_index, cells }, ANALYSIS_TIMEOUT);
            if (!fillResp.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `채우기 실패: ${fillResp.error}` }) }], isError: true };
            }
            bridge.setCachedAnalysis(null);
            return {
                content: [{
                        type: 'text',
                        text: JSON.stringify({
                            status: 'ok',
                            reference_file: path.basename(resolvedRef),
                            mapped: mapData.mappings.map(m => ({ header: m.header, label: m.matched_label, value: m.text })),
                            unmapped: mapData.unmapped,
                            fill_result: fillResp.data,
                        }),
                    }],
            };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── Phase C: 복합 기능 ──
    server.tool('hwp_table_to_json', '표 데이터를 JSON 형식으로 추출합니다.', {
        table_index: z.number().int().min(0).describe('표 인덱스'),
    }, async ({ table_index }) => {
        if (!bridge.getCurrentDocument())
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('table_to_json', { table_index }, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_batch_convert', '폴더 내 모든 HWP 파일을 지정 형식으로 일괄 변환합니다.', {
        input_dir: z.string().describe('HWP 파일이 있는 디렉토리'),
        output_format: z.enum(['PDF', 'HTML', 'HWPX']).describe('변환할 형식'),
        output_dir: z.string().optional().describe('출력 디렉토리 (생략 시 input_dir)'),
    }, async ({ input_dir, output_format, output_dir }) => {
        try {
            await bridge.ensureRunning();
            const params = { input_dir: path.resolve(input_dir), output_format };
            if (output_dir)
                params.output_dir = path.resolve(output_dir);
            const r = await bridge.send('batch_convert', params, ANALYSIS_TIMEOUT * 5);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_compare_documents', '두 HWP 문서의 텍스트를 비교하여 차이점을 반환합니다.', {
        file_path_1: z.string().describe('첫 번째 문서 경로'),
        file_path_2: z.string().describe('두 번째 문서 경로'),
    }, async ({ file_path_1, file_path_2 }) => {
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('compare_documents', {
                file_path_1: path.resolve(file_path_1), file_path_2: path.resolve(file_path_2),
            }, ANALYSIS_TIMEOUT * 2);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── 목차 자동 생성 ──
    server.tool('hwp_generate_toc', '현재 문서의 제목 패턴(Ⅰ., 1., 가. 등)을 자동 감지하여 목차를 생성하고 현재 커서 위치에 삽입합니다.', {
        dot_leader: z.boolean().optional().describe('점선 리더 사용 여부 (기본 true)'),
    }, async ({ dot_leader }) => {
        if (!bridge.getCurrentDocument())
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        try {
            await bridge.ensureRunning();
            const params = {};
            if (dot_leader !== undefined)
                params.dot_leader = dot_leader;
            const r = await bridge.send('generate_toc', params, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            bridge.setCachedAnalysis(null);
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── 간트차트 추진일정 표 ──
    server.tool('hwp_create_gantt_chart', '추진일정 간트차트 표를 자동 생성합니다. 작업 목록과 기간을 입력하면 ■ 표시가 있는 일정표를 만듭니다.', {
        tasks: z.array(z.object({
            name: z.string().describe('작업명'),
            desc: z.string().optional().describe('수행내용'),
            start: z.number().int().min(1).describe('시작 월 (1부터)'),
            end: z.number().int().min(1).describe('종료 월'),
            weight: z.string().optional().describe('비중(%)'),
        })).describe('작업 목록'),
        months: z.number().int().min(1).max(24).describe('총 기간 (월 수)'),
        month_label: z.string().optional().describe('월 라벨 형식 (기본 "M+N")'),
    }, async ({ tasks, months, month_label }) => {
        if (!bridge.getCurrentDocument())
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        try {
            await bridge.ensureRunning();
            const params = { tasks, months };
            if (month_label)
                params.month_label = month_label;
            const r = await bridge.send('create_gantt_chart', params, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            bridge.setCachedAnalysis(null);
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_word_count', '현재 문서의 글자수, 단어수, 문단수, 페이지수를 반환합니다.', {}, async () => {
        if (!bridge.getCurrentDocument())
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('word_count', {}, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── Phase D: HWPX XML 엔진 도구 ──
    server.tool('hwp_template_list', '사용 가능한 22종 문서 템플릿 목록을 반환합니다. 공문서, 기업, 학술, 개인 카테고리별 템플릿을 확인할 수 있습니다. 한글 프로그램 없이 동작합니다.', {
        category: z.string().optional().describe('필터링할 카테고리 (공문서/기업/학술/개인, 생략 시 전체)'),
    }, async ({ category }) => {
        try {
            const { TEMPLATES } = await import('../hwpx-engine.js');
            let list = TEMPLATES;
            if (category) {
                list = TEMPLATES.filter(t => t.category === category);
            }
            return { content: [{ type: 'text', text: JSON.stringify({
                            templates: list.map(t => ({ id: t.id, name: t.name, category: t.category, fields: t.fields })),
                            total: list.length,
                            categories: [...new Set(TEMPLATES.map(t => t.category))],
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_document_create', '빈 HWPX 문서를 생성합니다. 한글 프로그램 없이 동작합니다. 생성된 파일은 한글에서 열 수 있습니다.', {
        output_path: z.string().describe('생성할 HWPX 파일 경로'),
        title: z.string().optional().describe('문서 제목 (선택)'),
    }, async ({ output_path, title }) => {
        try {
            const { createBlankHwpx } = await import('../hwpx-engine.js');
            const resolved = path.resolve(output_path);
            await createBlankHwpx(resolved, title);
            return { content: [{ type: 'text', text: JSON.stringify({
                            status: 'ok', path: resolved, title: title || '(빈 문서)',
                            note: '한글 프로그램에서 열어서 확인하세요.',
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_template_generate', '템플릿 기반으로 HWPX 문서를 생성합니다. 변수(기업명, 대표자 등)를 치환하여 완성된 문서를 만듭니다. 한글 프로그램 없이 동작합니다.', {
        template_id: z.string().describe('템플릿 ID (hwp_template_list로 확인)'),
        variables: z.record(z.string(), z.string()).describe('채울 변수 { "기업명": "플랜아이", "대표자": "이명기" }'),
        output_path: z.string().describe('생성할 HWPX 파일 경로'),
    }, async ({ template_id, variables, output_path }) => {
        try {
            const { generateFromTemplate } = await import('../hwpx-engine.js');
            const resolved = path.resolve(output_path);
            const result = await generateFromTemplate(template_id, variables, resolved);
            return { content: [{ type: 'text', text: JSON.stringify({
                            status: 'ok', path: resolved, template: template_id,
                            filled_fields: result.filledFields,
                            empty_fields: result.emptyFields,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_xml_edit_text', 'HWPX 파일의 텍스트를 직접 찾아 바꿉니다. 한글 프로그램 없이 동작합니다. HWPX(ZIP+XML) 내부의 텍스트를 수정합니다.', {
        file_path: z.string().describe('수정할 HWPX 파일 경로'),
        find: z.string().describe('찾을 텍스트'),
        replace: z.string().describe('바꿀 텍스트'),
        output_path: z.string().optional().describe('저장 경로 (생략 시 원본 덮어쓰기)'),
    }, async ({ file_path, find, replace, output_path }) => {
        try {
            const { readHwpxXml, writeHwpxXml, replaceTextInSection } = await import('../hwpx-engine.js');
            const resolved = path.resolve(file_path);
            const outResolved = output_path ? path.resolve(output_path) : resolved;
            if (!fs.existsSync(resolved)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
            }
            const doc = await readHwpxXml(resolved, 'Contents/section0.xml');
            const count = replaceTextInSection(doc, find, replace);
            await writeHwpxXml(resolved, outResolved, 'Contents/section0.xml', doc);
            return { content: [{ type: 'text', text: JSON.stringify({
                            status: 'ok', path: outResolved, find, replace,
                            replacements: count,
                            note: 'linesegarray가 자동 삭제되었습니다 (CLAUDE.md 규칙)',
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.0 신규: HWPX 표 셀 단위 텍스트 직접 편집 (XML 경로, 한글 프로그램 불필요)
    server.tool('hwp_xml_edit_table_cell', 'HWPX 파일의 특정 표 셀(row, col)에서 텍스트를 직접 찾아 바꿉니다. (v0.7.0 신규) 한글 프로그램 없이 XML 경로(tc → subList → p → run → t)로 처리. linesegarray 셀 내부만 자동 삭제. charPrIDRef(서식 ID)는 보존. .hwp는 미지원 (hwp_fill_table_cells 사용).', {
        file_path: z.string().describe('수정할 HWPX 파일 경로 (.hwpx만 지원)'),
        table_index: z.number().int().min(0).describe('0-based 평탄화 표 인덱스 (중첩 표는 v0.7.2.1)'),
        row: z.number().int().min(0).describe('0-based 행 인덱스'),
        col: z.number().int().min(0).describe('0-based 열 인덱스'),
        find: z.string().describe('찾을 텍스트'),
        replace: z.string().describe('바꿀 텍스트'),
        occurrence: z.number().int().min(0).optional().describe('치환 횟수 (0=전체, 1+=N번째). 기본 0'),
        output_path: z.string().optional().describe('저장 경로 (생략 시 원본 덮어쓰기)'),
    }, async ({ file_path, table_index, row, col, find, replace, occurrence, output_path }) => {
        try {
            if (!file_path.toLowerCase().endsWith('.hwpx')) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: 'XML_ONLY_HWPX: 이 도구는 .hwpx 파일만 지원합니다. .hwp 파일은 hwp_fill_table_cells (COM 경로) 사용.',
                                file_path,
                            }) }], isError: true };
            }
            const { readHwpxXml, writeHwpxXml, replaceInTableCell } = await import('../hwpx-engine.js');
            const resolved = path.resolve(file_path);
            const outResolved = output_path ? path.resolve(output_path) : resolved;
            if (!fs.existsSync(resolved)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
            }
            const doc = await readHwpxXml(resolved, 'Contents/section0.xml');
            const result = replaceInTableCell(doc, {
                tableIndex: table_index,
                rowIndex: row,
                colIndex: col,
                find,
                replace,
                occurrence: occurrence ?? 0,
            });
            await writeHwpxXml(resolved, outResolved, 'Contents/section0.xml', doc);
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true,
                            path: outResolved,
                            matched: result.matched,
                            cell_text: result.cellText,
                            char_pr_id_ref: result.charPrIDRef,
                            warnings: result.warnings,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.0 신규: HWPX 자동 계산 필드(목차/페이지번호 등) dirty mark 후 한글이 다음 열 때 자동 재계산
    server.tool('hwp_refresh_fields', 'HWPX 파일의 자동 계산 필드(목차, 페이지번호, 작성일, 인덱스 등)에 dirty="1" 마크를 추가하여 한글이 다음에 파일을 열 때 자동 재계산하도록 합니다. (v0.7.0 신규, composite tool) 처리 대상: PageNum, TotalPage, Date, Time, TOC, Index, CrossRef, FieldFormula. .hwp는 미지원.', {
        file_path: z.string().describe('대상 HWPX 파일 경로 (.hwpx만)'),
        field_types: z.array(z.enum(['PageNum', 'TotalPage', 'Date', 'Time', 'TOC', 'Index', 'CrossRef', 'FieldFormula', 'all'])).optional().describe('처리할 필드 종류 (기본: ["all"])'),
        output_path: z.string().optional().describe('저장 경로 (생략 시 원본 덮어쓰기)'),
    }, async ({ file_path, field_types, output_path }) => {
        try {
            if (!file_path.toLowerCase().endsWith('.hwpx')) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: 'XML_ONLY_HWPX: 이 도구는 .hwpx 파일만 지원합니다. .hwp는 한글에서 직접 F9 또는 export hwpx 후 사용.',
                                file_path,
                            }) }], isError: true };
            }
            const { readHwpxXml, writeHwpxXml, markFieldsForRecalc } = await import('../hwpx-engine.js');
            const resolved = path.resolve(file_path);
            const outResolved = output_path ? path.resolve(output_path) : resolved;
            if (!fs.existsSync(resolved)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
            }
            const doc = await readHwpxXml(resolved, 'Contents/section0.xml');
            const result = markFieldsForRecalc(doc, field_types);
            await writeHwpxXml(resolved, outResolved, 'Contents/section0.xml', doc);
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true,
                            path: outResolved,
                            marked: result.marked,
                            by_type: result.byType,
                            unsupported: result.unsupported,
                            note: '한글이 다음에 이 파일을 열 때 자동 재계산됩니다.',
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ──────────────────────────────────────────────────────────
    // v0.7.1: 양식 학습 + Workload Estimate (사용자 핵심 니즈)
    // ──────────────────────────────────────────────────────────
    // v0.7.1 신규: 양식의 트리 구조 추출 (목차/섹션/표/필드)
    server.tool('hwp_extract_template_structure', '양식 문서의 트리 구조(목차/섹션/표/필드)를 추출합니다. (v0.7.1 신규) heading 정규식 휴리스틱(제 N 장/조/절, I./II., 1./1.1, 가./나., (1)/(가))으로 섹션 인식. analyze_document + traverse_all_ctrls 재활용 (95%). 사용자 양식 분석의 진입점.', {
        file_path: z.string().describe('양식 파일 경로 (HWP/HWPX)'),
        max_depth: z.number().int().min(1).max(6).optional().describe('인식할 heading 깊이 (기본 4)'),
    }, async ({ file_path, max_depth }) => {
        if (!bridge.getCurrentDocument() && file_path) {
            try {
                await bridge.send('open_document', { file_path });
            }
            catch { }
        }
        try {
            await bridge.ensureRunning();
            const params = { file_path };
            if (max_depth !== undefined)
                params.max_depth = max_depth;
            const r = await bridge.send('extract_template_structure', params, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.1 신규: 양식의 서식 패턴 학습
    server.tool('hwp_analyze_writing_patterns', '양식의 서식 패턴(폰트/줄간격/들여쓰기/번호 체계)을 학습합니다. (v0.7.1 신규) extract_full_profile + extract_style_profile + get_table_format_summary 재활용 (90%). 출력은 후속 hwp_extend_section, hwp_apply_style_profile에 입력으로 사용.', {
        file_path: z.string().describe('양식 파일 경로'),
    }, async ({ file_path }) => {
        if (!bridge.getCurrentDocument()) {
            try {
                await bridge.send('open_document', { file_path });
            }
            catch { }
        }
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('analyze_writing_patterns', { file_path }, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.1 신규 ★ — Workload 사전 분석 (사용자 의사결정 도구)
    server.tool('hwp_estimate_workload', '★ 작성 작업의 워크로드를 사전 추정합니다. (v0.7.1 신규) 입력: 양식 + 사용자 요청 + 참고 자료. 출력: 예상 페이지 수, 토큰 사용량, 소요 시간, 위험 항목, 권장 조치(proceed/split_into_sessions/reduce_scope). 사용자가 결과 보고 진행 여부 결정. 추정 공식: chars_per_page=1100, tokens=chars/3.5, output_tokens_per_page=500, seconds_per_token=0.011 (Opus 4.6 한국어).', {
        file_path: z.string().optional().describe('양식 파일 경로 (옵션, 있으면 자동 분석)'),
        user_request: z.string().describe('사용자 요청 (예: "AI 스타트업 사업계획서 A4 10쪽 격식체")'),
        reference_files: z.array(z.string()).optional().describe('참고 자료 파일 경로 목록 (옵셔널)'),
        mode: z.enum(['new', 'extend']).optional().describe('작성 모드 (new=새 문서, extend=양식 확장)'),
        constraints: z.object({
            max_reference_files: z.number().int().optional().describe('참고 자료 최대 개수 (기본 5)'),
            max_reference_mb: z.number().int().optional().describe('참고 자료 최대 크기 MB (기본 10)'),
            context_window_tokens: z.number().int().optional().describe('컨텍스트 윈도우 토큰 (기본 200000)'),
        }).optional().describe('정책 제약'),
    }, async ({ file_path, user_request, reference_files, mode, constraints }) => {
        if (file_path && !bridge.getCurrentDocument()) {
            try {
                await bridge.send('open_document', { file_path });
            }
            catch { }
        }
        try {
            await bridge.ensureRunning();
            const params = { user_request };
            if (file_path)
                params.file_path = file_path;
            if (reference_files)
                params.reference_files = reference_files;
            if (mode)
                params.mode = mode;
            if (constraints)
                params.constraints = constraints;
            const r = await bridge.send('estimate_workload', params, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.1 신규: 기존 양식 섹션 확장
    server.tool('hwp_extend_section', '기존 양식의 특정 섹션 끝에 콘텐츠를 추가합니다. (v0.7.1 신규) 양식 서식 보존하며 LLM 생성 콘텐츠를 삽입. section_identifier로 섹션 위치를 제목 텍스트로 검색.', {
        section_identifier: z.object({
            by: z.enum(['title', 'index']).describe('식별 방식'),
            value: z.union([z.string(), z.number()]).describe('제목 텍스트 또는 인덱스'),
        }).describe('섹션 식별자'),
        content: z.string().describe('추가할 콘텐츠 (단락 단위, \\n으로 구분)'),
        preserve_format: z.boolean().optional().describe('양식 서식 보존 (기본 true)'),
    }, async ({ section_identifier, content, preserve_format }) => {
        if (!bridge.getCurrentDocument())
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('extend_section', { section_identifier, content, preserve_format: preserve_format ?? true }, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            bridge.setCachedAnalysis(null);
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.1 신규: 패턴 프로파일 일괄 적용
    server.tool('hwp_apply_style_profile', '추출된 서식 패턴 프로파일(WritingPatterns)을 현재 문서에 적용합니다. (v0.7.1 신규) hwp_analyze_writing_patterns의 출력을 입력으로 받아 set_paragraph_style 반복 호출.', {
        profile: z.object({
            body_style: z.any().optional(),
            title_styles: z.any().optional(),
            table_styles: z.any().optional(),
        }).passthrough().describe('WritingPatterns 객체'),
        target: z.enum(['all', 'section_index', 'range']).optional().describe('적용 대상 (기본 all)'),
    }, async ({ profile, target }) => {
        if (!bridge.getCurrentDocument())
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        try {
            await bridge.ensureRunning();
            const r = await bridge.send('apply_style_profile', { profile, target: target ?? 'all' }, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ===== v0.7.5.4 P2-1: Form Attachment Workflow Orchestrator (read-only default) =====
    // hwp_form_workflow — 양식 파일 첨부 → 학습 → 계획 → 미리보기 → (명시) 채우기 → (명시) 검증
    // v0.7.5.4 변경:
    //   - phase='all' 은 learn→plan→preview 에서 stop (이전: auto_fix 까지 자동 실행)
    //   - auto_fix 는 완전 opt-in, 기본값은 no-op (runAutoFixLoop 이 P0-2 에서 이미 무력화됨)
    //   - 원본 양식 절대 덮어쓰지 않음 (save_as 로 새 경로 저장만)
    server.tool('hwp_form_workflow', '양식 파일 첨부 워크플로우. v0.7.5.4 read-only 기본: phase="all" 은 learn→plan→preview 에서 정지 (fill 은 사용자 명시 호출 필요). phase 별: learn(학습)→plan(계획)→preview(미리보기)→fill(명시)→verify(명시)→rollback(원본 복원). auto_fix 는 v0.7.5.4 부터 no-op (원본 서식 보호). table_cell_overrides/field_overrides 로 Claude host 가 직접 제어.', {
        phase: z.enum(['learn', 'plan', 'preview', 'fill', 'verify', 'auto_fix', 'all', 'rollback']).describe('실행할 단계. "all" 은 read-only (learn→plan→preview 만)'),
        form_file: z.string().optional().describe('양식 HWP/HWPX 파일 (learn 또는 all 첫 호출 시 필수)'),
        user_request: z.string().optional().describe('사용자 요청 (plan 단계에서 estimate_workload 입력)'),
        reference_file: z.string().optional().describe('참고 자료 (Excel/CSV/JSON/PDF/DOCX/HTML)'),
        session_id: z.string().optional().describe('세션 ID (없으면 자동 생성)'),
        output_path: z.string().optional().describe('저장 경로 (생략 시 form_file 옆에 _filled 접미사)'),
        field_overrides: z.record(z.string(), z.string()).optional().describe('필드명→값 직접 지정 (auto_map 결과 덮어쓰기)'),
        table_cell_overrides: z.array(z.object({
            table_index: z.number().int().min(0),
            cells: z.array(z.object({
                tab: z.number().int().min(0).optional(),
                label: z.string().optional(),
                text: z.string(),
            })),
        })).optional().describe('표 셀 직접 지정 (auto_map 결과 덮어쓰기)'),
        confirm_fill: z.boolean().optional().describe('fill 단계 진행 확인 (사용자 승인 완료)'),
        auto_fix_enabled: z.boolean().optional().describe('v0.7.5.4: auto_fix 활성화 여부 (기본 false). true 여도 P0-2 runAutoFixLoop no-op 이므로 validate 만 수행.'),
        auto_fix_threshold: z.number().min(0).max(100).optional().describe('auto_fix 점수 임계 (기본 85, auto_fix_enabled=true 일 때만 의미)'),
        auto_fix_max_iterations: z.number().int().min(1).max(5).optional().describe('auto_fix 최대 반복 (기본 2, auto_fix_enabled=true 일 때만 의미)'),
    }, async (args) => {
        const sid = safeId(args.session_id || `form_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
        const sessionPath = path.join(STATE_DIR, `${sid}.json`);
        const snapshotPath = path.join(STATE_DIR, `${sid}.snapshot.bin`);
        if (!fs.existsSync(STATE_DIR))
            fs.mkdirSync(STATE_DIR, { recursive: true });
        const loadCheckpoint = () => {
            if (!fs.existsSync(sessionPath))
                return {};
            try {
                return JSON.parse(fs.readFileSync(sessionPath, 'utf8'));
            }
            catch {
                return {};
            }
        };
        const saveCheckpoint = (extra) => {
            const existing = loadCheckpoint();
            const merged = {
                ...existing,
                ...extra,
                session_id: sid,
                workflow: 'form_workflow',
                last_saved: new Date().toISOString(),
            };
            fs.writeFileSync(sessionPath, JSON.stringify(merged, null, 2), 'utf8');
            return merged;
        };
        try {
            await bridge.ensureRunning();
            const cp = loadCheckpoint();
            const formFile = args.form_file ?? cp.form_file;
            const outputPath = args.output_path
                ?? cp.output_path
                ?? (formFile ? formFile.replace(/(\.hwpx?)$/i, '_filled$1') : undefined);
            // ─── Phase 1: LEARN ───
            const runLearn = async () => {
                if (!formFile)
                    throw new Error('form_file required for learn phase');
                // v0.7.4.9 S2-NEW-2 Fix: 이전 문서 (특히 hwp_pdf_clone 의 FileNew 결과) 를 명시 close
                // → cursor state 정리. 실패해도 무시 (no document to close).
                try {
                    await bridge.send('close_document', {}, 5000);
                }
                catch { }
                const openR = await bridge.send('open_document', { file_path: formFile }, ANALYSIS_TIMEOUT);
                if (!openR.success)
                    throw new Error(`open_document failed: ${openR.error}`);
                let formDetect = null;
                try {
                    const r = await bridge.send('form_detect', {}, ANALYSIS_TIMEOUT);
                    if (r.success)
                        formDetect = r.data;
                }
                catch { }
                let structure = null;
                try {
                    const r = await bridge.send('extract_template_structure', { file_path: formFile }, ANALYSIS_TIMEOUT);
                    if (r.success)
                        structure = r.data;
                }
                catch { }
                let patterns = null;
                try {
                    const r = await bridge.send('analyze_writing_patterns', { file_path: formFile }, ANALYSIS_TIMEOUT);
                    if (r.success)
                        patterns = r.data;
                }
                catch { }
                // Enumerate tables and map cells
                let tablesInfo = [];
                let fieldsInfo = [];
                try {
                    const r = await bridge.send('analyze_document', { file_path: formFile }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data) {
                        const d = r.data;
                        tablesInfo = Array.isArray(d.tables) ? d.tables : [];
                        fieldsInfo = Array.isArray(d.fields) ? d.fields : [];
                    }
                }
                catch { }
                const tableMaps = [];
                for (let i = 0; i < tablesInfo.length; i++) {
                    try {
                        const m = await bridge.send('map_table_cells', { table_index: i }, ANALYSIS_TIMEOUT);
                        if (m.success && m.data) {
                            tableMaps.push({ table_index: i, ...m.data });
                        }
                    }
                    catch { }
                }
                return saveCheckpoint({
                    phase: 'learn',
                    form_file: formFile,
                    output_path: outputPath,
                    form_detect: formDetect,
                    template_structure: structure,
                    writing_patterns: patterns,
                    tables_info: tablesInfo,
                    fields_info: fieldsInfo,
                    table_maps: tableMaps,
                });
            };
            // ─── Phase 2: PLAN ───
            const runPlan = async () => {
                const learned = (cp.phase === 'learn' || cp.table_maps) ? cp : await runLearn();
                let estimate = {};
                try {
                    const estParams = {
                        user_request: args.user_request || '양식 채우기',
                    };
                    if (formFile)
                        estParams.file_path = formFile;
                    if (args.reference_file)
                        estParams.reference_files = [args.reference_file];
                    const r = await bridge.send('estimate_workload', estParams, ANALYSIS_TIMEOUT);
                    if (r.success)
                        estimate = r.data;
                }
                catch { }
                // Heuristic: detect "label-shaped" cells (ends with colon or short Korean noun)
                // v0.7.4.9 S3-NEW-1 Fix: Python hwp_analyzer.map_table_cells 는 "cell_map" 키로 반환하지만
                // 이전 v0.7.4.8 은 "cells" 로 잘못 읽어서 suggested_fills 가 항상 0 이었음.
                // 호환: cell_map 우선, 레거시 cells 도 fallback
                const suggestedFills = [];
                const tableMaps = learned.table_maps || [];
                for (const tm of tableMaps) {
                    const cells = tm.cell_map
                        || tm.cells
                        || [];
                    for (const c of cells) {
                        const txt = String(c.text || '').trim();
                        if (txt && txt.length <= 20 && /[:：]$|^[가-힣]{2,8}$/.test(txt)) {
                            suggestedFills.push({
                                table_index: tm.table_index,
                                label: txt,
                                tab: c.tab,
                                suggested_value: '',
                            });
                        }
                    }
                }
                return saveCheckpoint({
                    phase: 'plan',
                    estimate,
                    suggested_fills: suggestedFills,
                    status: 'awaiting_preview',
                });
            };
            // ─── Phase 3: PREVIEW ───
            const runPreview = () => {
                return saveCheckpoint({
                    phase: 'preview',
                    status: 'awaiting_confirm',
                    preview: {
                        form_file: cp.form_file,
                        output_path: cp.output_path,
                        form_detect: cp.form_detect,
                        suggested_fills: cp.suggested_fills,
                        estimate: cp.estimate,
                        instructions: 'Claude host: 사용자 확인 후 phase=fill, confirm_fill=true, field_overrides/table_cell_overrides 를 포함해 재호출',
                    },
                });
            };
            // ─── Phase 4: FILL ───
            const runFill = async () => {
                if (!args.confirm_fill) {
                    return saveCheckpoint({ phase: 'fill', status: 'denied', reason: 'confirm_fill=false' });
                }
                if (!formFile)
                    throw new Error('form_file missing in checkpoint');
                if (!outputPath)
                    throw new Error('output_path missing');
                // Snapshot pre-fill state for rollback (항상 원본 form_file 을 백업)
                if (fs.existsSync(formFile)) {
                    try {
                        fs.copyFileSync(formFile, snapshotPath);
                    }
                    catch { }
                }
                // Ensure document is open
                const openR = await bridge.send('open_document', { file_path: formFile }, ANALYSIS_TIMEOUT);
                if (!openR.success)
                    throw new Error(`open_document failed: ${openR.error}`);
                const fillResults = [];
                // 4a) reference auto-map (reference_file 제공 시)
                let refMappings = null;
                if (args.reference_file) {
                    try {
                        const refR = await bridge.send('read_reference', { file_path: args.reference_file }, ANALYSIS_TIMEOUT);
                        if (refR.success && refR.data) {
                            const refData = refR.data;
                            // Extract headers and first data row (tolerant of different schemas)
                            let headers = [];
                            let row = [];
                            // v0.7.4.8 Fix C3: hwp_structured (.hwp/.hwpx) 우선 지원 — tables[0] 의 headers/data 사용
                            if (refData.format === 'hwp_structured' && Array.isArray(refData.tables)) {
                                const tables = refData.tables;
                                if (tables.length > 0) {
                                    const firstTable = tables[0];
                                    headers = firstTable.headers || [];
                                    const dataRows = firstTable.data;
                                    if (Array.isArray(dataRows) && dataRows.length > 0) {
                                        row = dataRows[0].map(String);
                                    }
                                }
                            }
                            if (headers.length === 0 && Array.isArray(refData.headers)) {
                                headers = refData.headers;
                                const dataRows = refData.data;
                                if (Array.isArray(dataRows) && dataRows.length > 0) {
                                    const first = dataRows[0];
                                    row = Array.isArray(first) ? first.map(String) : Object.values(first).map(String);
                                }
                            }
                            else if (headers.length === 0 && Array.isArray(refData.sheets) && refData.sheets.length > 0) {
                                const sheet0 = refData.sheets[0];
                                headers = sheet0.headers || [];
                                const dataRows = sheet0.data;
                                if (Array.isArray(dataRows) && dataRows.length > 0) {
                                    row = dataRows[0].map(String);
                                }
                            }
                            const tableMaps = cp.table_maps || [];
                            // v0.7.4.8 Fix B4: 모든 테이블 순회 (기존: firstTableIdx 만)
                            // 각 테이블 별로 auto_map 호출 후 결과를 flatten 해 refMappings 배열로
                            if (headers.length > 0 && tableMaps.length > 0) {
                                const allMappings = [];
                                const perTableResults = [];
                                for (const tm of tableMaps) {
                                    const tblIdx = tm.table_index ?? 0;
                                    try {
                                        const m = await bridge.send('auto_map_reference', { table_index: tblIdx, ref_headers: headers, ref_row: row }, ANALYSIS_TIMEOUT);
                                        if (m.success && m.data) {
                                            const mData = m.data;
                                            const tblMappings = mData.mappings || [];
                                            for (const mp of tblMappings) {
                                                allMappings.push({ ...mp, table_index: tblIdx });
                                            }
                                            perTableResults.push({
                                                table_index: tblIdx,
                                                total_matched: mData.total_matched || 0,
                                                unmapped: mData.unmapped || [],
                                            });
                                        }
                                    }
                                    catch (e) {
                                        perTableResults.push({
                                            table_index: tblIdx,
                                            error: e.message,
                                        });
                                    }
                                }
                                if (allMappings.length > 0) {
                                    refMappings = {
                                        mappings: allMappings,
                                        per_table: perTableResults,
                                        total_matched: allMappings.length,
                                    };
                                }
                            }
                        }
                    }
                    catch { }
                }
                // 4b) Apply table_cell_overrides (Claude host 의 권위 있는 값)
                if (args.table_cell_overrides) {
                    for (const t of args.table_cell_overrides) {
                        try {
                            const r = await bridge.send('smart_fill', {
                                table_index: t.table_index,
                                cells: t.cells.map(c => ({ tab: c.tab, label: c.label, text: c.text })),
                            }, ANALYSIS_TIMEOUT);
                            fillResults.push({ type: 'table_cell_overrides', table_index: t.table_index, ok: r.success, data: r.data, error: r.error });
                        }
                        catch (e) {
                            fillResults.push({ type: 'table_cell_overrides', table_index: t.table_index, ok: false, error: e.message });
                        }
                    }
                }
                // 4c) Apply ref auto-map (overrides 로 채워지지 않은 셀만)
                // v0.7.4.8 Fix B4: refMappings.mappings 가 여러 table 의 mapping 을 병합한 배열 (각 항목에 table_index 포함)
                if (refMappings && Array.isArray(refMappings.mappings)) {
                    const overrideTabs = new Set((args.table_cell_overrides ?? [])
                        .flatMap(t => t.cells
                        .filter(c => c.tab !== undefined)
                        .map(c => `${t.table_index}:${c.tab}`)));
                    // Group by table_index → 각 테이블별로 smart_fill 호출
                    const byTable = new Map();
                    for (const mp of refMappings.mappings) {
                        const tIdx = mp.table_index ?? 0;
                        const key = `${tIdx}:${mp.tab}`;
                        if (overrideTabs.has(key))
                            continue;
                        if (!byTable.has(tIdx))
                            byTable.set(tIdx, []);
                        byTable.get(tIdx).push({ tab: mp.tab, text: mp.text });
                    }
                    for (const [tIdx, cells] of byTable) {
                        if (cells.length === 0)
                            continue;
                        try {
                            const r = await bridge.send('smart_fill', { table_index: tIdx, cells }, ANALYSIS_TIMEOUT);
                            fillResults.push({ type: 'ref_auto_map', table_index: tIdx, ok: r.success, cells: cells.length, data: r.data, error: r.error });
                        }
                        catch (e) {
                            fillResults.push({ type: 'ref_auto_map', table_index: tIdx, ok: false, error: e.message });
                        }
                    }
                }
                // 4d) Apply field_overrides
                if (args.field_overrides && Object.keys(args.field_overrides).length > 0) {
                    try {
                        const r = await bridge.send('fill_fields', { fields: args.field_overrides }, ANALYSIS_TIMEOUT);
                        fillResults.push({ type: 'field_overrides', ok: r.success, count: Object.keys(args.field_overrides).length, data: r.data, error: r.error });
                    }
                    catch (e) {
                        fillResults.push({ type: 'field_overrides', ok: false, error: e.message });
                    }
                }
                // Save to output_path
                const saveFmt = outputPath.toLowerCase().endsWith('.hwpx') ? 'HWPX' : 'HWP';
                const saveR = await bridge.send('save_as', { path: outputPath, format: saveFmt }, ANALYSIS_TIMEOUT);
                // v0.7.4.8 Fix B5: save 성공 시 runVerify 자동 호출 — 5단계 cross-check
                // (이전: runVerify 는 phase="verify" 명시 호출 시에만 작동)
                let autoVerify = null;
                if (saveR.success) {
                    try {
                        const verifyResult = await runVerify();
                        autoVerify = verifyResult.verify || null;
                    }
                    catch (e) {
                        autoVerify = { error: e.message };
                    }
                }
                return saveCheckpoint({
                    phase: 'fill',
                    status: saveR.success ? 'filled' : 'fill_failed',
                    fill_results: fillResults,
                    ref_mappings: refMappings,
                    save_result: saveR.data,
                    save_error: saveR.error,
                    output_path: outputPath,
                    rollback_snapshot: snapshotPath,
                    // v0.7.4.8: auto-verify 결과
                    auto_verify: autoVerify,
                });
            };
            // ─── Phase 5: VERIFY ───
            const runVerify = async () => {
                if (!outputPath || !fs.existsSync(outputPath)) {
                    throw new Error('output file not saved yet — run phase=fill first');
                }
                const stat = fs.statSync(outputPath);
                const minBytes = outputPath.toLowerCase().endsWith('.hwpx') ? 19000 : 24000;
                const sizeOk = stat.size >= minBytes;
                let wcData = null;
                try {
                    const r = await bridge.send('word_count', {}, ANALYSIS_TIMEOUT);
                    if (r.success && r.data)
                        wcData = r.data;
                }
                catch { }
                let textOk = false;
                let textPreview = '';
                try {
                    const r = await bridge.send('get_document_text', { file_path: outputPath, max_chars: 5000 }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data) {
                        const d = r.data;
                        const txt = d.text || d.full_text || '';
                        textOk = txt.length > 100;
                        textPreview = txt.slice(0, 300);
                    }
                }
                catch { }
                let consistencyScore = null;
                try {
                    const params = { file_path: outputPath };
                    if (cp.writing_patterns)
                        params.expected_profile = cp.writing_patterns;
                    const r = await bridge.send('validate_consistency', params, ANALYSIS_TIMEOUT);
                    if (r.success && r.data) {
                        const d = r.data;
                        if (typeof d.consistency_score === 'number')
                            consistencyScore = d.consistency_score;
                    }
                }
                catch { }
                let privacy = null;
                try {
                    const r = await bridge.send('privacy_scan', { file_path: outputPath }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data)
                        privacy = r.data;
                }
                catch { }
                const crossCheckPassed = sizeOk && textOk && consistencyScore !== null && consistencyScore >= 50;
                // v0.7.5.4 P4-2: 5단계 검증 자동 병합 (TEST_CHECKLIST Phase 19)
                let verify5Stage = null;
                try {
                    const r = await bridge.send('verify_5stage', {
                        file_path: outputPath,
                        expected_text_snippet: textPreview.slice(0, 100),
                        run_layout: false,
                    }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data)
                        verify5Stage = r.data;
                }
                catch { }
                return saveCheckpoint({
                    phase: 'verify',
                    verify: {
                        file_size: stat.size,
                        min_required: minBytes,
                        file_size_ok: sizeOk,
                        word_count: wcData,
                        text_ok: textOk,
                        text_preview: textPreview,
                        consistency_score: consistencyScore,
                        privacy,
                        cross_check_passed: crossCheckPassed,
                        // v0.7.5.4 P4-2: 5단계 검증 결과 병합
                        verify_5stage: verify5Stage,
                        overall_pass: verify5Stage?.overall_pass ?? crossCheckPassed,
                    },
                });
            };
            // ─── Phase 6: AUTO_FIX ───
            const runAutoFixPhase = async () => {
                if (!outputPath)
                    throw new Error('output_path missing');
                const saveFmt = outputPath.toLowerCase().endsWith('.hwpx') ? 'HWPX' : 'HWP';
                const loopResult = await runAutoFixLoop({
                    outputPath,
                    styleProfile: cp.writing_patterns,
                    threshold: args.auto_fix_threshold ?? 85,
                    maxIter: Math.min(args.auto_fix_max_iterations ?? 2, 5),
                    saveFmt,
                });
                return saveCheckpoint({
                    phase: 'auto_fix',
                    auto_fix: loopResult,
                });
            };
            // ─── Rollback ───
            const runRollback = () => {
                if (!formFile)
                    throw new Error('form_file missing');
                if (!fs.existsSync(snapshotPath))
                    throw new Error('no rollback snapshot — fill phase 를 먼저 실행해야 합니다');
                fs.copyFileSync(snapshotPath, formFile);
                return saveCheckpoint({
                    phase: 'rollback',
                    status: 'restored',
                    restored_from: snapshotPath,
                    restored_to: formFile,
                });
            };
            // ─── Dispatcher ───
            let result;
            if (args.phase === 'learn')
                result = await runLearn();
            else if (args.phase === 'plan')
                result = await runPlan();
            else if (args.phase === 'preview')
                result = runPreview();
            else if (args.phase === 'fill')
                result = await runFill();
            else if (args.phase === 'verify')
                result = await runVerify();
            else if (args.phase === 'auto_fix')
                result = await runAutoFixPhase();
            else if (args.phase === 'rollback')
                result = runRollback();
            else if (args.phase === 'all') {
                // learn → plan → preview 자동. fill 은 confirm 필요하므로 제외.
                await runLearn();
                await runPlan();
                result = runPreview();
            }
            else {
                throw new Error(`unknown phase: ${args.phase}`);
            }
            return { content: [{ type: 'text', text: JSON.stringify({ ok: true, session_id: sid, ...result }) }] };
        }
        catch (err) {
            saveCheckpoint({ phase: args.phase, status: 'failed', error: err.message });
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message, session_id: sid }) }], isError: true };
        }
    });
    // ===== v0.7.5.4 P3-4: 공무원 양식 원스톱 자동 작성 =====
    // 사업계획서/공문/보고서 양식 첨부 시 한 번에 detect → snapshot → delete guide → fill → save → verify
    // 원본 양식 절대 보존 (read-only). 사용자 최종 목표: 공무원 사무업무 자동화.
    server.tool('hwp_korean_business_fill', '공무원 양식 원스톱 자동 작성. (v0.7.5.4 신규) 양식 파일 + 본문 채우기 맵만 주면 자동으로: (1) 문서 타입 감지 + 공무원 표준 프리셋 선택, (2) 원본 서식 스냅샷 캐시, (3) 작성요령 박스 정리 (scope=both), (4) 표 셀 + 본문 일괄 삽입, (5) 새 경로 저장 (원본 절대 보존), (6) 5단계 검증. 사업계획서/공문/보고서 자동화의 권장 진입점.', {
        form_file: z.string().describe('양식 HWP/HWPX 파일 절대 경로'),
        output_path: z.string().describe('저장할 새 파일 경로 (원본과 달라야 함)'),
        body_fills: z.array(z.object({
            heading: z.string().describe('소제목 텍스트 (예: "(1) 산업의 특성")'),
            body_text: z.string().describe('삽입할 본문'),
        })).optional().describe('본문 삽입 맵 (heading → body_text)'),
        table_cell_overrides: z.array(z.object({
            table_index: z.number().int().min(0),
            cells: z.array(z.object({
                tab: z.number().int().min(0),
                text: z.string(),
            })),
        })).optional().describe('표 셀 직접 지정'),
        doc_type_hint: z.enum(['business_plan', 'official_document', 'report', 'form', 'general']).optional().describe('문서 타입 힌트 (생략 시 자동 감지)'),
        delete_guides: z.boolean().optional().describe('작성요령 자동 삭제 (기본 true)'),
        run_verify: z.boolean().optional().describe('5단계 검증 실행 (기본 true)'),
    }, async (args) => {
        const startTime = Date.now();
        const log = [];
        const pushLog = (step, data) => log.push({ step, elapsed_ms: Date.now() - startTime, ...data });
        try {
            // 경로 검증
            const formPath = path.resolve(args.form_file);
            const outPath = path.resolve(args.output_path);
            if (formPath === outPath) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: 'output_path 는 form_file 과 달라야 합니다 (원본 보호).' }) }], isError: true };
            }
            if (!fs.existsSync(formPath)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `양식 파일이 없습니다: ${formPath}` }) }], isError: true };
            }
            const originalMtime = fs.statSync(formPath).mtime.getTime();
            const originalSize = fs.statSync(formPath).size;
            await bridge.ensureRunning();
            // Step 1: Open form
            const openR = await bridge.send('open_document', { file_path: formPath }, ANALYSIS_TIMEOUT);
            pushLog('open_document', { ok: openR.success, pages: openR.data?.pages });
            if (!openR.success)
                throw new Error(`open failed: ${openR.error}`);
            // Step 2: Detect document type — hint 있으면 type_override 로 프리셋 강제 (v0.7.9 fix)
            let docType = args.doc_type_hint || 'general';
            let preset = null;
            try {
                const detectParams = {};
                if (args.doc_type_hint)
                    detectParams.type_override = args.doc_type_hint;
                const r = await bridge.send('detect_document_type', detectParams, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const d = r.data;
                    if (!args.doc_type_hint)
                        docType = d.type || 'general';
                    preset = d.recommended_preset || null;
                    pushLog('detect_document_type', { type: docType, confidence: d.confidence });
                }
            }
            catch (e) {
                pushLog('detect_document_type', { error: e.message });
            }
            // Step 3: Snapshot template style (optional, 실패해도 continue)
            try {
                const r = await bridge.send('snapshot_template_style', {}, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    pushLog('snapshot_template_style', { snapshot_id: r.data.snapshot_id });
                }
            }
            catch (e) {
                pushLog('snapshot_template_style', { error: e.message });
            }
            // Step 4: Extract + Delete guide text (v0.7.7: extract_first)
            const deleteGuides = args.delete_guides ?? true;
            let extractedGuides = [];
            if (deleteGuides) {
                try {
                    const r = await bridge.send('delete_guide_text', { scope: 'both', extract_first: true }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data) {
                        const d = r.data;
                        extractedGuides = d.extracted_guides || [];
                        pushLog('delete_guide_text', { ...d, extracted_guide_count: extractedGuides.length });
                    }
                }
                catch (e) {
                    pushLog('delete_guide_text', { error: e.message });
                }
            }
            // Step 5: Fill table cells (if provided)
            // fill_by_tab 을 각 table 별로 반복 호출 (composite 내부 분할)
            const tableFillResults = [];
            if (args.table_cell_overrides && args.table_cell_overrides.length > 0) {
                for (const tbl of args.table_cell_overrides) {
                    try {
                        const r = await bridge.send('fill_by_tab', {
                            table_index: tbl.table_index,
                            cells: tbl.cells,
                        }, 120000);
                        if (r.success && r.data) {
                            tableFillResults.push({ table_index: tbl.table_index, ...r.data });
                        }
                        else {
                            tableFillResults.push({ table_index: tbl.table_index, error: r.error });
                        }
                    }
                    catch (e) {
                        tableFillResults.push({ table_index: tbl.table_index, error: e.message });
                    }
                }
                pushLog('fill_by_tab_batch', { tables_processed: tableFillResults.length });
            }
            const tableFillResult = tableFillResults.length > 0 ? { tables: tableFillResults } : null;
            // Step 6: Insert body after each heading
            const bodyResults = [];
            if (args.body_fills && args.body_fills.length > 0) {
                // v0.7.9: 양식 서식 우선 — preset body_style 전달하지 않음
                // insert_body_after_heading 이 양식 heading에서 직접 서식 상속
                // (자간, 장평, 글꼴, 줄간격 등 양식 원본 100% 유지)
                for (const fill of args.body_fills) {
                    try {
                        const params = {
                            heading: fill.heading,
                            body_text: fill.body_text,
                        };
                        const r = await bridge.send('insert_body_after_heading', params, ANALYSIS_TIMEOUT);
                        if (r.success && r.data) {
                            const d = r.data;
                            bodyResults.push({ heading: fill.heading, status: d.status, total_matches: d.total_matches });
                        }
                        else {
                            bodyResults.push({ heading: fill.heading, status: 'error', error: r.error });
                        }
                    }
                    catch (e) {
                        bodyResults.push({ heading: fill.heading, status: 'error', error: e.message });
                    }
                }
                pushLog('insert_body_after_heading_batch', { count: bodyResults.length, success: bodyResults.filter(r => r.status === 'ok').length });
            }
            // Step 7: Save as (새 경로, 원본 절대 보존)
            const saveFmt = outPath.toLowerCase().endsWith('.hwpx') ? 'HWPX' : 'HWP';
            const saveR = await bridge.send('save_as', { path: outPath, format: saveFmt }, ANALYSIS_TIMEOUT);
            pushLog('save_as', { ok: saveR.success, path: outPath });
            if (!saveR.success)
                throw new Error(`save failed: ${saveR.error}`);
            // Step 8: 원본 mtime 검증 (원본 보존 확인)
            const postMtime = fs.statSync(formPath).mtime.getTime();
            const postSize = fs.statSync(formPath).size;
            const originalPreserved = (originalMtime === postMtime) && (originalSize === postSize);
            pushLog('original_preservation_check', {
                preserved: originalPreserved,
                original_mtime: originalMtime,
                post_mtime: postMtime,
                original_size: originalSize,
                post_size: postSize,
            });
            // Step 9: 5단계 검증 (if enabled)
            let verify5Stage = null;
            if (args.run_verify !== false) {
                try {
                    const snippet = args.body_fills?.[0]?.body_text?.slice(0, 50) || '';
                    const r = await bridge.send('verify_5stage', {
                        file_path: outPath,
                        expected_text_snippet: snippet,
                        run_layout: false,
                    }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data)
                        verify5Stage = r.data;
                    pushLog('verify_5stage', { overall_pass: verify5Stage?.overall_pass, passed_stages: verify5Stage?.passed_stages });
                }
                catch (e) {
                    pushLog('verify_5stage', { error: e.message });
                }
            }
            return {
                content: [{
                        type: 'text',
                        text: JSON.stringify({
                            status: 'ok',
                            form_file: formPath,
                            output_path: outPath,
                            doc_type: docType,
                            preset_used: preset ? (preset.description || docType) : 'none',
                            original_preserved: originalPreserved,
                            extracted_guides: extractedGuides.length > 0 ? extractedGuides : undefined,
                            table_fill: tableFillResult,
                            body_fills: bodyResults,
                            verify_5stage: verify5Stage,
                            overall_pass: verify5Stage?.overall_pass ?? null,
                            total_duration_ms: Date.now() - startTime,
                            log,
                        }),
                    }],
            };
        }
        catch (err) {
            return {
                content: [{ type: 'text', text: JSON.stringify({
                            error: err.message,
                            log,
                        }) }],
                isError: true,
            };
        }
    });
    // ─── v0.7.9 — 사업계획서 자동분석 (prepare 단계) ───────────────────────
    server.tool('hwp_business_plan_prepare', '★ 사업계획서 양식을 분석하여 AI 자동 작성 컨텍스트를 생성합니다 (v0.7.9). 양식 파일(+선택적 참고자료)을 입력하면: (1) 문서 타입 감지, (2) 양식 구조 추출, (3) 작성요령 파싱, (4) 참고자료 읽기+섹션 매핑, (5) 섹션별 AI 컨텍스트 반환. 이 결과로 AI가 body_fills 를 생성한 후 hwp_korean_business_fill 로 실제 삽입합니다.', {
        form_file: z.string().describe('양식 HWP/HWPX 파일 절대 경로'),
        reference_files: z.array(z.string()).optional().describe('참고자료 파일 경로 목록 (PDF/HWP/Excel/DOCX)'),
    }, async (args) => {
        const startTime = Date.now();
        const log = [];
        const pushLog = (step, data) => log.push({ step, elapsed_ms: Date.now() - startTime, ...data });
        try {
            const formPath = path.resolve(args.form_file);
            if (!fs.existsSync(formPath)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `양식 파일 없음: ${formPath}` }) }], isError: true };
            }
            await bridge.ensureRunning();
            // Step 1: Open form
            const openR = await bridge.send('open_document', { file_path: formPath }, ANALYSIS_TIMEOUT);
            pushLog('open_document', { ok: openR.success, pages: openR.data?.pages });
            if (!openR.success)
                throw new Error(`open failed: ${openR.error}`);
            // Step 2: Detect document type
            let docType = 'general';
            let preset = null;
            try {
                const r = await bridge.send('detect_document_type', {}, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const d = r.data;
                    docType = d.type || 'general';
                    preset = d.recommended_preset || null;
                    pushLog('detect_document_type', { type: docType, confidence: d.confidence });
                }
            }
            catch (e) {
                pushLog('detect_document_type', { error: e.message });
            }
            // Step 3: Extract template structure
            let templateSections = [];
            try {
                const r = await bridge.send('extract_template_structure', { file_path: formPath }, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const d = r.data;
                    templateSections = d.sections || [];
                    pushLog('extract_template_structure', { sections: templateSections.length });
                }
            }
            catch (e) {
                pushLog('extract_template_structure', { error: e.message });
            }
            // Step 4: Snapshot template style
            let snapshotId = '';
            try {
                const r = await bridge.send('snapshot_template_style', {}, ANALYSIS_TIMEOUT);
                if (r.success && r.data)
                    snapshotId = String(r.data.snapshot_id || '');
                pushLog('snapshot_template_style', { snapshot_id: snapshotId });
            }
            catch (e) {
                pushLog('snapshot_template_style', { error: e.message });
            }
            // Step 5: Extract guide text (before deletion)
            let guideConstraints = [];
            try {
                const r = await bridge.send('extract_guide_text', {}, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    guideConstraints = r.data.guides || [];
                    pushLog('extract_guide_text', { count: guideConstraints.length });
                }
            }
            catch (e) {
                pushLog('extract_guide_text', { error: e.message });
            }
            // Step 6: Read reference files (if provided)
            let refData = {};
            if (args.reference_files && args.reference_files.length > 0) {
                const allTexts = [];
                const allTables = [];
                for (const refFile of args.reference_files) {
                    try {
                        const r = await bridge.send('read_reference', { file_path: path.resolve(refFile) }, ANALYSIS_TIMEOUT);
                        if (r.success && r.data) {
                            const d = r.data;
                            if (d.full_text)
                                allTexts.push(String(d.full_text));
                            else if (d.text)
                                allTexts.push(String(d.text));
                            if (d.tables)
                                allTables.push(...d.tables);
                        }
                    }
                    catch (e) {
                        pushLog('read_reference', { file: refFile, error: e.message });
                    }
                }
                refData = { full_text: allTexts.join('\n\n---\n\n'), tables: allTables };
                pushLog('read_references', { files: args.reference_files.length, total_chars: allTexts.join('').length, tables: allTables.length });
            }
            // Step 7: Map reference to sections
            let sectionMappings = [];
            try {
                const r = await bridge.send('map_reference_to_sections', {
                    reference_data: refData,
                    template_sections: templateSections,
                    guide_constraints: guideConstraints,
                }, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const d = r.data;
                    sectionMappings = d.section_mappings || [];
                    pushLog('map_reference_to_sections', { total: d.total_sections, with_data: d.sections_with_data });
                }
            }
            catch (e) {
                pushLog('map_reference_to_sections', { error: e.message });
            }
            // Step 8: Build section contexts
            let contexts = [];
            try {
                const r = await bridge.send('build_section_context', {
                    section_mappings: sectionMappings,
                    template_style: snapshotId ? { snapshot_id: snapshotId } : {},
                }, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const d = r.data;
                    contexts = d.contexts || [];
                    pushLog('build_section_context', { total: d.total, with_ref: d.sections_with_reference, total_chars: d.total_suggested_chars });
                }
            }
            catch (e) {
                pushLog('build_section_context', { error: e.message });
            }
            // Close document (prepare 단계 완료, 실제 삽입은 korean_business_fill 에서)
            try {
                await bridge.send('close_document', {}, 10000);
            }
            catch (_) { /* ignore */ }
            return {
                content: [{
                        type: 'text',
                        text: JSON.stringify({
                            status: 'ok',
                            form_file: formPath,
                            doc_type: docType,
                            preset: preset ? { description: preset.description } : null,
                            template_sections: templateSections.length,
                            guide_constraints: guideConstraints,
                            section_contexts: contexts,
                            total_duration_ms: Date.now() - startTime,
                            log,
                            hint: 'section_contexts 의 각 항목으로 body_fills 를 생성한 후 hwp_korean_business_fill 의 body_fills 파라미터로 전달하세요. heading 은 fuzzy matching 지원됩니다.',
                        }),
                    }],
            };
        }
        catch (err) {
            return {
                content: [{ type: 'text', text: JSON.stringify({ error: err.message, log }) }],
                isError: true,
            };
        }
    });
    // v0.7.2.1 신규: 중첩 표 셀 텍스트 직접 편집 (재귀 path)
    server.tool('hwp_xml_edit_nested_cell', 'HWPX 중첩 표(표 안의 표)의 특정 셀 텍스트를 재귀 경로로 편집합니다. (v0.7.2.1 신규) path 배열로 중첩 깊이 표현 — 예: [{tableIndex:0,row:0,col:0},{tableIndex:0,row:1,col:1}]은 "0번 표 (0,0) 셀 안의 0번 nested 표 (1,1) 셀". v0.7.0 hwp_xml_edit_table_cell의 다단계 확장. linesegarray 셀 내부만 자동 삭제, charPrIDRef 보존.', {
        file_path: z.string().describe('수정할 HWPX 파일 경로'),
        path: z.array(z.object({
            tableIndex: z.number().int().min(0).describe('이 단계의 표 인덱스 (각 단계는 부모 셀 안의 평탄화 인덱스)'),
            row: z.number().int().min(0).describe('0-based 행'),
            col: z.number().int().min(0).describe('0-based 열'),
        })).min(1).describe('중첩 경로 배열 (length=1: 단일 셀, length≥2: 재귀)'),
        find: z.string().describe('찾을 텍스트'),
        replace: z.string().describe('바꿀 텍스트'),
        output_path: z.string().optional().describe('저장 경로 (생략 시 원본 덮어쓰기)'),
    }, async ({ file_path, path: cellPath, find, replace, output_path }) => {
        try {
            if (!file_path.toLowerCase().endsWith('.hwpx')) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: 'XML_ONLY_HWPX: 이 도구는 .hwpx 파일만 지원합니다.',
                                file_path,
                            }) }], isError: true };
            }
            const { readHwpxXml, writeHwpxXml, replaceInNestedTable } = await import('../hwpx-engine.js');
            const resolved = path.resolve(file_path);
            const outResolved = output_path ? path.resolve(output_path) : resolved;
            if (!fs.existsSync(resolved)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
            }
            const doc = await readHwpxXml(resolved, 'Contents/section0.xml');
            const result = replaceInNestedTable(doc, cellPath, find, replace);
            await writeHwpxXml(resolved, outResolved, 'Contents/section0.xml', doc);
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true,
                            path: outResolved,
                            path_depth: cellPath.length,
                            matched: result.matched,
                            cell_text: result.cellText,
                            char_pr_id_ref: result.charPrIDRef,
                            warnings: result.warnings,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.2.1 신규: 모든 표(중첩 포함)를 트리로 열거
    server.tool('hwp_enumerate_nested_tables', 'HWPX 문서의 모든 표(중첩 포함)를 트리 구조로 열거합니다. (v0.7.2.1 신규) DFS 순회로 top-level 표 + 각 셀 내부의 nested 표를 재귀적으로 발견. 출력: NestedTableNode[] (path, rows, cols, children).', {
        file_path: z.string().describe('대상 HWPX 파일 경로'),
    }, async ({ file_path }) => {
        try {
            if (!file_path.toLowerCase().endsWith('.hwpx')) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: 'XML_ONLY_HWPX', file_path }) }], isError: true };
            }
            const { readHwpxXml, enumerateNestedTables } = await import('../hwpx-engine.js');
            const resolved = path.resolve(file_path);
            if (!fs.existsSync(resolved)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
            }
            const doc = await readHwpxXml(resolved, 'Contents/section0.xml');
            const tree = enumerateNestedTables(doc);
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true,
                            path: resolved,
                            top_level_count: tree.length,
                            tree,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.1 신규: 작성된 결과의 양식 일관성 검증
    server.tool('hwp_validate_consistency', '작성된 문서의 양식 일관성을 검증합니다. (v0.7.1 신규) expected_profile(WritingPatterns)과 비교하여 deviations와 0~100 점수 반환. 작성 중간/완료 후 호출하여 양식 준수 확인.', {
        file_path: z.string().describe('검증 대상 파일 경로'),
        expected_profile: z.object({
            body_style: z.any().optional(),
        }).passthrough().optional().describe('기대 프로파일 (없으면 placeholder 100점)'),
    }, async ({ file_path, expected_profile }) => {
        if (!bridge.getCurrentDocument()) {
            try {
                await bridge.send('open_document', { file_path });
            }
            catch { }
        }
        try {
            await bridge.ensureRunning();
            const params = { file_path };
            if (expected_profile)
                params.expected_profile = expected_profile;
            const r = await bridge.send('validate_consistency', params, ANALYSIS_TIMEOUT);
            if (!r.success)
                return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
            return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ===== v0.7.2.2: Reference Policy + Session State + Template Library =====
    // 모두 순수 JSON I/O — Python 경유 불필요. 사용자 home 디렉토리 기반 영속화.
    // v0.7.2.5: path traversal 가드 — id는 영문/숫자/언더스코어/하이픈만 허용
    const safeId = (id) => {
        if (typeof id !== 'string' || !/^[a-zA-Z0-9_\-]+$/.test(id)) {
            throw new Error(`invalid id (allowed: [a-zA-Z0-9_-]): ${id}`);
        }
        return id;
    };
    const CONFIG_PATH = path.join(os.homedir(), '.hwp_studio_config.json');
    const STATE_DIR = path.join(os.homedir(), '.hwp_studio_state');
    const TEMPLATE_DIR = path.join(os.homedir(), '.hwp_studio_templates');
    // v0.7.4.8 Part 4: 볼륨 가이드 정책 — 권장 임계값 + 경고 기준
    // 🟢 최적 1-3개 / 총 60KB, 🟡 적정 ≤5개 / 150KB, 🟠 주의 >5개 / >150KB, 🔴 최대 10개 / 500KB
    const DEFAULT_POLICY = {
        max_reference_files: 10, // v0.7.4.8: 5 → 10 (기존 허용, 최대값)
        max_total_size_mb: 5, // v0.7.4.8: 10 → 5 (기본 엄격하게 — 권장 상한 150KB 기준 여유)
        max_tokens_input_percent: 80,
        allowed_formats: ['xlsx', 'csv', 'json', 'pdf', 'docx', 'txt', 'html', 'xml', 'pptx', 'hwp', 'hwpx'],
        prefer_summary: true,
        summary_threshold_mb: 3,
        // v0.7.4.8 신규: 권장 임계값 (경고 발생 기준)
        recommend_threshold_kb: 150, // 이 값 초과 시 warning 반환 (focus degradation 시작)
        optimal_threshold_kb: 60, // 이 값 이하면 최적 (LLM focus 최상)
        optimal_file_count: 3, // 이 값 이하면 최적 파일 수
    };
    function readConfig() {
        try {
            if (fs.existsSync(CONFIG_PATH)) {
                const raw = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
                return { reference_policy: { ...DEFAULT_POLICY, ...(raw.reference_policy || {}) } };
            }
        }
        catch { }
        return { reference_policy: { ...DEFAULT_POLICY } };
    }
    function writeConfig(cfg) {
        fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), 'utf8');
    }
    // v0.7.2.2 — 1c) hwp_reference_policy
    server.tool('hwp_reference_policy', '참고자료 정책 메타 도구. (v0.7.2.2 신규) ~/.hwp_studio_config.json에 max_reference_files/max_total_size_mb/max_tokens_input_percent/allowed_formats/prefer_summary 저장. mode: get|set|reset', {
        mode: z.enum(['get', 'set', 'reset']).describe('동작 모드'),
        policy: z.object({
            max_reference_files: z.number().int().positive().optional(),
            max_total_size_mb: z.number().positive().optional(),
            max_tokens_input_percent: z.number().min(1).max(100).optional(),
            allowed_formats: z.array(z.string()).optional(),
            prefer_summary: z.boolean().optional(),
            summary_threshold_mb: z.number().positive().optional(),
            // v0.7.4.8 신규 필드
            recommend_threshold_kb: z.number().positive().optional().describe('v0.7.4.8: 이 값 초과 시 warning (기본 150KB)'),
            optimal_threshold_kb: z.number().positive().optional().describe('v0.7.4.8: 이 값 이하면 최적 (기본 60KB)'),
            optimal_file_count: z.number().int().positive().optional().describe('v0.7.4.8: 이 값 이하면 최적 파일 수 (기본 3)'),
        }).optional().describe('set 모드 시 부분 업데이트할 정책 필드'),
    }, async ({ mode, policy }) => {
        try {
            if (mode === 'reset') {
                writeConfig({ reference_policy: { ...DEFAULT_POLICY } });
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, reference_policy: DEFAULT_POLICY }) }] };
            }
            if (mode === 'set') {
                const current = readConfig();
                const merged = { reference_policy: { ...current.reference_policy, ...(policy || {}) } };
                writeConfig(merged);
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, ...merged }) }] };
            }
            // get
            const cfg = readConfig();
            return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, ...cfg, config_path: CONFIG_PATH }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.2.2 — 2a) hwp_session_state ★
    server.tool('hwp_session_state', '긴 작성 작업의 진행 상태를 저장/재개합니다. (v0.7.2.2 신규 ★) ~/.hwp_studio_state/{session_id}.json에 sections_total/done, current_section, progress_percent, checkpoints 영속화. mode: save|load|list|delete|cancel', {
        mode: z.enum(['save', 'load', 'list', 'delete', 'cancel']).describe('동작 모드'),
        session_id: z.string().optional().describe('세션 ID (save 시 미지정이면 자동 생성)'),
        state: z.object({
            current_doc: z.string().optional(),
            sections_total: z.array(z.string()).optional(),
            sections_done: z.array(z.string()).optional(),
            current_section: z.string().optional(),
            progress_percent: z.number().optional(),
            checkpoints: z.array(z.any()).optional(),
        }).passthrough().optional().describe('save 모드 시 저장할 상태 객체'),
    }, async ({ mode, session_id, state }) => {
        try {
            if (!fs.existsSync(STATE_DIR))
                fs.mkdirSync(STATE_DIR, { recursive: true });
            if (mode === 'list') {
                const files = fs.readdirSync(STATE_DIR).filter(f => f.endsWith('.json'));
                const sessions = files.map(f => {
                    try {
                        const raw = JSON.parse(fs.readFileSync(path.join(STATE_DIR, f), 'utf8'));
                        return {
                            session_id: raw.session_id || f.replace('.json', ''),
                            current_doc: raw.current_doc,
                            progress_percent: raw.progress_percent,
                            last_saved: raw.last_saved,
                            cancelled: raw.cancelled || false,
                        };
                    }
                    catch {
                        return { session_id: f.replace('.json', ''), error: 'parse_failed' };
                    }
                });
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, count: sessions.length, sessions }) }] };
            }
            if (mode === 'save') {
                const sid = safeId(session_id || `sess_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
                const filePath = path.join(STATE_DIR, `${sid}.json`);
                let existing = {};
                if (fs.existsSync(filePath)) {
                    try {
                        existing = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                    }
                    catch { }
                }
                const merged = {
                    ...existing,
                    ...(state || {}),
                    session_id: sid,
                    last_saved: new Date().toISOString(),
                };
                fs.writeFileSync(filePath, JSON.stringify(merged, null, 2), 'utf8');
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, session_id: sid, state: merged }) }] };
            }
            if (!session_id) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `session_id required for mode=${mode}` }) }], isError: true };
            }
            const safeSid = safeId(session_id);
            const filePath = path.join(STATE_DIR, `${safeSid}.json`);
            if (mode === 'load') {
                if (!fs.existsSync(filePath)) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: `session not found: ${session_id}` }) }], isError: true };
                }
                const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, state: raw }) }] };
            }
            if (mode === 'delete') {
                if (fs.existsSync(filePath))
                    fs.unlinkSync(filePath);
                // v0.7.4.1: form_workflow 가 남긴 snapshot 바이너리도 함께 정리
                const snapshotPath = path.join(STATE_DIR, `${safeId(session_id)}.snapshot.bin`);
                if (fs.existsSync(snapshotPath)) {
                    try {
                        fs.unlinkSync(snapshotPath);
                    }
                    catch { }
                }
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, session_id, deleted: true }) }] };
            }
            if (mode === 'cancel') {
                if (!fs.existsSync(filePath)) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: `session not found: ${session_id}` }) }], isError: true };
                }
                const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                raw.cancelled = true;
                raw.last_saved = new Date().toISOString();
                fs.writeFileSync(filePath, JSON.stringify(raw, null, 2), 'utf8');
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, session_id, cancelled: true }) }] };
            }
            return { content: [{ type: 'text', text: JSON.stringify({ error: `unknown mode: ${mode}` }) }], isError: true };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // v0.7.2.2 — 2c) hwp_template_library
    server.tool('hwp_template_library', '문서 템플릿 라이브러리. (v0.7.2.2 신규) 사용자 등록 템플릿을 ~/.hwp_studio_templates/{id}.json + files/{id}.hwpx에 저장. mode: list|get|register|delete|search. 빌트인 35개 템플릿은 기존 hwp_template_list 도구 사용.', {
        mode: z.enum(['list', 'get', 'register', 'delete', 'search']).describe('동작 모드'),
        template_id: z.string().optional().describe('템플릿 ID (get/delete/register 필수)'),
        template: z.object({
            name: z.string(),
            description: z.string().optional(),
            tags: z.array(z.string()).optional(),
            category: z.string().optional(),
            source_path: z.string().optional().describe('register 시 복사할 .hwpx 원본 경로'),
        }).passthrough().optional().describe('register 모드 메타데이터'),
        query: z.string().optional().describe('search 모드 검색어'),
    }, async ({ mode, template_id, template, query }) => {
        try {
            if (!fs.existsSync(TEMPLATE_DIR))
                fs.mkdirSync(TEMPLATE_DIR, { recursive: true });
            const filesDir = path.join(TEMPLATE_DIR, 'files');
            if (!fs.existsSync(filesDir))
                fs.mkdirSync(filesDir, { recursive: true });
            const loadAll = () => {
                const metas = fs.readdirSync(TEMPLATE_DIR).filter(f => f.endsWith('.json'));
                return metas.map(f => {
                    try {
                        const raw = JSON.parse(fs.readFileSync(path.join(TEMPLATE_DIR, f), 'utf8'));
                        return raw;
                    }
                    catch {
                        return null;
                    }
                }).filter(Boolean);
            };
            if (mode === 'list') {
                const items = loadAll();
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, count: items.length, templates: items }) }] };
            }
            if (mode === 'search') {
                const q = (query || '').toLowerCase().trim();
                if (!q) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: 'query required for search' }) }], isError: true };
                }
                const items = loadAll();
                const scored = items.map(item => {
                    const name = String(item.name || '').toLowerCase();
                    const desc = String(item.description || '').toLowerCase();
                    const tags = (item.tags || []).map(t => t.toLowerCase());
                    let score = 0;
                    if (name.includes(q))
                        score += 3;
                    for (const t of tags)
                        if (t.includes(q))
                            score += 2;
                    if (desc.includes(q))
                        score += 1;
                    if (item.last_used)
                        score += 0.5;
                    return { ...item, _score: score };
                }).filter(x => x._score > 0).sort((a, b) => b._score - a._score);
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, count: scored.length, results: scored }) }] };
            }
            if (mode === 'get') {
                if (!template_id)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: 'template_id required' }) }], isError: true };
                const tid = safeId(template_id);
                const metaPath = path.join(TEMPLATE_DIR, `${tid}.json`);
                if (!fs.existsSync(metaPath)) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: `template not found: ${tid}` }) }], isError: true };
                }
                const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
                const filePath = path.join(filesDir, `${tid}.hwpx`);
                meta.file_path = fs.existsSync(filePath) ? filePath : null;
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, template: meta }) }] };
            }
            if (mode === 'register') {
                if (!template_id || !template) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: 'template_id + template required' }) }], isError: true };
                }
                const tid = safeId(template_id);
                const metaPath = path.join(TEMPLATE_DIR, `${tid}.json`);
                const meta = {
                    template_id: tid,
                    name: template.name,
                    description: template.description || '',
                    tags: template.tags || [],
                    category: template.category || 'user',
                    registered_at: new Date().toISOString(),
                };
                if (template.source_path) {
                    const src = String(template.source_path);
                    if (!fs.existsSync(src)) {
                        return { content: [{ type: 'text', text: JSON.stringify({ error: `source_path not found: ${src}` }) }], isError: true };
                    }
                    const dest = path.join(filesDir, `${tid}.hwpx`);
                    fs.copyFileSync(src, dest);
                    meta.file_path = dest;
                }
                fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2), 'utf8');
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, template: meta }) }] };
            }
            if (mode === 'delete') {
                if (!template_id)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: 'template_id required' }) }], isError: true };
                const tid = safeId(template_id);
                const metaPath = path.join(TEMPLATE_DIR, `${tid}.json`);
                const filePath = path.join(filesDir, `${tid}.hwpx`);
                if (fs.existsSync(metaPath))
                    fs.unlinkSync(metaPath);
                if (fs.existsSync(filePath))
                    fs.unlinkSync(filePath);
                return { content: [{ type: 'text', text: JSON.stringify({ ok: true, mode, template_id: tid, deleted: true }) }] };
            }
            return { content: [{ type: 'text', text: JSON.stringify({ error: `unknown mode: ${mode}` }) }], isError: true };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ===== v0.7.2.3: Review + Compare + Progress Polling =====
    // 2b) hwp_review_and_edit — validate_consistency + privacy_scan + (옵션) auto_fix
    server.tool('hwp_review_and_edit', '문서 종합 리뷰. (v0.7.2.3 신규) consistency/privacy/formatting 검사 후 점수 산출. auto_fix=true면 안전한 자동 수정 시도. score_before/after 반환.', {
        file_path: z.string().describe('리뷰 대상 파일'),
        checks: z.array(z.enum(['consistency', 'privacy', 'formatting', 'typos'])).optional().describe('실행할 검사 (기본 consistency+privacy)'),
        auto_fix: z.boolean().optional().describe('자동 수정 시도 여부 (기본 false)'),
        expected_profile: z.any().optional().describe('consistency 검사용 기대 프로파일'),
    }, async ({ file_path, checks, auto_fix, expected_profile }) => {
        try {
            const checkSet = new Set(checks && checks.length > 0 ? checks : ['consistency', 'privacy']);
            const issues = [];
            const auto_fixed = [];
            const requires_manual = [];
            let score_before = 100;
            if (!bridge.getCurrentDocument()) {
                try {
                    await bridge.send('open_document', { file_path });
                }
                catch { }
            }
            await bridge.ensureRunning();
            if (checkSet.has('consistency')) {
                const params = { file_path };
                if (expected_profile)
                    params.expected_profile = expected_profile;
                const r = await bridge.send('validate_consistency', params, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const data = r.data;
                    const cscore = typeof data.consistency_score === 'number' ? data.consistency_score : 100;
                    score_before = Math.min(score_before, cscore);
                    const deviations = data.deviations || [];
                    for (const d of deviations)
                        issues.push({ check: 'consistency', detail: d });
                }
            }
            if (checkSet.has('privacy')) {
                const r = await bridge.send('privacy_scan', { file_path }, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const data = r.data;
                    const findings = data.findings || data.matches || [];
                    for (const f of findings) {
                        issues.push({ check: 'privacy', severity: 'high', detail: f });
                        requires_manual.push({ check: 'privacy', detail: f, reason: '개인정보는 사람 확인 필수' });
                    }
                    if (findings.length > 0)
                        score_before = Math.min(score_before, Math.max(0, 100 - findings.length * 10));
                }
            }
            if (checkSet.has('typos')) {
                requires_manual.push({ check: 'typos', reason: 'pyhwpx SpellCheck 미지원 — 외부 도구 필요' });
            }
            if (checkSet.has('formatting')) {
                // formatting은 consistency에 사실상 포함됨
            }
            let score_after = score_before;
            if (auto_fix && issues.length > 0) {
                // v0.7.3 #9: consistency deviations 중 안전한 항목만 자동 수정.
                // (1) consistency 류 deviation 에 대해 apply_style_profile 호출 시도
                // (2) privacy 는 수동 필수 — 자동 fix 안 함 (이미 requires_manual 에 분류)
                const consistencyIssues = issues.filter(i => i.check === 'consistency');
                if (consistencyIssues.length > 0 && expected_profile) {
                    try {
                        const fixR = await bridge.send('apply_style_profile', { profile: expected_profile, target: 'all' }, ANALYSIS_TIMEOUT);
                        if (fixR.success) {
                            auto_fixed.push({ check: 'consistency', method: 'apply_style_profile', applied: fixR.data });
                            // 재검증
                            const reR = await bridge.send('validate_consistency', { file_path, expected_profile }, ANALYSIS_TIMEOUT);
                            if (reR.success && reR.data) {
                                const newScore = reR.data.consistency_score;
                                if (typeof newScore === 'number')
                                    score_after = newScore;
                            }
                        }
                        else {
                            auto_fixed.push({ check: 'consistency', method: 'apply_style_profile', error: fixR.error });
                        }
                    }
                    catch (e) {
                        auto_fixed.push({ check: 'consistency', error: e.message });
                    }
                }
                else if (consistencyIssues.length > 0) {
                    auto_fixed.push({ note: 'expected_profile 없음 → consistency auto_fix skip' });
                }
            }
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true, file_path, checks: Array.from(checkSet),
                            issues, auto_fixed, requires_manual,
                            score_before, score_after,
                            issue_count: issues.length,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // 2f) hwp_compare_with_template
    server.tool('hwp_compare_with_template', '결과 문서를 템플릿과 비교하여 format/structure/content 점수 산출. (v0.7.2.3 신규) extract_template_structure + analyze_writing_patterns 양쪽 호출 후 가중 합산. self vs self = 100.', {
        result_path: z.string().describe('비교 대상 결과 문서'),
        template_path: z.string().describe('기준 템플릿 문서'),
        weight: z.object({
            format: z.number().min(0).max(1).optional(),
            structure: z.number().min(0).max(1).optional(),
            content: z.number().min(0).max(1).optional(),
        }).optional().describe('가중치 (기본 0.4/0.4/0.2)'),
    }, async ({ result_path, template_path, weight }) => {
        try {
            const w = { format: 0.4, structure: 0.4, content: 0.2, ...(weight || {}) };
            await bridge.ensureRunning();
            // 템플릿 구조 추출
            const tStruct = await bridge.send('extract_template_structure', { file_path: template_path }, ANALYSIS_TIMEOUT);
            const rStruct = await bridge.send('extract_template_structure', { file_path: result_path }, ANALYSIS_TIMEOUT);
            const tPattern = await bridge.send('analyze_writing_patterns', { file_path: template_path }, ANALYSIS_TIMEOUT);
            const rPattern = await bridge.send('analyze_writing_patterns', { file_path: result_path }, ANALYSIS_TIMEOUT);
            if (!tStruct.success || !rStruct.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: 'extract_template_structure failed', t: tStruct.error, r: rStruct.error }) }], isError: true };
            }
            const tData = (tStruct.data || {});
            const rData = (rStruct.data || {});
            const tSections = tData.sections || [];
            const rSections = rData.sections || [];
            const tNames = new Set(tSections.map(s => String(s.title || s.name || '')));
            const rNames = new Set(rSections.map(s => String(s.title || s.name || '')));
            const missing_sections = [...tNames].filter(n => !rNames.has(n));
            const extra_sections = [...rNames].filter(n => !tNames.has(n));
            const matched = [...tNames].filter(n => rNames.has(n)).length;
            const structure_score = tNames.size === 0 ? 100 : Math.round((matched / tNames.size) * 100);
            // format: writing patterns 비교 (v0.7.2.5: Python 실제 shape body_style:{char,para}에 맞춤)
            const tPat = (tPattern.data || {});
            const rPat = (rPattern.data || {});
            const format_deviations = [];
            let format_match = 0;
            let format_total = 0;
            const tBody = (tPat.body_style || {});
            const rBody = (rPat.body_style || {});
            for (const sub of ['char', 'para']) {
                const tSub = tBody[sub] || {};
                const rSub = rBody[sub] || {};
                for (const key of Object.keys(tSub)) {
                    format_total++;
                    if (JSON.stringify(tSub[key]) === JSON.stringify(rSub[key])) {
                        format_match++;
                    }
                    else {
                        format_deviations.push({ field: `body_style.${sub}.${key}`, expected: tSub[key], actual: rSub[key] });
                    }
                }
            }
            const format_score = format_total === 0 ? 100 : Math.round((format_match / format_total) * 100);
            // v0.7.3 #10: content_score 실제 계산 (이전: self=100/else=80 placeholder)
            // get_document_text 양쪽 호출 → 단어 set Jaccard 유사도 × 100
            let content_score = result_path === template_path ? 100 : 0;
            if (result_path !== template_path) {
                try {
                    const tText = await bridge.send('get_document_text', { file_path: template_path, max_chars: 50000 }, ANALYSIS_TIMEOUT);
                    const rText = await bridge.send('get_document_text', { file_path: result_path, max_chars: 50000 }, ANALYSIS_TIMEOUT);
                    const tStr = String(tText.data?.text || '').trim();
                    const rStr = String(rText.data?.text || '').trim();
                    if (tStr && rStr) {
                        const tokenize = (s) => new Set(s.split(/\s+|[.,!?;:()\[\]{}<>'"`~\-_=+|\\\/]+/).filter(w => w.length >= 2));
                        const tSet = tokenize(tStr);
                        const rSet = tokenize(rStr);
                        const union = new Set([...tSet, ...rSet]);
                        const inter = [...tSet].filter(w => rSet.has(w));
                        const jaccard = union.size === 0 ? 0 : inter.length / union.size;
                        content_score = Math.round(jaccard * 100);
                    }
                }
                catch {
                    content_score = 0;
                }
            }
            const overall_score = Math.round(format_score * w.format + structure_score * w.structure + content_score * w.content);
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true,
                            format_score, structure_score, content_score, overall_score,
                            missing_sections, extra_sections, format_deviations,
                            weight: w,
                            self_compare: result_path === template_path,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // 2d) hwp_get_progress — session_state 폴링
    server.tool('hwp_get_progress', '진행 중인 long-running 작업의 진행률 조회. (v0.7.2.3 신규) hwp_session_state(load)의 단축 wrapper. progress_percent/current_section/cancelled 즉시 반환.', {
        session_id: z.string().describe('조회할 session_id'),
    }, async ({ session_id }) => {
        try {
            const sid = safeId(session_id);
            const filePath = path.join(STATE_DIR, `${sid}.json`);
            if (!fs.existsSync(filePath)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `session not found: ${session_id}` }) }], isError: true };
            }
            const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
            const total = (raw.sections_total || []).length;
            const done = (raw.sections_done || []).length;
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true,
                            session_id,
                            progress_percent: typeof raw.progress_percent === 'number' ? raw.progress_percent : (total ? Math.round((done / total) * 100) : 0),
                            current_section: raw.current_section || null,
                            sections_done: done,
                            sections_total: total,
                            cancelled: raw.cancelled || false,
                            last_saved: raw.last_saved || null,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ===== v0.7.4.0: LLM Plan Wrapper =====
    // hwp_autopilot_plan — autopilot 파이프라인 실행 전 구조화된 계획을 반환.
    // LLM 은 직접 호출하지 않음 — estimate_workload + extract_template_structure + analyze_writing_patterns 를
    // 조합해 skeleton(섹션/테이블/스타일/추정치)을 session_state 에 저장.
    // Claude host 가 sections[].content 를 채운 뒤 동일 session_id 로 hwp_autopilot_create 호출.
    server.tool('hwp_autopilot_plan', '문서 자동 생성을 위한 구조화된 계획을 반환합니다. (v0.7.4.0 신규) estimate_workload + extract_template_structure + analyze_writing_patterns 를 조합해 sections[]/tables[]/style_profile skeleton 을 만들고 ~/.hwp_studio_state/{session_id}.json 에 저장. 이 도구는 LLM 을 호출하지 않습니다 — Claude host 가 반환된 plan.sections[].content 를 채운 뒤 동일 session_id 로 hwp_autopilot_create(plan_session_id=...) 를 호출해 실행합니다.', {
        user_request: z.string().describe('사용자 요청 (예: "AI 스타트업 사업계획서 10섹션, A4, 격식체")'),
        template_path: z.string().optional().describe('양식 파일 .hwp/.hwpx (있으면 sections/style_profile 자동 추출)'),
        template_id: z.string().optional().describe('hwp_template_library 등록 ID (template_path 대신)'),
        target_pages: z.number().int().min(1).max(200).optional().describe('목표 페이지 수 (없으면 estimate 에서 유추)'),
        target_sections: z.number().int().min(1).max(50).optional().describe('목표 섹션 수 (없으면 template/estimate 에서 유추)'),
        target_tables: z.number().int().min(0).max(30).optional().describe('목표 표 개수 (기본 0)'),
        reference_files: z.array(z.string()).optional().describe('참고 자료 경로 (estimate_workload 입력)'),
        output_path: z.string().optional().describe('최종 저장 경로 — autopilot_create 에 그대로 전달됨'),
        session_id: z.string().optional().describe('세션 ID (생략 시 자동 생성)'),
    }, async (args) => {
        const sid = safeId(args.session_id || `plan_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
        const sessionPath = path.join(STATE_DIR, `${sid}.json`);
        if (!fs.existsSync(STATE_DIR))
            fs.mkdirSync(STATE_DIR, { recursive: true });
        try {
            // Resolve template_id → file path
            let resolvedTemplate = args.template_path;
            if (!resolvedTemplate && args.template_id) {
                const tid = safeId(args.template_id);
                const tplFile = path.join(TEMPLATE_DIR, 'files', `${tid}.hwpx`);
                if (fs.existsSync(tplFile))
                    resolvedTemplate = tplFile;
            }
            await bridge.ensureRunning();
            // Phase 1: estimate_workload (template 없이도 작동)
            let estimate = {};
            try {
                const estParams = { user_request: args.user_request };
                if (resolvedTemplate)
                    estParams.file_path = resolvedTemplate;
                if (args.reference_files)
                    estParams.reference_files = args.reference_files;
                const r = await bridge.send('estimate_workload', estParams, ANALYSIS_TIMEOUT);
                if (r.success && r.data)
                    estimate = r.data;
            }
            catch { }
            // Phase 2: template structure + writing patterns (template 있을 때만)
            let templateStructure = null;
            let writingPatterns = null;
            if (resolvedTemplate) {
                try {
                    const r = await bridge.send('extract_template_structure', { file_path: resolvedTemplate }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data)
                        templateStructure = r.data;
                }
                catch { }
                try {
                    const r = await bridge.send('analyze_writing_patterns', { file_path: resolvedTemplate }, ANALYSIS_TIMEOUT);
                    if (r.success && r.data)
                        writingPatterns = r.data;
                }
                catch { }
            }
            // Phase 3: derive section skeleton
            const tplSections = templateStructure?.sections || [];
            const targetSec = args.target_sections
                ?? (tplSections.length > 0 ? tplSections.length : (typeof estimate.sections_estimate === 'number' ? estimate.sections_estimate : 5));
            const targetTbl = args.target_tables
                ?? (typeof estimate.tables_estimate === 'number' ? estimate.tables_estimate : 0);
            const targetPgs = args.target_pages
                ?? (typeof estimate.pages_estimate === 'number' ? estimate.pages_estimate : 10);
            // Build sections[] — template 있으면 title 재사용, 없으면 placeholder
            const sectionsSkel = tplSections.length > 0
                ? tplSections.slice(0, targetSec).map((s, i) => ({
                    id: `sec_${i + 1}`,
                    title: String(s.title || `섹션 ${i + 1}`),
                    outline_level: typeof s.level === 'number' ? s.level : 1,
                    content: '',
                    target_chars: Math.round((targetPgs * 1100) / Math.max(1, targetSec)),
                }))
                : Array.from({ length: targetSec }, (_unused, i) => ({
                    id: `sec_${i + 1}`,
                    title: `섹션 ${i + 1}`,
                    outline_level: 1,
                    content: '',
                    target_chars: Math.round((targetPgs * 1100) / Math.max(1, targetSec)),
                }));
            // Build tables[] skeleton (Claude host 가 data[]를 채움)
            const tablesSkel = Array.from({ length: targetTbl }, (_unused, i) => ({
                id: `tbl_${i + 1}`,
                caption: '',
                suggested_headers: [],
                rows_hint: 5,
                cols_hint: 3,
                data: [],
            }));
            // Phase 4: extract style_profile from writing patterns
            const styleProfileOut = writingPatterns ? {
                body_style: writingPatterns.body_style,
                title_styles: writingPatterns.title_styles,
                table_styles: writingPatterns.table_styles,
                page_setup: writingPatterns.page_setup,
                numbering_pattern: writingPatterns.numbering_pattern,
            } : null;
            // Persist plan — hwp_autopilot_create 가 plan_session_id 로 hydrate
            const planObject = {
                plan_version: '0.7.4.0',
                session_id: sid,
                user_request: args.user_request,
                template_path: resolvedTemplate || null,
                output_path: args.output_path || null,
                target_pages: targetPgs,
                target_sections: targetSec,
                target_tables: targetTbl,
                sections: sectionsSkel,
                tables: tablesSkel,
                style_profile: styleProfileOut,
                estimate,
                template_structure: templateStructure,
                next_action: {
                    tool: 'hwp_autopilot_create',
                    required_fill: ['sections[].content', 'tables[].data', 'tables[].caption'],
                    guidance: 'plan.sections 와 plan.tables 에 콘텐츠를 채운 후 동일 session_id 또는 plan_session_id 로 hwp_autopilot_create 호출',
                },
                created_at: new Date().toISOString(),
                status: 'awaiting_content',
            };
            fs.writeFileSync(sessionPath, JSON.stringify(planObject, null, 2), 'utf8');
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true, mode: 'plan', plan: planObject,
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ===== v0.7.2.4: End-to-End Autopilot =====
    // hwp_autopilot_create — 12-step pipeline orchestrator (사용자 핵심 니즈)
    // 콘텐츠 생성은 호출자(LLM)가 미리 sections[]에 담아 전달.
    server.tool('hwp_autopilot_create', '문서 자동 생성 12단계 파이프라인. (v0.7.2.4 / v0.7.4.0 / v0.7.5.4) sections[]/tables[]를 받아 template 기반으로 작성→스타일→TOC→검증→저장→PDF→layout 검증. v0.7.5.4: auto_fix 기본값 false + preserve_template_style 기본값 true (원본 템플릿 서식 보존). mode=plan은 estimate만, mode=execute는 실제 파이프. 각 step 마다 session_state 자동 save + cancel 체크.', {
        output_path: z.string().optional().describe('생성할 .hwp/.hwpx 절대 경로 (plan_session_id 사용 시 생략 가능)'),
        sections: z.array(z.object({
            title: z.string(),
            content: z.string(),
            outline_level: z.number().int().min(1).max(7).optional(),
        })).optional().describe('미리 생성된 섹션 콘텐츠 (호출자가 LLM으로 작성). plan_session_id 사용 시 생략 가능.'),
        template_path: z.string().optional().describe('템플릿 .hwpx 경로 (없으면 빈 문서로 시작)'),
        template_id: z.string().optional().describe('hwp_template_library 등록 ID (template_path 대신)'),
        tables: z.array(z.object({
            caption: z.string().optional(),
            data: z.array(z.array(z.string())),
        })).optional().describe('삽입할 표 목록 (sections 다음에 일괄 삽입)'),
        style_profile: z.any().optional().describe('apply_style_profile에 전달할 프로파일'),
        approve_threshold_seconds: z.number().optional().describe('estimate가 이를 넘으면 awaiting_approval 반환 (기본 600)'),
        export_pdf: z.boolean().optional().describe('완료 후 PDF 변환 (기본 true)'),
        mode: z.enum(['plan', 'execute']).optional().describe('plan=estimate만, execute=실제 실행 (기본 execute)'),
        session_id: z.string().optional().describe('재개할 세션 (생략 시 자동 생성)'),
        prompt: z.string().optional().describe('원본 사용자 프롬프트 (메타데이터)'),
        // v0.7.5.4 P2-2: auto_fix 기본값 false (원본 서식 override 방지)
        auto_fix: z.boolean().optional().describe('v0.7.5.4: 기본 false. validate_consistency 점수 < threshold 여도 자동 override 안 함. true 여도 runAutoFixLoop 가 P0-2 에서 no-op 전환됨 (validate 만 수행).'),
        auto_fix_threshold: z.number().min(0).max(100).optional().describe('auto_fix 트리거 점수 (기본 85, auto_fix=true 일 때만 의미)'),
        auto_fix_max_iterations: z.number().int().min(1).max(5).optional().describe('auto_fix 최대 반복 횟수 (기본 2, 안전 한도 5)'),
        // v0.7.5.4 P2-2: 원본 템플릿 서식 보존 (기본 true)
        preserve_template_style: z.boolean().optional().describe('v0.7.5.4: 원본 템플릿의 서식 (heading/body/cell) 을 변경하지 않음. true 시 style_profile override 무시. 공무원 양식 작업 시 권장 (기본 true).'),
        plan_session_id: z.string().optional().describe('hwp_autopilot_plan 이 만든 plan session_id. 지정 시 sections/tables/style_profile/output_path/template_path 를 해당 세션에서 자동 로드 (개별 인자 우선)'),
    }, async (args) => {
        const startTime = Date.now();
        const mode = args.mode || 'execute';
        const threshold = args.approve_threshold_seconds ?? 600;
        const exportPdf = args.export_pdf ?? true;
        const sid = safeId(args.session_id || `auto_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
        const sessionPath = path.join(STATE_DIR, `${sid}.json`);
        if (!fs.existsSync(STATE_DIR))
            fs.mkdirSync(STATE_DIR, { recursive: true });
        // v0.7.5.4 P2-2: auto_fix 기본값 false (원본 서식 override 방지)
        // preserve_template_style=true (기본) 면 style_profile 무시 + auto_fix 강제 off
        const preserveTemplateStyle = args['preserve_template_style'] ?? true;
        const autoFix = preserveTemplateStyle ? false : (args.auto_fix ?? false);
        const autoFixThreshold = args.auto_fix_threshold ?? 85;
        const autoFixMaxIter = Math.min(args.auto_fix_max_iterations ?? 2, 5);
        // v0.7.4.0: plan_session_id hydration — hwp_autopilot_plan 세션에서 sections/tables/style_profile/output_path/template_path 자동 로드
        // NOTE: Zod schema field names have snake_case but we already replaced args.sections→sections etc.
        // Read directly from the args object via bracket notation to avoid self-reference.
        const rawArgs = args;
        let hSections = rawArgs['sections'];
        let hTables = rawArgs['tables'];
        let hStyleProfile = rawArgs['style_profile'];
        let hOutputPath = rawArgs['output_path'];
        let hTemplatePath = rawArgs['template_path'];
        const templateId = rawArgs['template_id'];
        if (args.plan_session_id) {
            try {
                const planPath = path.join(STATE_DIR, `${safeId(args.plan_session_id)}.json`);
                if (fs.existsSync(planPath)) {
                    const plan = JSON.parse(fs.readFileSync(planPath, 'utf8'));
                    if ((!hSections || hSections.length === 0) && Array.isArray(plan.sections)) {
                        hSections = plan.sections
                            .filter((s) => typeof s.content === 'string' && s.content.trim().length > 0)
                            .map((s) => ({
                            title: String(s.title || ''),
                            content: String(s.content || ''),
                            outline_level: typeof s.outline_level === 'number' ? s.outline_level : undefined,
                        }));
                    }
                    if ((!hTables || hTables.length === 0) && Array.isArray(plan.tables)) {
                        hTables = plan.tables
                            .filter((t) => Array.isArray(t.data) && t.data.length > 0)
                            .map((t) => ({ caption: t.caption, data: t.data }));
                    }
                    if (!hStyleProfile && plan.style_profile)
                        hStyleProfile = plan.style_profile;
                    if (!hOutputPath && plan.output_path)
                        hOutputPath = String(plan.output_path);
                    if (!hTemplatePath && plan.template_path)
                        hTemplatePath = String(plan.template_path);
                }
            }
            catch { }
        }
        if (!hSections || hSections.length === 0) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: 'sections 가 비어있습니다. sections[]를 직접 전달하거나 plan_session_id 로 hwp_autopilot_plan 결과를 참조하세요.',
                        }) }], isError: true };
        }
        if (!hOutputPath) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: 'output_path 가 비어있습니다. 직접 전달하거나 plan_session_id 에 포함시켜 주세요.',
                        }) }], isError: true };
        }
        // Non-null re-bind so body code can reuse familiar names
        const sections = hSections;
        const tables = hTables;
        const styleProfile = hStyleProfile;
        const outputPath = hOutputPath;
        const templatePath = hTemplatePath;
        const sectionTitles = sections.map(s => s.title);
        const totalSteps = 8 + sections.length + (tables?.length || 0);
        let stepsDone = 0;
        const stepLog = [];
        const saveSession = (extra = {}) => {
            let existing = {};
            if (fs.existsSync(sessionPath)) {
                try {
                    existing = JSON.parse(fs.readFileSync(sessionPath, 'utf8'));
                }
                catch { }
            }
            const merged = {
                ...existing,
                session_id: sid,
                current_doc: outputPath,
                sections_total: sectionTitles,
                progress_percent: Math.round((stepsDone / totalSteps) * 100),
                last_saved: new Date().toISOString(),
                ...extra,
            };
            fs.writeFileSync(sessionPath, JSON.stringify(merged, null, 2), 'utf8');
            return merged;
        };
        const isCancelled = () => {
            if (!fs.existsSync(sessionPath))
                return false;
            try {
                return !!JSON.parse(fs.readFileSync(sessionPath, 'utf8')).cancelled;
            }
            catch {
                return false;
            }
        };
        const recordStep = (name, ok, detail) => {
            stepsDone++;
            stepLog.push({ step: stepsDone, name, ok, detail });
            saveSession({ current_section: name });
        };
        try {
            await bridge.ensureRunning();
            // STEP 1: estimate (plan/execute 공통) — v0.7.2.5: Python contract에 맞춘 user_request 전달
            let estimate = {};
            try {
                const userRequest = args.prompt || `${sections.length}개 섹션 자동 생성`;
                const estParams = { user_request: userRequest };
                if (templatePath)
                    estParams.file_path = templatePath;
                const r = await bridge.send('estimate_workload', estParams, ANALYSIS_TIMEOUT);
                if (r.success)
                    estimate = r.data;
            }
            catch { }
            const estSeconds = typeof estimate.duration_seconds_estimate === 'number'
                ? estimate.duration_seconds_estimate
                : 0;
            if (mode === 'plan') {
                return { content: [{ type: 'text', text: JSON.stringify({
                                ok: true, mode: 'plan', session_id: sid, estimate,
                                steps_total: totalSteps, sections: sectionTitles,
                                tables: tables?.length || 0,
                                requires_approval: estSeconds > threshold,
                            }) }] };
            }
            if (estSeconds > threshold) {
                saveSession({ status: 'awaiting_approval' });
                return { content: [{ type: 'text', text: JSON.stringify({
                                ok: true, status: 'awaiting_approval', session_id: sid, estimate,
                                message: `예상 ${estSeconds}초 > 임계 ${threshold}초. session_id로 재호출 시 approve_threshold_seconds를 늘려 진행하세요.`,
                            }) }] };
            }
            // STEP 2: 문서 생성 또는 템플릿 열기
            if (templatePath) {
                const r = await bridge.send('open_document', { file_path: templatePath }, ANALYSIS_TIMEOUT);
                recordStep('open_template', r.success, r.error);
                if (!r.success)
                    throw new Error(`open_template failed: ${r.error}`);
            }
            else if (templateId) {
                const safeTid = safeId(templateId);
                const tplFile = path.join(TEMPLATE_DIR, 'files', `${safeTid}.hwpx`);
                if (!fs.existsSync(tplFile))
                    throw new Error(`template file not found: ${safeTid}`);
                const r = await bridge.send('open_document', { file_path: tplFile }, ANALYSIS_TIMEOUT);
                recordStep('open_template_library', r.success, { template_id: safeTid });
                if (!r.success)
                    throw new Error(`open template_id failed: ${r.error}`);
            }
            else {
                // v0.7.2.5: document_create → document_new (Python에 신규 추가)
                const r = await bridge.send('document_new', {}, ANALYSIS_TIMEOUT);
                recordStep('document_new', r.success, r.error);
                if (!r.success)
                    throw new Error(`document_new failed: ${r.error}`);
            }
            if (isCancelled())
                throw new Error('cancelled');
            // STEP 3..N: 섹션 작성 루프 (v0.7.2.9: 본문 길이 cross-check)
            let lastBodyChars = 0;
            for (const section of sections) {
                if (isCancelled())
                    throw new Error('cancelled');
                const headingR = await bridge.send('insert_heading', {
                    text: section.title,
                    outline_level: section.outline_level || 1,
                }, ANALYSIS_TIMEOUT);
                if (!headingR.success) {
                    // fallback: insert_text
                    await bridge.send('insert_text', { text: section.title + '\n' }, ANALYSIS_TIMEOUT);
                }
                const textR = await bridge.send('insert_text', { text: section.content + '\n\n' }, ANALYSIS_TIMEOUT);
                // v0.7.2.9: 본문 검증 — word_count 로 누적 글자 수 측정
                let bodyVerified = false;
                let bodyDelta = 0;
                let currentBodyChars = lastBodyChars;
                try {
                    const wc = await bridge.send('word_count', {}, ANALYSIS_TIMEOUT);
                    if (wc.success && wc.data && typeof wc.data.chars_total === 'number') {
                        currentBodyChars = wc.data.chars_total;
                        bodyDelta = currentBodyChars - lastBodyChars;
                        const expectedDelta = section.title.length + section.content.length;
                        // 50% 이상 들어갔으면 OK (단락 마커/줄바꿈 차이 허용)
                        bodyVerified = bodyDelta >= expectedDelta * 0.5;
                        lastBodyChars = currentBodyChars;
                    }
                }
                catch { }
                recordStep(`section:${section.title}`, textR.success && bodyVerified, {
                    chars: section.content.length,
                    body_chars_after: currentBodyChars,
                    body_delta: bodyDelta,
                    body_verified: bodyVerified,
                });
                if (!bodyVerified) {
                    throw new Error(`section "${section.title}" body verification failed: delta=${bodyDelta} expected~=${section.title.length + section.content.length}. autopilot이 본문을 쓰지 못함 — cursor 위치 확인 필요`);
                }
                saveSession({
                    sections_done: stepLog.filter(s => String(s.name).startsWith('section:')).map(s => String(s.name).slice(8)),
                    current_section: section.title,
                });
            }
            // STEP: 표 삽입
            if (tables) {
                for (let i = 0; i < tables.length; i++) {
                    if (isCancelled())
                        throw new Error('cancelled');
                    const t = tables[i];
                    const r = await bridge.send('table_create_from_data', {
                        data: t.data,
                        caption: t.caption,
                    }, ANALYSIS_TIMEOUT);
                    recordStep(`table:${i}`, r.success, { rows: t.data.length, caption: t.caption });
                }
            }
            // STEP: 스타일 프로파일 적용
            if (styleProfile) {
                const r = await bridge.send('apply_style_profile', { profile: styleProfile }, ANALYSIS_TIMEOUT);
                recordStep('apply_style_profile', r.success, r.error);
            }
            // STEP: TOC — v0.7.2.9: outline_level 있는 섹션이 1개 이상일 때만 호출
            const hasOutline = sections.some(s => typeof s.outline_level === 'number' && s.outline_level >= 1);
            if (hasOutline) {
                try {
                    const r = await bridge.send('generate_toc', {}, ANALYSIS_TIMEOUT);
                    recordStep('generate_toc', r.success, r.error);
                }
                catch (e) {
                    recordStep('generate_toc', false, e.message);
                }
            }
            else {
                recordStep('generate_toc', true, { skipped: true, reason: 'no outline_level sections' });
            }
            if (isCancelled())
                throw new Error('cancelled');
            // STEP: 저장 (v0.7.2.7: save_document는 현재경로만 저장 → 새 문서는 save_as 필수)
            const saveFmt = outputPath.toLowerCase().endsWith('.hwpx') ? 'HWPX' : 'HWP';
            const saveR = await bridge.send('save_as', { path: outputPath, format: saveFmt }, ANALYSIS_TIMEOUT);
            recordStep('save_as', saveR.success, saveR.error);
            if (!saveR.success)
                throw new Error(`save_as failed: ${saveR.error}`);
            // v0.7.2.11: 파일 사이즈 하한선 검증 — body_verified 가 primary, size 는 secondary
            // 이전 22KB 는 짧은 본문(150자)에서 false positive. body_verified 가 이미 본문 존재를 증명.
            // 하한선은 19KB (빈 HWPX ~18KB + 마진 1KB). body_verified 가 모두 true 면 size 미달이어도 pass.
            try {
                const stat = fs.statSync(outputPath);
                const minBytes = outputPath.toLowerCase().endsWith('.hwpx') ? 19000 : 24000;
                const sizeOk = stat.size >= minBytes;
                const allSectionsVerified = stepLog
                    .filter(s => String(s.name).startsWith('section:'))
                    .every(s => (s.detail)?.body_verified === true);
                if (!sizeOk && !allSectionsVerified) {
                    // primary + secondary 모두 실패 = 진짜 빈 파일
                    recordStep('file_size_check', false, { size: stat.size, min_required: minBytes });
                    throw new Error(`saved file size ${stat.size} < ${minBytes} bytes and body not verified — likely empty HWPX`);
                }
                recordStep('file_size_check', true, {
                    size: stat.size,
                    min_required: minBytes,
                    primary_body_verified: allSectionsVerified,
                    note: sizeOk ? 'size ok' : 'under threshold but body_verified primary pass',
                });
            }
            catch (e) {
                if (e.message.includes('not verified'))
                    throw e;
                recordStep('file_size_check', false, e.message);
            }
            // STEP: validate_consistency + auto_fix loop (v0.7.4.0 — 실구현)
            // runAutoFixLoop 이 내부적으로 validate 를 수행하고, score < threshold 면
            // apply_style_profile → save_as → word_count cross-check → re-validate 를 반복.
            // autoFix=false 면 threshold=0 으로 넘겨 단순 validate 만 수행.
            let score = 100;
            let scoreBefore = null;
            let autoFixIterations = 0;
            let autoFixLog = [];
            let autoFixStopped = 'skipped';
            try {
                const loopResult = await runAutoFixLoop({
                    outputPath,
                    styleProfile: autoFix ? styleProfile : undefined,
                    threshold: autoFix ? autoFixThreshold : 0,
                    maxIter: autoFixMaxIter,
                    saveFmt,
                    isCancelled,
                    onIteration: (entry) => {
                        saveSession({ auto_fix_last_iteration: entry });
                    },
                });
                scoreBefore = loopResult.score_before;
                score = loopResult.score_after;
                autoFixIterations = loopResult.iterations;
                autoFixLog = loopResult.log;
                autoFixStopped = loopResult.stopped_reason;
                recordStep('validate_consistency', true, { score, score_before: scoreBefore });
                if (autoFixIterations > 0) {
                    const loopOk = autoFixStopped === 'threshold_reached'
                        || autoFixStopped === 'max_iterations'
                        || autoFixStopped === 'plateau';
                    recordStep(`auto_fix:${autoFixIterations}`, loopOk, {
                        iterations: autoFixIterations,
                        stopped_reason: autoFixStopped,
                        score_before: scoreBefore,
                        score_after: score,
                    });
                }
                saveSession({ auto_fix_log: autoFixLog });
            }
            catch (e) {
                recordStep('validate_consistency', false, e.message);
            }
            // STEP: PDF — v0.7.2.5: export_pdf → export_format(format:PDF)
            let pdf_path = null;
            if (exportPdf) {
                const pdfTarget = outputPath.replace(/\.(hwp|hwpx)$/i, '.pdf');
                const r = await bridge.send('export_format', { path: pdfTarget, format: 'PDF' }, ANALYSIS_TIMEOUT);
                recordStep('export_format', r.success, r.error);
                if (r.success)
                    pdf_path = pdfTarget;
            }
            // STEP: verify_layout
            try {
                const r = await bridge.send('verify_layout', { file_path: outputPath }, ANALYSIS_TIMEOUT);
                recordStep('verify_layout', r.success, r.error);
            }
            catch (e) {
                recordStep('verify_layout', false, e.message);
            }
            const finalSession = saveSession({
                status: 'completed',
                progress_percent: 100,
            });
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: true, status: 'completed', session_id: sid,
                            saved_path: outputPath, pdf_path,
                            score,
                            score_before: scoreBefore,
                            auto_fix_iterations: autoFixIterations,
                            auto_fix_log: autoFixLog,
                            auto_fix_stopped_reason: autoFixStopped,
                            steps_done: stepsDone, steps_total: totalSteps,
                            duration_seconds: Math.round((Date.now() - startTime) / 1000),
                            step_log: stepLog,
                            estimate,
                        }) }] };
        }
        catch (err) {
            const message = err.message;
            const cancelled = message === 'cancelled';
            saveSession({ status: cancelled ? 'cancelled' : 'failed', error: message });
            return { content: [{ type: 'text', text: JSON.stringify({
                            ok: false, status: cancelled ? 'cancelled' : 'failed',
                            session_id: sid, error: message,
                            steps_done: stepsDone, steps_total: totalSteps,
                            step_log: stepLog,
                            duration_seconds: Math.round((Date.now() - startTime) / 1000),
                        }) }], isError: !cancelled };
        }
    });
    // ===== v0.7.4.2: PDF OCR → HWP Clone =====
    // PDF (native 선택 텍스트 또는 스캔 한국어) → 편집 가능한 HWP/HWPX 복원.
    // v0.7.4.2: native PDF only (pdfplumber + PyMuPDF get_text dict → 순차 단락)
    // v0.7.4.3: + PaddleOCR 스캔 지원, 전처리, 제목 감지, hybrid per-page dispatch
    // v0.7.4.4: + 표/이미지 재구성, fidelity scoring, column 감지 경고
    const PDF_CLONE_TIMEOUT = 600000; // 10분 — OCR 페이지 수에 따라 가변
    server.tool('hwp_pdf_clone', 'PDF (native 또는 스캔 한국어) 를 편집 가능한 HWP/HWPX 로 복원합니다. (v0.7.4.4) native PDF 는 PyMuPDF get_text("dict") 로 bbox + 폰트 직접 추출, 스캔 PDF 는 PaddleOCR (lang=korean, ~150MB 모델 최초 자동 다운로드) + opencv 전처리 (deskew + denoise + threshold). hybrid PDF 는 페이지별 자동 dispatch. 제목 감지, 표 재구성 (pdfplumber find_tables), 이미지 임베딩 (page.get_images + extract_image), 2-column 감지 경고, 4-component fidelity score (text/page/layout/structure). 출력은 원본과 시각적으로 유사한 클론 (픽셀 단위 일치 아님).', {
        pdf_path: z.string().describe('원본 PDF 경로 (절대 또는 상대)'),
        output_path: z.string().describe('출력 HWP/HWPX 경로 (.hwp 또는 .hwpx, 확장자에 따라 형식 결정)'),
        options: z.object({
            ocr_engine: z.enum(['paddle', 'none', 'auto']).optional()
                .describe('"auto"(기본): native PDF 는 OCR 미사용, 스캔 PDF 만 PaddleOCR (v0.7.4.3+). "none": 강제 native 만. "paddle": 강제 OCR.'),
            preserve_images: z.boolean().optional()
                .describe('PDF 내 이미지를 추출해 HWP 에 삽입 (기본 true, v0.7.4.4 부터 유효)'),
            detect_tables: z.boolean().optional()
                .describe('표 구조 자동 감지 + table_create_from_data 호출 (기본 true, v0.7.4.4 부터 유효)'),
            max_pages: z.number().int().min(0).optional()
                .describe('처리할 최대 페이지 수 (0 = 전체, 기본 0). 큰 PDF 테스트 시 1-3 권장'),
            preprocess: z.boolean().optional()
                .describe('스캔 이미지 전처리 (deskew + denoise + threshold). 기본 true, v0.7.4.3 부터 유효'),
            min_native_chars_per_page: z.number().int().min(0).optional()
                .describe('native PDF 로 분류하는 페이지당 최소 글자수 (기본 30)'),
            page_setup_from_pdf: z.boolean().optional()
                .describe('PDF 첫 페이지의 가로/세로 mm 를 HWP 용지 설정에 적용 (기본 true)'),
            lang: z.string().optional().describe('OCR 언어 (기본 "korean", v0.7.4.3 부터 유효)'),
        }).optional().describe('PDF clone 옵션'),
    }, async ({ pdf_path, output_path, options }) => {
        const pdfResolved = path.resolve(pdf_path);
        const outResolved = path.resolve(output_path);
        if (!fs.existsSync(pdfResolved)) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: `PDF 파일을 찾을 수 없습니다: ${pdfResolved}`,
                        }) }], isError: true };
        }
        const outExt = path.extname(outResolved).toLowerCase();
        if (outExt !== '.hwp' && outExt !== '.hwpx') {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: 'output_path 는 .hwp 또는 .hwpx 확장자여야 합니다.',
                        }) }], isError: true };
        }
        const outDir = path.dirname(outResolved);
        if (!fs.existsSync(outDir)) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: `출력 디렉토리가 존재하지 않습니다: ${outDir}`,
                        }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const startedAt = Date.now();
            const response = await bridge.send('clone_pdf_to_hwp', {
                pdf_path: pdfResolved,
                output_path: outResolved,
                options: options ?? {},
            }, PDF_CLONE_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: response.error,
                            }) }], isError: true };
            }
            const data = (response.data || {});
            // Success → 현재 문서를 새 출력으로 갱신 (hwp_save_document / hwp_export_pdf 체이닝 가능)
            if ((data.status === 'ok' || data.status === 'partial') && fs.existsSync(outResolved)) {
                bridge.setCurrentDocument(outResolved);
                bridge.setCachedAnalysis(null);
            }
            return { content: [{ type: 'text', text: JSON.stringify({
                            ...data,
                            elapsed_seconds: Math.round((Date.now() - startedAt) / 1000),
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: err.message,
                        }) }], isError: true };
        }
    });
}
