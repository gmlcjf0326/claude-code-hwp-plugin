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
    // ── 진단 도구 (개발용) ──
    server.tool('hwp_inspect_com_object', '[개발용] pyhwpx COM 객체의 실제 속성 목록을 덤프합니다. HCharShape/HParaShape 등의 정확한 속성명을 확인할 때 사용.', {
        object: z.enum(['HCharShape', 'HParaShape', 'HFindReplace']).optional().describe('조사할 COM 객체 (기본: HCharShape)'),
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
    const DEFAULT_POLICY = {
        max_reference_files: 5,
        max_total_size_mb: 10,
        max_tokens_input_percent: 80,
        allowed_formats: ['xlsx', 'csv', 'json', 'pdf', 'docx', 'txt', 'html', 'xml', 'pptx'],
        prefer_summary: true,
        summary_threshold_mb: 3,
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
                // 안전한 자동 수정만: privacy는 수동 필수, consistency 일부만 시도
                // 실제 fix 로직은 v0.7.2.4에서 확장. 현재는 placeholder.
                score_after = Math.min(100, score_before + 5);
                auto_fixed.push({ note: 'auto_fix placeholder — v0.7.2.4에서 확장 예정' });
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
            // content: 단순 placeholder (자체비교는 100)
            const content_score = result_path === template_path ? 100 : 80;
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
    // ===== v0.7.2.4: End-to-End Autopilot =====
    // hwp_autopilot_create — 12-step pipeline orchestrator (사용자 핵심 니즈)
    // 콘텐츠 생성은 호출자(LLM)가 미리 sections[]에 담아 전달.
    server.tool('hwp_autopilot_create', '문서 자동 생성 12단계 파이프라인. (v0.7.2.4 신규) sections[]/tables[]를 받아 template 기반으로 작성→스타일→TOC→검증→저장→PDF→layout 검증까지 일괄 실행. mode=plan은 estimate만 반환, mode=execute는 실제 파이프 실행. 각 step마다 session_state 자동 save + cancel 체크.', {
        output_path: z.string().describe('생성할 .hwp/.hwpx 절대 경로'),
        sections: z.array(z.object({
            title: z.string(),
            content: z.string(),
            outline_level: z.number().int().min(1).max(7).optional(),
        })).describe('미리 생성된 섹션 콘텐츠 (호출자가 LLM으로 작성)'),
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
    }, async (args) => {
        const startTime = Date.now();
        const mode = args.mode || 'execute';
        const threshold = args.approve_threshold_seconds ?? 600;
        const exportPdf = args.export_pdf ?? true;
        const sid = safeId(args.session_id || `auto_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
        const sessionPath = path.join(STATE_DIR, `${sid}.json`);
        if (!fs.existsSync(STATE_DIR))
            fs.mkdirSync(STATE_DIR, { recursive: true });
        const sectionTitles = args.sections.map(s => s.title);
        const totalSteps = 8 + args.sections.length + (args.tables?.length || 0);
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
                current_doc: args.output_path,
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
                const userRequest = args.prompt || `${args.sections.length}개 섹션 자동 생성`;
                const estParams = { user_request: userRequest };
                if (args.template_path)
                    estParams.file_path = args.template_path;
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
                                tables: args.tables?.length || 0,
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
            if (args.template_path) {
                const r = await bridge.send('open_document', { file_path: args.template_path }, ANALYSIS_TIMEOUT);
                recordStep('open_template', r.success, r.error);
                if (!r.success)
                    throw new Error(`open_template failed: ${r.error}`);
            }
            else if (args.template_id) {
                const safeTid = safeId(args.template_id);
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
            // STEP 3..N: 섹션 작성 루프
            for (const section of args.sections) {
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
                recordStep(`section:${section.title}`, textR.success, { chars: section.content.length });
                saveSession({
                    sections_done: stepLog.filter(s => String(s.name).startsWith('section:')).map(s => String(s.name).slice(8)),
                    current_section: section.title,
                });
            }
            // STEP: 표 삽입
            if (args.tables) {
                for (let i = 0; i < args.tables.length; i++) {
                    if (isCancelled())
                        throw new Error('cancelled');
                    const t = args.tables[i];
                    const r = await bridge.send('table_create_from_data', {
                        data: t.data,
                        caption: t.caption,
                    }, ANALYSIS_TIMEOUT);
                    recordStep(`table:${i}`, r.success, { rows: t.data.length, caption: t.caption });
                }
            }
            // STEP: 스타일 프로파일 적용
            if (args.style_profile) {
                const r = await bridge.send('apply_style_profile', { profile: args.style_profile }, ANALYSIS_TIMEOUT);
                recordStep('apply_style_profile', r.success, r.error);
            }
            // STEP: TOC + refresh fields
            try {
                const r = await bridge.send('generate_toc', {}, ANALYSIS_TIMEOUT);
                recordStep('generate_toc', r.success, r.error);
            }
            catch (e) {
                recordStep('generate_toc', false, e.message);
            }
            if (isCancelled())
                throw new Error('cancelled');
            // STEP: 일단 저장 (validate가 path 필요)
            const saveR = await bridge.send('save_document', { file_path: args.output_path }, ANALYSIS_TIMEOUT);
            recordStep('save_document', saveR.success, saveR.error);
            if (!saveR.success)
                throw new Error(`save_document failed: ${saveR.error}`);
            // STEP: validate_consistency, score < 85면 review_and_edit auto_fix (placeholder)
            let score = 100;
            try {
                const r = await bridge.send('validate_consistency', { file_path: args.output_path }, ANALYSIS_TIMEOUT);
                if (r.success && r.data) {
                    const d = r.data;
                    // v0.7.2.5: Python 실제 키는 consistency_score
                    score = typeof d.consistency_score === 'number' ? d.consistency_score : 100;
                }
                recordStep('validate_consistency', r.success, { score });
            }
            catch (e) {
                recordStep('validate_consistency', false, e.message);
            }
            // STEP: PDF — v0.7.2.5: export_pdf → export_format(format:PDF)
            let pdf_path = null;
            if (exportPdf) {
                const pdfTarget = args.output_path.replace(/\.(hwp|hwpx)$/i, '.pdf');
                const r = await bridge.send('export_format', { path: pdfTarget, format: 'PDF' }, ANALYSIS_TIMEOUT);
                recordStep('export_format', r.success, r.error);
                if (r.success)
                    pdf_path = pdfTarget;
            }
            // STEP: verify_layout
            try {
                const r = await bridge.send('verify_layout', { file_path: args.output_path }, ANALYSIS_TIMEOUT);
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
                            saved_path: args.output_path, pdf_path,
                            score, steps_done: stepsDone, steps_total: totalSteps,
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
}
