/**
 * Analysis tools: analyze, get text, get tables, get fields
 * HWPX 파일은 XML 직접 검색으로 라우팅 (COM 우회)
 */
import { z } from 'zod';
import path from 'node:path';
import fs from 'node:fs';
const HWP_EXTENSIONS = new Set(['.hwp', '.hwpx']);
const ANALYSIS_TIMEOUT = 60000;
async function ensureAnalysis(bridge, filePath) {
    await bridge.ensureRunning();
    if (filePath) {
        const resolved = path.resolve(filePath);
        if (!fs.existsSync(resolved)) {
            throw new Error(`파일을 찾을 수 없습니다: ${resolved}`);
        }
        const ext = path.extname(resolved).toLowerCase();
        if (!HWP_EXTENSIONS.has(ext)) {
            throw new Error('HWP 또는 HWPX 파일만 지원합니다.');
        }
        // COM 메모리 변경사항을 파일에 반영 (미저장 표/텍스트 감지용)
        try {
            await bridge.send('save_document', {});
        }
        catch { }
        const response = await bridge.send('analyze_document', { file_path: resolved }, ANALYSIS_TIMEOUT);
        if (!response.success)
            throw new Error(response.error ?? '분석 실패');
        bridge.setCachedAnalysis(response.data);
        bridge.setCurrentDocument(resolved);
        return response.data;
    }
    const cached = bridge.getCachedAnalysis();
    const current = bridge.getCurrentDocument();
    // P1 #8: 캐시가 현재 열린 문서와 일치하는 경우에만 반환
    if (cached && current && cached.file_path === current)
        return cached;
    if (!current) {
        throw new Error('열린 문서가 없습니다. hwp_open_document로 문서를 열거나 file_path를 지정하세요. Python 프로세스 재시작 시 열린 문서 상태가 초기화됩니다.');
    }
    const response = await bridge.send('analyze_document', { file_path: current }, ANALYSIS_TIMEOUT);
    if (!response.success)
        throw new Error(response.error ?? '분석 실패');
    bridge.setCachedAnalysis(response.data);
    return response.data;
}
export function registerAnalysisTools(server, bridge, toolset = 'standard') {
    server.tool('hwp_analyze_document', 'HWP/HWPX 문서의 전체 구조를 분석합니다. 페이지 수, 표(데이터 포함), 필드(양식), 본문 텍스트를 반환합니다. 문서를 처음 다룰 때 반드시 이 도구를 먼저 호출하세요.', {
        file_path: z.string().describe('HWP/HWPX 파일 경로'),
    }, async ({ file_path }) => {
        try {
            const result = await ensureAnalysis(bridge, file_path);
            return { content: [{ type: 'text', text: JSON.stringify(result) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_get_document_text', '현재 열린 문서 또는 지정 파일의 본문 텍스트를 추출합니다. 문서 내용을 읽거나 검색할 때 사용하세요.', {
        file_path: z.string().optional().describe('HWP/HWPX 파일 경로 (생략 시 현재 문서)'),
        max_chars: z.number().optional().describe('최대 문자 수 (기본: 15000)'),
    }, async ({ file_path, max_chars }) => {
        try {
            const analysis = await ensureAnalysis(bridge, file_path);
            const limit = max_chars ?? 15000;
            const text = analysis.full_text || '';
            const truncated = text.length > limit;
            return {
                content: [{
                        type: 'text',
                        text: JSON.stringify({
                            text: truncated ? text.slice(0, limit) : text,
                            char_count: text.length,
                            truncated,
                        }),
                    }],
            };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_get_tables', '문서의 표 데이터를 조회합니다. 특정 표 인덱스를 지정하거나 전체 표를 반환합니다. 표 셀을 채우기 전에 구조를 확인할 때 사용하세요.', {
        file_path: z.string().optional().describe('HWP/HWPX 파일 경로 (생략 시 현재 문서)'),
        table_index: z.number().optional().describe('특정 표 인덱스 (생략 시 전체)'),
    }, async ({ file_path, table_index }) => {
        try {
            const analysis = await ensureAnalysis(bridge, file_path);
            const tables = analysis.tables || [];
            if (table_index !== undefined) {
                if (table_index < 0 || table_index >= tables.length) {
                    return {
                        content: [{ type: 'text', text: JSON.stringify({ error: `표 인덱스 ${table_index}이 범위를 벗어났습니다. (총 ${tables.length}개)` }) }],
                        isError: true,
                    };
                }
                return {
                    content: [{ type: 'text', text: JSON.stringify({ table: tables[table_index], total_count: tables.length }) }],
                };
            }
            return {
                content: [{ type: 'text', text: JSON.stringify({ tables, total_count: tables.length }) }],
            };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_map_table_cells', '표의 셀을 Tab 순서로 순회하여 각 셀의 Tab 인덱스와 내용을 매핑합니다. 병합 셀이 있는 표에서 hwp_fill_table_cells의 tab 파라미터에 사용할 인덱스를 확인할 때 사용하세요.', {
        table_index: z.number().int().min(0).describe('표 인덱스 (0부터 시작)'),
    }, async ({ table_index }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                        }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('map_table_cells', { table_index }, ANALYSIS_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // --- minimal에 포함: get_document_info, text_search ---
    server.tool('hwp_get_document_info', '현재 열린 문서의 경량 메타데이터(페이지 수, 파일 경로)를 빠르게 반환합니다. analyze_document보다 훨씬 빠릅니다. 문서가 열려 있는지, 몇 페이지인지만 빠르게 확인할 때 사용하세요.', {}, async () => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('get_document_info', {}, 10000);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_text_search', '문서에서 텍스트를 검색하고 발견된 위치와 횟수를 반환합니다. 치환 없이 검색만 합니다. 특정 텍스트가 문서에 있는지 확인하거나, 몇 번 등장하는지 파악할 때 사용하세요.', {
        search: z.string().describe('검색할 텍스트'),
        max_results: z.number().int().min(1).optional().describe('최대 검색 결과 수 (기본 50)'),
    }, async ({ search, max_results }) => {
        const filePath = bridge.getCurrentDocument();
        if (!filePath) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.' }) }], isError: true };
        }
        try {
            // COM 우선 검색 (HWP/HWPX 모두 — XML 단절 문제 근본 해결)
            await bridge.ensureRunning();
            const params = { search };
            if (max_results)
                params.max_results = max_results;
            const response = await bridge.send('text_search', params, ANALYSIS_TIMEOUT);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // --- standard 이상에서만 등록되는 도구 ---
    if (toolset !== 'minimal') {
        server.tool('hwp_get_cell_format', '특정 표 셀의 글자 서식(글꼴, 크기, 자간, 장평, 굵기 등)과 단락 서식(정렬, 줄간격 등)을 조회합니다. 표 셀에 내용을 채우기 전에 해당 셀의 서식을 파악하여 동일한 서식으로 입력할 때 사용하세요.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
            cell_tab: z.number().int().min(0).describe('셀 Tab 인덱스 (hwp_map_table_cells로 확인)'),
        }, async ({ table_index, cell_tab }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const response = await bridge.send('get_cell_format', { table_index, cell_tab }, ANALYSIS_TIMEOUT);
                if (!response.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
                }
                return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_get_table_format_summary', '표 전체의 서식 요약을 반환합니다. 샘플 셀들의 글꼴/크기/자간/장평/줄간격을 한번에 파악합니다. 표에 내용을 채우기 전에 서식 패턴을 파악할 때 사용하세요.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
            sample_tabs: z.array(z.number().int().min(0)).optional().describe('조회할 Tab 인덱스 목록 (생략 시 첫 5개+마지막)'),
        }, async ({ table_index, sample_tabs }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const params = { table_index };
                if (sample_tabs)
                    params.sample_tabs = sample_tabs;
                const response = await bridge.send('get_table_format_summary', params, ANALYSIS_TIMEOUT);
                if (!response.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
                }
                return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_read_reference', '참고자료 파일(txt, csv, xlsx, json, md)의 내용을 추출합니다. 사업계획서 작성 시 참고 데이터를 가져올 때 사용하세요. HWP 파일은 hwp_analyze_document를 사용하세요.', {
            file_path: z.string().describe('참고자료 파일 경로'),
            max_chars: z.number().optional().describe('최대 문자 수 (기본 30000)'),
        }, async ({ file_path, max_chars }) => {
            const resolved = path.resolve(file_path);
            if (!fs.existsSync(resolved)) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `파일을 찾을 수 없습니다: ${resolved}` }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const params = { file_path: resolved };
                if (max_chars)
                    params.max_chars = max_chars;
                const response = await bridge.send('read_reference', params, ANALYSIS_TIMEOUT);
                if (!response.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
                }
                return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_get_fields', '문서의 필드(양식) 목록과 현재 값을 조회합니다. 필드를 채우기 전에 어떤 필드가 있는지 확인할 때 사용하세요.', {
            file_path: z.string().optional().describe('HWP/HWPX 파일 경로 (생략 시 현재 문서)'),
        }, async ({ file_path }) => {
            try {
                const analysis = await ensureAnalysis(bridge, file_path);
                const fields = analysis.fields || [];
                const emptyCount = fields.filter(f => !f.value || f.value.trim() === '').length;
                return {
                    content: [{
                            type: 'text',
                            text: JSON.stringify({
                                fields,
                                total_count: fields.length,
                                empty_count: emptyCount,
                            }),
                        }],
                };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_get_as_markdown', '문서 내용을 마크다운 형식으로 변환하여 반환합니다. 표는 마크다운 테이블로, 텍스트는 구조화됩니다. AI가 문서 내용을 이해하기 가장 좋은 형식입니다.', {
            file_path: z.string().optional().describe('HWP/HWPX 파일 경로 (생략 시 현재 문서)'),
        }, async ({ file_path }) => {
            try {
                const analysis = await ensureAnalysis(bridge, file_path);
                const parts = [];
                // 제목
                parts.push(`# ${analysis.file_name}`);
                parts.push(`> ${analysis.file_format} | ${analysis.pages}페이지 | 표 ${analysis.tables?.length ?? 0}개 | 필드 ${analysis.fields?.length ?? 0}개\n`);
                // 본문 텍스트
                if (analysis.full_text) {
                    parts.push('## 본문');
                    parts.push(analysis.full_text);
                    parts.push('');
                }
                // 표 데이터를 마크다운 테이블로
                const tables = analysis.tables || [];
                for (const table of tables) {
                    parts.push(`## 표 ${table.index}`);
                    if (table.headers && table.headers.length > 0) {
                        // 헤더에서 빈 문자열을 '-'로 대체
                        const headers = table.headers.map((h) => h.replace(/\r?\n/g, ' ').trim() || '-');
                        parts.push('| ' + headers.join(' | ') + ' |');
                        parts.push('| ' + headers.map(() => '---').join(' | ') + ' |');
                    }
                    if (table.data) {
                        for (const row of table.data) {
                            const cells = row.map((c) => (c ?? '').replace(/\r?\n/g, ' ').trim() || '-');
                            parts.push('| ' + cells.join(' | ') + ' |');
                        }
                    }
                    parts.push('');
                }
                // 필드
                if (analysis.fields && analysis.fields.length > 0) {
                    parts.push('## 필드');
                    for (const field of analysis.fields) {
                        parts.push(`- **${field.name}**: ${field.value || '(비어있음)'}`);
                    }
                    parts.push('');
                }
                const markdown = parts.join('\n');
                return {
                    content: [{
                            type: 'text',
                            text: JSON.stringify({
                                markdown,
                                char_count: markdown.length,
                                tables: tables.length,
                                fields: analysis.fields?.length ?? 0,
                            }),
                        }],
                };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_get_page_text', '특정 페이지의 텍스트만 추출합니다. 전체 문서가 아닌 특정 페이지 내용만 필요할 때 사용하세요.', {
            page: z.number().int().min(1).describe('페이지 번호 (1부터 시작)'),
        }, async ({ page }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                // 전체 텍스트를 가져온 뒤 페이지별로 분리하는 방식 (COM API에 페이지별 추출 없음)
                const response = await bridge.send('analyze_document', {
                    file_path: bridge.getCurrentDocument(),
                }, ANALYSIS_TIMEOUT);
                if (!response.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
                }
                const analysis = response.data;
                const totalPages = analysis.pages || 1;
                if (page > totalPages) {
                    return { content: [{ type: 'text', text: JSON.stringify({
                                    error: `페이지 ${page}은 범위를 벗어났습니다 (총 ${totalPages}페이지)`,
                                }) }], isError: true };
                }
                // 전체 텍스트를 페이지 수로 균등 분할 (근사치)
                const text = analysis.full_text || '';
                const charsPerPage = Math.ceil(text.length / totalPages);
                const start = (page - 1) * charsPerPage;
                const pageText = text.slice(start, start + charsPerPage);
                return { content: [{ type: 'text', text: JSON.stringify({
                                page,
                                total_pages: totalPages,
                                text: pageText,
                                char_count: pageText.length,
                                note: '페이지 분할은 텍스트 길이 기반 근사치입니다',
                            }) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_image_extract', '문서의 모든 이미지를 지정 디렉토리에 추출합니다. 문서 내 이미지를 파일로 가져올 때 사용하세요.', {
            output_dir: z.string().describe('이미지 저장 디렉토리 경로'),
        }, async ({ output_dir }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const resolved = path.resolve(output_dir);
                const r = await bridge.send('image_extract', { output_dir: resolved }, ANALYSIS_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_document_split', '문서를 페이지 단위로 분할하여 별도 파일로 저장합니다. 긴 문서를 나눌 때 사용하세요.', {
            output_dir: z.string().describe('분할 파일 저장 디렉토리'),
            pages_per_split: z.number().int().min(1).optional().describe('분할 단위 페이지 수 (기본: 1)'),
        }, async ({ output_dir, pages_per_split }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const resolved = path.resolve(output_dir);
                const params = { output_dir: resolved };
                if (pages_per_split)
                    params.pages_per_split = pages_per_split;
                const r = await bridge.send('document_split', params, ANALYSIS_TIMEOUT * 2);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_form_detect', '문서에서 양식 필드(빈 괄호, 체크박스, 밑줄 등)를 자동 감지합니다. 양식을 채우기 전에 어떤 필드가 있는지 파악할 때 사용하세요.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('form_detect', {}, ANALYSIS_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_extract_style_profile', '양식 문서에서 서식 프로파일(글꼴/크기/자간/장평/줄간격/들여쓰기/여백)을 추출합니다. 양식 파일을 제공받았을 때 서식을 파악하여 동일하게 적용할 때 사용하세요.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('extract_style_profile', {}, ANALYSIS_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 시각적 레이아웃 검증 ──
        server.tool('hwp_verify_layout', '현재 문서를 PNG 이미지로 변환하여 경로를 반환합니다. Claude가 Read 도구로 이미지를 읽어 표 구조, 셀 병합, 열 너비, 정렬 등을 시각적으로 검증합니다. 공문서 생성 후 결과물 확인에 사용하세요. PyMuPDF 필요(pip install PyMuPDF).', {
            pages: z.string().optional().describe('확인할 페이지 (예: "1", "5-7". 생략 시 전체)'),
        }, async ({ pages }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = {};
                if (pages)
                    params.pages = pages;
                const r = await bridge.send('verify_layout', params, 60000);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 페이지 설정 ──
        server.tool('hwp_set_page_setup', '페이지 여백, 용지 크기, 방향을 설정합니다. 공문서 작성 전 페이지 설정에 사용하세요.', {
            top_margin: z.number().optional().describe('위쪽 여백 (mm)'),
            bottom_margin: z.number().optional().describe('아래쪽 여백 (mm)'),
            left_margin: z.number().optional().describe('왼쪽 여백 (mm)'),
            right_margin: z.number().optional().describe('오른쪽 여백 (mm)'),
            header_margin: z.number().optional().describe('머리말 여백 (mm)'),
            footer_margin: z.number().optional().describe('꼬리말 여백 (mm)'),
            orientation: z.enum(['portrait', 'landscape']).optional().describe('용지 방향'),
            paper_width: z.number().optional().describe('용지 너비 (mm, 기본 A4=210)'),
            paper_height: z.number().optional().describe('용지 높이 (mm, 기본 A4=297)'),
        }, async (params) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('set_page_setup', params);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 셀 속성 설정 ──
        server.tool('hwp_set_cell_property', '표 셀의 여백, 수직 정렬, 텍스트 방향, 보호 등 속성을 설정합니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
            tab: z.number().int().min(0).describe('셀 탭 인덱스'),
            vert_align: z.enum(['top', 'middle', 'bottom']).optional().describe('수직 정렬'),
            margin_left: z.number().optional().describe('셀 왼쪽 여백 (mm)'),
            margin_right: z.number().optional().describe('셀 오른쪽 여백 (mm)'),
            margin_top: z.number().optional().describe('셀 위쪽 여백 (mm)'),
            margin_bottom: z.number().optional().describe('셀 아래쪽 여백 (mm)'),
            text_direction: z.number().int().min(0).max(1).optional().describe('텍스트 방향 (0=가로, 1=세로)'),
            protected: z.boolean().optional().describe('셀 보호'),
        }, async (params) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('set_cell_property', params);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 글상자 생성 ──
        server.tool('hwp_insert_textbox', '글상자(텍스트박스)를 생성합니다. x/y로 위치, width/height로 크기를 지정합니다. 결재란 등 위치 지정이 필요한 요소에 사용하세요.', {
            x: z.number().optional().describe('X 위치 (mm, 페이지 기준, 기본 0)'),
            y: z.number().optional().describe('Y 위치 (mm, 페이지 기준, 기본 0)'),
            width: z.number().optional().describe('너비 (mm, 기본 60)'),
            height: z.number().optional().describe('높이 (mm, 기본 30)'),
            text: z.string().optional().describe('글상자 내 텍스트'),
            border: z.boolean().optional().describe('테두리 표시 (기본 true)'),
        }, async (params) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_textbox', params);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 선 그리기 (강화) ──
        server.tool('hwp_draw_line', '선을 그립니다. 두께, 색상, 스타일을 지정할 수 있습니다. hwp_insert_line보다 상세한 제어가 가능합니다.', {
            width: z.number().optional().describe('선 두께'),
            color: z.string().optional().describe('선 색상 (#RRGGBB 또는 [R,G,B])'),
            style: z.number().int().min(0).max(5).optional().describe('선 스타일 (0=실선, 1=파선, 2=점선, 3=1점쇄선, 4=2점쇄선)'),
        }, async (params) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('draw_line', params);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 머리글/바닥글 ──
        server.tool('hwp_set_header_footer', '머리글 또는 바닥글을 설정합니다. 기관명, 페이지번호 등을 삽입할 때 사용하세요. style로 서식(굵게/정렬/크기)을 지정할 수 있습니다.', {
            type: z.enum(['header', 'footer']).describe('머리글 또는 바닥글'),
            text: z.string().optional().describe('삽입할 텍스트'),
            style: z.object({
                font_size: z.number().optional().describe('글자 크기 (pt)'),
                bold: z.boolean().optional().describe('굵게'),
                align: z.enum(['left', 'center', 'right']).optional().describe('정렬'),
                font_name: z.string().optional().describe('글꼴'),
                color: z.array(z.number()).optional().describe('색상 [R,G,B]'),
            }).optional().describe('텍스트 서식'),
        }, async ({ type, text, style }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = { type, text };
                if (style)
                    params.style = style;
                const r = await bridge.send('set_header_footer', params);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 스타일 적용 ──
        server.tool('hwp_apply_style', '현재 커서 위치에 문단 스타일을 적용합니다. "제목1", "본문", "개요1" 등 한글에 정의된 스타일을 사용합니다.', {
            style_name: z.string().describe('스타일 이름 (예: "제목1", "본문", "개요 1")'),
        }, async ({ style_name }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('apply_style', { style_name });
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 다단 설정 ──
        server.tool('hwp_set_column', '현재 섹션의 다단을 설정합니다. 2단/3단 레이아웃에 사용하세요.', {
            count: z.number().int().min(1).max(10).describe('단 수 (기본 2)'),
            gap: z.number().optional().describe('단 간격 (mm, 기본 10)'),
            line_type: z.number().int().min(0).max(5).optional().describe('구분선 종류 (0=없음, 1=실선)'),
        }, async ({ count, gap, line_type }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('set_column', { count, gap, line_type });
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 캡션 삽입 ──
        server.tool('hwp_insert_caption', '표나 그림에 캡션(제목)을 삽입합니다.', {
            text: z.string().optional().describe('캡션 텍스트'),
            side: z.number().int().min(0).max(3).optional().describe('캡션 위치 (0=왼쪽, 1=오른쪽, 2=위, 3=아래)'),
        }, async ({ text, side }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_caption', { text, side });
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 용지 설정 읽기 (F7 용지편집) ──
        server.tool('hwp_get_page_setup', '현재 문서의 용지 설정을 읽습니다. 용지 크기, 방향, 여백(위/아래/좌/우/머리말/꼬리말), 제본 여백, 사용 가능 영역을 반환합니다.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('get_page_setup', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 표 치수 추출 ──
        server.tool('hwp_get_table_dimensions', '표의 전체 너비, 셀 여백, 바깥 여백을 반환합니다. 양식 분석 시 표 구조를 정확히 재현하기 위해 사용합니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스 (0부터)'),
        }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('get_table_dimensions', { table_index });
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 양식 종합 프로파일 ──
        server.tool('hwp_extract_full_profile', '문서의 양식을 정밀 분석합니다. 용지 설정 + 본문 글자/문단 서식(19개 속성) + 표 치수(최대 5개)를 한번에 반환합니다. 양식 기반 문서 작성 전 반드시 호출하세요.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('extract_full_profile', {}, 60000);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 폰트 목록 조회 ──
        server.tool('hwp_get_font_list', '사용 가능한 한글 폰트 목록을 반환합니다. 카테고리별(serif/sans/display) 또는 공문서용(gov) 필터 가능. 40+종 한국어 폰트 포함.', {
            category: z.string().optional().describe('폰트 카테고리 필터 (serif/sans/display/mono 등)'),
            gov_only: z.boolean().optional().describe('공문서 표준 폰트만 (기본: false)'),
        }, async ({ category, gov_only }) => {
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('get_font_list', { category, gov_only: gov_only || false });
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 문서 프리셋 목록 ──
        server.tool('hwp_get_preset_list', '사용 가능한 문서/표 프리셋 목록을 반환합니다. 공문서, 사업계획서, 제안서, 보고서 등 6종 문서 프리셋 + 4종 표 스타일.', {}, async () => {
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('get_preset_list', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 문서 프리셋 적용 ──
        server.tool('hwp_apply_document_preset', '문서 프리셋을 적용합니다. 용지 설정 + 기본 폰트/줄간격을 일괄 적용합니다. 프리셋: 공문서, 사업계획서, 제안서, 보고서, 계약서, 동의서.', {
            preset_name: z.string().describe('프리셋 이름 (공문서/사업계획서/제안서/보고서/계약서/동의서)'),
        }, async ({ preset_name }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('apply_document_preset', { preset_name });
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
    } // end toolset !== 'minimal'
}
