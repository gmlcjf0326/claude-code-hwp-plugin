/**
 * Analysis tools: analyze, get text, get tables, get fields
 * HWPX 파일은 XML 직접 검색으로 라우팅 (COM 우회)
 */
import { z } from 'zod';
import path from 'node:path';
import fs from 'node:fs';
import { readHwpxXml, searchTextInSection } from '../hwpx-engine.js';
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
            // HWPX → XML 직접 검색 시도. EBUSY 시 COM 폴백.
            if (bridge.getCurrentDocumentFormat() === 'HWPX') {
                try {
                    const doc = await readHwpxXml(filePath, 'Contents/section0.xml');
                    const result = searchTextInSection(doc, search);
                    const limited = max_results ? result.results.slice(0, max_results) : result.results.slice(0, 50);
                    return { content: [{ type: 'text', text: JSON.stringify({
                                    search, total_found: result.total, results: limited, engine: 'xml',
                                }) }] };
                }
                catch (xmlErr) {
                    console.error('[text_search] XML failed, falling back to COM:', xmlErr.message);
                }
            }
            // COM 경로 (HWP 또는 HWPX XML 실패 시 폴백)
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
        server.tool('hwp_verify_layout', '현재 문서를 PDF로 내보내고 파일 경로를 반환합니다. Claude가 Read 도구로 PDF를 읽어 표 구조, 셀 병합, 열 너비, 정렬 등 레이아웃을 시각적으로 검증할 수 있습니다. 공문서 생성 후 결과물 확인에 사용하세요.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('verify_layout', {}, 60000);
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
