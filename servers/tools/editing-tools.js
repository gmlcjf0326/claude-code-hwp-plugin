/**
 * Editing tools: fill fields, fill table cells, find/replace, insert text
 * HWPX 파일은 XML 직접 조작으로 라우팅 (COM 우회)
 */
import { z } from 'zod';
import path from 'node:path';
import { readHwpxXml, writeHwpxXml, replaceTextInSection, replaceTextNthInSection, findAndAppendInSection } from '../hwpx-engine.js';
const FILL_TIMEOUT = 60000;
export function registerEditingTools(server, bridge, toolset = 'standard') {
    // --- standard 이상에서만: fill_fields ---
    if (toolset !== 'minimal') {
        server.tool('hwp_fill_fields', '문서의 필드(양식)에 값을 채웁니다. 반드시 먼저 hwp_get_fields로 필드 이름을 확인한 후 사용하세요. 필드가 없는 문서에는 hwp_fill_table_cells나 hwp_insert_text를 사용하세요.', {
            file_path: z.string().optional().describe('HWP 파일 경로 (생략 시 현재 열린 문서)'),
            fields: z.record(z.string(), z.string()).describe('채울 필드 객체 { "필드이름": "값" }'),
        }, async ({ file_path, fields }) => {
            try {
                await bridge.ensureRunning();
                const params = { fields };
                if (file_path) {
                    params.file_path = path.resolve(file_path);
                }
                const response = await bridge.send('fill_document', params, FILL_TIMEOUT);
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
    } // end fill_fields (standard+)
    // --- minimal에 포함: fill_table_cells, find_replace, insert_text ---
    server.tool('hwp_fill_table_cells', '문서의 표 셀에 값을 채웁니다. 반드시 먼저 hwp_get_tables로 표 구조를 확인하세요. 병합 셀이 있으면 label 파라미터로 라벨 텍스트 기반 매칭을 추천합니다. 예: {label: "계약금액", text: "50,000,000원"}. tab이나 row/col도 사용 가능합니다.', {
        file_path: z.string().optional().describe('HWP 파일 경로 (생략 시 현재 열린 문서)'),
        tables: z.array(z.object({
            index: z.number().int().min(0).describe('표 인덱스 (0부터 시작)'),
            cells: z.array(z.object({
                row: z.number().int().min(0).optional().describe('행 번호 (0부터, row/col 방식)'),
                col: z.number().int().min(0).optional().describe('열 번호 (0부터, row/col 방식)'),
                tab: z.number().int().min(0).optional().describe('Tab 인덱스 (0부터, 병합 셀용). hwp_map_table_cells로 확인'),
                label: z.string().optional().describe('라벨 텍스트로 셀 찾기 (예: "계약금액"). 해당 라벨 오른쪽 셀에 값을 채움'),
                row_label: z.string().optional().describe('행 라벨 텍스트 (예: "전체기간"). label과 함께 사용하면 교차점 셀을 찾음'),
                direction: z.enum(['right', 'below']).optional().default('right').describe('라벨 기준 값 셀 방향'),
                text: z.string().describe('채울 텍스트'),
                style: z.object({
                    color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('글자 색상 [R, G, B]'),
                    bold: z.boolean().optional().describe('굵게'),
                    italic: z.boolean().optional().describe('기울임'),
                    underline: z.boolean().optional().describe('밑줄'),
                    font_size: z.number().positive().optional().describe('글자 크기 (pt)'),
                    font_name: z.string().optional().describe('글꼴 이름'),
                    char_spacing: z.number().optional().describe('자간 (%)'),
                    width_ratio: z.number().optional().describe('장평 (%)'),
                    align: z.enum(['left', 'center', 'right', 'justify']).optional().describe('셀 텍스트 정렬'),
                }).optional().describe('셀별 서식 (생략 시 기존 셀 서식 상속)'),
                vert_align: z.enum(['top', 'middle', 'bottom']).optional().describe('셀 수직 정렬'),
            })).describe('채울 셀 목록. label, tab, row+col 중 하나 필수'),
        })).describe('채울 표 배열'),
    }, async ({ file_path, tables }) => {
        try {
            await bridge.ensureRunning();
            // Check if any cell uses tab-based navigation
            const results = [];
            for (const table of tables) {
                // Split cells into three groups: label, tab, row/col
                const labelCells = table.cells.filter(c => c.label !== undefined);
                const tabCells = table.cells.filter(c => c.label === undefined && c.tab !== undefined);
                const rowColCells = table.cells.filter(c => c.label === undefined && c.tab === undefined && c.row !== undefined && c.col !== undefined);
                // Label-based cells → fill_by_label
                if (labelCells.length > 0) {
                    const resp = await bridge.send('fill_by_label', {
                        table_index: table.index,
                        cells: labelCells.map(c => ({ label: c.label, text: c.text, direction: c.direction ?? 'right', ...(c.row_label ? { row_label: c.row_label } : {}) })),
                    }, FILL_TIMEOUT);
                    if (!resp.success) {
                        return { content: [{ type: 'text', text: JSON.stringify({ error: resp.error }) }], isError: true };
                    }
                    results.push(resp.data);
                }
                // Tab-based cells → fill_by_tab
                if (tabCells.length > 0) {
                    const resp = await bridge.send('fill_by_tab', {
                        table_index: table.index,
                        cells: tabCells.map(c => ({ tab: c.tab, text: c.text, ...(c.style ? { style: c.style } : {}), ...(c.vert_align ? { vert_align: c.vert_align } : {}) })),
                    }, FILL_TIMEOUT);
                    if (!resp.success) {
                        return { content: [{ type: 'text', text: JSON.stringify({ error: resp.error }) }], isError: true };
                    }
                    results.push(resp.data);
                }
                // Row/col-based cells → fill_document (legacy)
                if (rowColCells.length > 0) {
                    const params = {
                        tables: [{ index: table.index, cells: rowColCells }],
                    };
                    if (file_path)
                        params.file_path = path.resolve(file_path);
                    const resp = await bridge.send('fill_document', params, FILL_TIMEOUT);
                    if (!resp.success) {
                        return { content: [{ type: 'text', text: JSON.stringify({ error: resp.error }) }], isError: true };
                    }
                    results.push(resp.data);
                }
            }
            bridge.setCachedAnalysis(null);
            // Merge results
            const merged = { filled: 0, failed: 0, errors: [] };
            for (const r of results) {
                const d = r;
                merged.filled += d.filled ?? 0;
                merged.failed += d.failed ?? 0;
                if (d.errors)
                    merged.errors.push(...d.errors);
            }
            return { content: [{ type: 'text', text: JSON.stringify(merged) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_find_replace', '문서 전체에서 텍스트를 찾아 바꿉니다. use_regex=true로 정규식 패턴도 사용 가능합니다. case_sensitive=false로 대소문자 무시 검색 가능.', {
        find: z.string().describe('찾을 텍스트 (use_regex=true 시 정규식 패턴)'),
        replace: z.string().describe('바꿀 텍스트'),
        use_regex: z.boolean().optional().describe('정규식 사용 여부 (기본: false)'),
        case_sensitive: z.boolean().optional().describe('대소문자 구분 (기본: true). false면 대소문자 무시'),
    }, async ({ find, replace, use_regex, case_sensitive }) => {
        const filePath = bridge.getCurrentDocument();
        if (!filePath) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                        }) }], isError: true };
        }
        try {
            // HWPX → XML 직접 치환 시도 (COM 우회). EBUSY 시 COM 폴백.
            if (bridge.getCurrentDocumentFormat() === 'HWPX' && !use_regex) {
                try {
                    // COM 메모리 변경사항을 파일에 반영 (XML 엔진이 최신 내용을 읽도록)
                    await bridge.ensureRunning();
                    await bridge.send('save_document', {});
                    const doc = await readHwpxXml(filePath, 'Contents/section0.xml');
                    const count = replaceTextInSection(doc, find, replace);
                    await writeHwpxXml(filePath, filePath, 'Contents/section0.xml', doc);
                    bridge.setCachedAnalysis(null);
                    return { content: [{ type: 'text', text: JSON.stringify({
                                    status: 'ok', find, replace, replaced: count > 0, count, engine: 'xml',
                                }) }] };
                }
                catch (xmlErr) {
                    // 파일 잠금(EBUSY) 등 XML 실패 시 COM 폴백
                    console.error('[find_replace] XML failed, falling back to COM:', xmlErr.message);
                }
            }
            // COM 경로 (HWP 또는 HWPX XML 실패 시 폴백)
            await bridge.ensureRunning();
            const params = { find, replace };
            if (use_regex)
                params.use_regex = true;
            if (case_sensitive === false)
                params.case_sensitive = false;
            const response = await bridge.send('find_replace', params, FILL_TIMEOUT);
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
    // --- standard 이상에서만: find_replace_multi, find_and_append ---
    if (toolset !== 'minimal') {
        server.tool('hwp_find_replace_multi', '여러 건의 찾기/바꾸기를 일괄 실행합니다. use_regex=true로 정규식도 가능.', {
            replacements: z.array(z.object({
                find: z.string().describe('찾을 텍스트'),
                replace: z.string().describe('바꿀 텍스트'),
            })).describe('치환 목록'),
            use_regex: z.boolean().optional().describe('정규식 사용 여부 (기본: false)'),
        }, async ({ replacements, use_regex }) => {
            const filePath = bridge.getCurrentDocument();
            if (!filePath) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                            }) }], isError: true };
            }
            try {
                // HWPX → XML 직접 다건 치환 시도. EBUSY 시 COM 폴백.
                if (bridge.getCurrentDocumentFormat() === 'HWPX' && !use_regex) {
                    try {
                        // COM 메모리 변경사항을 파일에 반영 + 파일 시스템 동기 대기
                        await bridge.ensureRunning();
                        await bridge.send('save_document', {});
                        await new Promise(r => setTimeout(r, 200)); // 파일 I/O 완료 대기
                        const doc = await readHwpxXml(filePath, 'Contents/section0.xml');
                        const results = [];
                        let totalCount = 0;
                        for (const item of replacements) {
                            const count = replaceTextInSection(doc, item.find, item.replace);
                            results.push({ find: item.find, replaced: count > 0, count });
                            totalCount += count;
                        }
                        if (totalCount > 0) {
                            await writeHwpxXml(filePath, filePath, 'Contents/section0.xml', doc);
                        }
                        bridge.setCachedAnalysis(null);
                        return { content: [{ type: 'text', text: JSON.stringify({
                                        status: 'ok', results, total: results.length,
                                        success: results.filter(r => r.replaced).length, engine: 'xml',
                                    }) }] };
                    }
                    catch (xmlErr) {
                        console.error('[find_replace_multi] XML failed, falling back to COM:', xmlErr.message);
                    }
                }
                // COM 경로 (HWP 또는 HWPX XML 실패 시 폴백)
                await bridge.ensureRunning();
                const params = { replacements };
                if (use_regex)
                    params.use_regex = true;
                const response = await bridge.send('find_replace_multi', params, FILL_TIMEOUT);
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
        server.tool('hwp_find_and_append', '문서에서 텍스트를 찾은 후 그 뒤에 텍스트를 추가합니다. 색상 지정 가능. 기존 텍스트 서식을 보존하면서 새 텍스트를 추가할 때 사용하세요.', {
            find: z.string().describe('찾을 텍스트'),
            append_text: z.string().describe('찾은 텍스트 뒤에 추가할 텍스트'),
            color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('텍스트 색상 [R, G, B] (0-255)'),
        }, async ({ find, append_text, color }) => {
            const filePath = bridge.getCurrentDocument();
            if (!filePath) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                            }) }], isError: true };
            }
            try {
                // HWPX → XML 직접 조작 시도. EBUSY 시 COM 폴백.
                if (bridge.getCurrentDocumentFormat() === 'HWPX' && !color) {
                    try {
                        await bridge.ensureRunning();
                        await bridge.send('save_document', {});
                        const doc = await readHwpxXml(filePath, 'Contents/section0.xml');
                        const found = findAndAppendInSection(doc, find, append_text);
                        if (!found) {
                            return { content: [{ type: 'text', text: JSON.stringify({ status: 'not_found', find, engine: 'xml' }) }] };
                        }
                        await writeHwpxXml(filePath, filePath, 'Contents/section0.xml', doc);
                        bridge.setCachedAnalysis(null);
                        return { content: [{ type: 'text', text: JSON.stringify({ status: 'ok', find, appended: true, engine: 'xml' }) }] };
                    }
                    catch (xmlErr) {
                        console.error('[find_and_append] XML failed, falling back to COM:', xmlErr.message);
                    }
                }
                // COM 경로 (HWP 또는 HWPX XML 실패 시 폴백)
                await bridge.ensureRunning();
                // COM도 최신 파일 상태에서 검색하도록 save 선행
                await bridge.send('save_document', {});
                const params = { find, append_text };
                if (color)
                    params.color = color;
                const response = await bridge.send('find_and_append', params);
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
    } // end find_replace_multi, find_and_append (standard+)
    // --- minimal에 포함: insert_text ---
    server.tool('hwp_insert_text', '현재 커서 위치에 텍스트를 삽입합니다. 필드가 없는 문서에 텍스트를 추가할 때 사용하세요. style로 글꼴/크기/굵기/색상 등 서식 지정 가능.', {
        text: z.string().describe('삽입할 텍스트'),
        color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('텍스트 색상 [R, G, B] (0-255). style.color와 동일 (하위 호환)'),
        style: z.object({
            color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('글자 색상 [R, G, B]'),
            bold: z.boolean().optional().describe('굵게'),
            italic: z.boolean().optional().describe('기울임'),
            underline: z.boolean().optional().describe('밑줄'),
            font_size: z.number().positive().optional().describe('글자 크기 (pt)'),
            font_name: z.string().optional().describe('글꼴 이름 (예: "맑은 고딕")'),
            bg_color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('배경 색상 [R, G, B]'),
            strikeout: z.boolean().optional().describe('취소선'),
            char_spacing: z.number().optional().describe('자간 (%, 기본 0. 음수=좁게, 양수=넓게)'),
            width_ratio: z.number().optional().describe('장평 (%, 기본 100. 100 미만=좁게, 100 초과=넓게)'),
            font_name_hanja: z.string().optional().describe('한자 글꼴 이름'),
            font_name_japanese: z.string().optional().describe('일본어 글꼴 이름'),
            font_name_latin: z.string().optional().describe('라틴(영문) 전용 글꼴'),
            underline_type: z.number().int().min(0).max(7).optional().describe('밑줄 종류 (0=없음,1=실선,2=이중,3=점선,4=파선,5=1점쇄선,6=물결,7=굵은실선)'),
            underline_color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('밑줄 색상 [R,G,B]'),
            strikeout_type: z.number().int().min(0).max(3).optional().describe('취소선 종류 (0=없음,1=단일,2=이중,3=굵은)'),
            strikeout_color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('취소선 색상 [R,G,B]'),
            superscript: z.boolean().optional().describe('위 첨자'),
            subscript: z.boolean().optional().describe('아래 첨자'),
            outline: z.boolean().optional().describe('외곽선'),
            shadow: z.boolean().optional().describe('그림자'),
            shadow_color: z.array(z.number().int().min(0).max(255)).length(3).optional().describe('그림자 색상 [R,G,B]'),
            shadow_offset_x: z.number().optional().describe('그림자 X 오프셋'),
            shadow_offset_y: z.number().optional().describe('그림자 Y 오프셋'),
            emboss: z.boolean().optional().describe('양각'),
            engrave: z.boolean().optional().describe('음각'),
            small_caps: z.boolean().optional().describe('작은 대문자'),
            underline_shape: z.number().int().optional().describe('밑줄 모양'),
            strikeout_shape: z.number().int().optional().describe('취소선 모양'),
            use_kerning: z.boolean().optional().describe('커닝 (자동 자간 조정)'),
        }).optional().describe('텍스트 서식 옵션'),
    }, async ({ text, color, style }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                            hint: 'Python 프로세스가 재시작되면 열린 문서 상태가 초기화됩니다.',
                        }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const params = { text };
            if (style)
                params.style = style;
            else if (color)
                params.color = color;
            const response = await bridge.send('insert_text', params);
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
    // --- standard 이상에서만: set_paragraph_style, find_replace_nth, insert_picture ---
    if (toolset !== 'minimal') {
        server.tool('hwp_set_paragraph_style', '현재 커서 위치의 단락 서식을 변경합니다. left_margin=나머지줄 시작위치, indent=첫줄 들여쓰기. 첫줄 시작위치 = left_margin + indent.', {
            align: z.enum(['left', 'center', 'right', 'justify']).optional().describe('정렬'),
            line_spacing: z.number().optional().describe('줄간격 (%, 예: 160)'),
            line_spacing_type: z.number().int().min(0).max(2).optional().describe('줄간격 타입 (0=퍼센트)'),
            space_before: z.number().optional().describe('문단 앞 간격 (pt)'),
            space_after: z.number().optional().describe('문단 뒤 간격 (pt)'),
            indent: z.number().optional().describe('첫 줄 들여쓰기 (pt, 양수=들여쓰기, 음수=내어쓰기)'),
            left_margin: z.number().optional().describe('왼쪽 여백/나머지 줄 시작위치 (pt)'),
            right_margin: z.number().optional().describe('오른쪽 여백 (pt)'),
            page_break_before: z.boolean().optional().describe('문단 앞 페이지 나누기'),
            keep_with_next: z.boolean().optional().describe('다음 문단과 함께 (제목+본문 분리 방지)'),
            widow_orphan: z.boolean().optional().describe('과부/고아 방지'),
            line_wrap: z.number().int().optional().describe('줄 바꿈 방식'),
            snap_to_grid: z.boolean().optional().describe('그리드에 맞춤'),
            auto_space_eAsian_eng: z.boolean().optional().describe('한영 자동 간격'),
            auto_space_eAsian_num: z.boolean().optional().describe('한숫자 자동 간격'),
            break_latin_word: z.number().int().optional().describe('영문 줄바꿈 (0=단어, 1=글자)'),
            heading_type: z.number().int().optional().describe('제목 수준 (개요)'),
            keep_lines_together: z.boolean().optional().describe('줄 함께 유지 (문단 분리 방지)'),
            condense: z.number().int().optional().describe('문단 압축'),
        }, async ({ align, line_spacing, line_spacing_type, space_before, space_after, indent, left_margin, right_margin, page_break_before, keep_with_next, widow_orphan, line_wrap, snap_to_grid, auto_space_eAsian_eng, auto_space_eAsian_num, break_latin_word, heading_type, keep_lines_together, condense }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const style = {};
                if (align)
                    style.align = align;
                if (line_spacing !== undefined)
                    style.line_spacing = line_spacing;
                if (line_spacing_type !== undefined)
                    style.line_spacing_type = line_spacing_type;
                if (space_before !== undefined)
                    style.space_before = space_before;
                if (space_after !== undefined)
                    style.space_after = space_after;
                if (indent !== undefined)
                    style.indent = indent;
                if (left_margin !== undefined)
                    style.left_margin = left_margin;
                if (right_margin !== undefined)
                    style.right_margin = right_margin;
                if (page_break_before !== undefined)
                    style.page_break_before = page_break_before;
                if (keep_with_next !== undefined)
                    style.keep_with_next = keep_with_next;
                if (widow_orphan !== undefined)
                    style.widow_orphan = widow_orphan;
                if (line_wrap !== undefined)
                    style.line_wrap = line_wrap;
                if (snap_to_grid !== undefined)
                    style.snap_to_grid = snap_to_grid;
                if (auto_space_eAsian_eng !== undefined)
                    style.auto_space_eAsian_eng = auto_space_eAsian_eng;
                if (auto_space_eAsian_num !== undefined)
                    style.auto_space_eAsian_num = auto_space_eAsian_num;
                if (break_latin_word !== undefined)
                    style.break_latin_word = break_latin_word;
                if (heading_type !== undefined)
                    style.heading_type = heading_type;
                if (keep_lines_together !== undefined)
                    style.keep_lines_together = keep_lines_together;
                if (condense !== undefined)
                    style.condense = condense;
                const response = await bridge.send('set_paragraph_style', { style });
                if (!response.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
                }
                return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_find_replace_nth', '문서에서 N번째로 나타나는 텍스트만 치환합니다. 같은 텍스트가 여러 곳에 있을 때 특정 위치만 바꿀 때 사용하세요. AllReplace는 전체를 바꾸지만 이 도구는 지정한 N번째만 바꿉니다.', {
            find: z.string().describe('찾을 텍스트'),
            replace: z.string().describe('바꿀 텍스트'),
            nth: z.number().int().min(1).describe('몇 번째 매칭을 치환할지 (1부터 시작)'),
        }, async ({ find, replace, nth }) => {
            const filePath = bridge.getCurrentDocument();
            if (!filePath) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                            }) }], isError: true };
            }
            try {
                // HWPX → XML 직접 N번째 치환 시도. EBUSY 시 COM 폴백.
                if (bridge.getCurrentDocumentFormat() === 'HWPX') {
                    try {
                        await bridge.ensureRunning();
                        await bridge.send('save_document', {});
                        const doc = await readHwpxXml(filePath, 'Contents/section0.xml');
                        const replaced = replaceTextNthInSection(doc, find, replace, nth);
                        if (replaced) {
                            await writeHwpxXml(filePath, filePath, 'Contents/section0.xml', doc);
                        }
                        bridge.setCachedAnalysis(null);
                        return { content: [{ type: 'text', text: JSON.stringify({
                                        status: 'ok', find, replace, nth, replaced, engine: 'xml',
                                    }) }] };
                    }
                    catch (xmlErr) {
                        console.error('[find_replace_nth] XML failed, falling back to COM:', xmlErr.message);
                    }
                }
                // HWP → Python COM
                await bridge.ensureRunning();
                const response = await bridge.send('find_replace_nth', { find, replace, nth });
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
        server.tool('hwp_insert_picture', '현재 커서 위치에 이미지를 삽입합니다. 표 셀 안에서도 사용 가능합니다. 사업계획서의 제품사진 등을 삽입할 때 사용하세요.', {
            file_path: z.string().describe('이미지 파일 경로 (jpg, png, bmp 등)'),
            width: z.number().min(0).optional().describe('가로 크기 (mm, 0이면 원본 크기)'),
            height: z.number().min(0).optional().describe('세로 크기 (mm, 0이면 원본 크기)'),
        }, async ({ file_path, width, height }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({
                                error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                            }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const params = { file_path: path.resolve(file_path) };
                if (width)
                    params.width = width;
                if (height)
                    params.height = height;
                const response = await bridge.send('insert_picture', params);
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
        server.tool('hwp_table_add_row', '표에 행을 추가합니다. 표의 현재 마지막 행 아래에 새 행이 추가됩니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스 (0부터)'),
        }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const response = await bridge.send('table_add_row', { table_index }, FILL_TIMEOUT);
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
        server.tool('hwp_insert_markdown', '마크다운 텍스트를 한글 서식으로 변환하여 현재 커서 위치에 삽입합니다. # 제목, **굵게**, - 목록 등을 지원합니다.', {
            text: z.string().describe('마크다운 텍스트'),
        }, async ({ text }) => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const response = await bridge.send('insert_markdown', { text }, FILL_TIMEOUT);
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
        server.tool('hwp_insert_page_break', '현재 커서 위치에 페이지 나누기를 삽입합니다.', {}, async () => {
            if (!bridge.getCurrentDocument()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            }
            try {
                await bridge.ensureRunning();
                const response = await bridge.send('insert_page_break', {});
                if (!response.success) {
                    return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
                }
                return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_delete_row', '표에서 현재 행을 삭제합니다.', { table_index: z.number().int().min(0).describe('표 인덱스') }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_delete_row', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_add_column', '표에 열을 추가합니다 (현재 열 오른쪽).', { table_index: z.number().int().min(0).describe('표 인덱스') }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_add_column', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_delete_column', '표에서 현재 열을 삭제합니다.', { table_index: z.number().int().min(0).describe('표 인덱스') }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_delete_column', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_merge_cells', '표의 셀을 병합합니다. start_row/col ~ end_row/col로 범위를 지정하면 해당 영역이 병합됩니다. 범위 미지정 시 현재 선택된 셀이 병합됩니다. 병합 순서는 하단→상단이 안전합니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
            start_row: z.number().int().min(0).optional().describe('시작 행 (0부터)'),
            start_col: z.number().int().min(0).optional().describe('시작 열 (0부터)'),
            end_row: z.number().int().min(0).optional().describe('끝 행 (0부터)'),
            end_col: z.number().int().min(0).optional().describe('끝 열 (0부터)'),
        }, async ({ table_index, start_row, start_col, end_row, end_col }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = { table_index };
                if (start_row !== undefined)
                    params.start_row = start_row;
                if (start_col !== undefined)
                    params.start_col = start_col;
                if (end_row !== undefined)
                    params.end_row = end_row;
                if (end_col !== undefined)
                    params.end_col = end_col;
                const r = await bridge.send('table_merge_cells', params, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_create_from_data', '2D 배열 데이터로 새 표를 생성합니다. col_widths로 열 너비(mm), row_heights로 행 높이(mm)를 지정할 수 있습니다. 공문서 표 등 정밀한 레이아웃에 사용하세요.', {
            data: z.array(z.array(z.string())).describe('2D 배열 데이터 [["헤더1","헤더2"],["값1","값2"]]'),
            header_style: z.boolean().optional().describe('첫 행을 헤더로 자동 스타일링 (Bold+배경색)'),
            col_widths: z.array(z.number()).optional().describe('열 너비 배열 (mm 단위, 예: [18, 65, 23, 23])'),
            row_heights: z.array(z.number()).optional().describe('행 높이 배열 (mm 단위, 예: [10, 12, 12])'),
            alignment: z.enum(['left', 'center', 'right']).optional().describe('표 정렬 (기본: left)'),
        }, async ({ data, header_style, col_widths, row_heights, alignment }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = { data };
                if (header_style)
                    params.header_style = header_style;
                if (col_widths)
                    params.col_widths = col_widths;
                if (row_heights)
                    params.row_heights = row_heights;
                if (alignment)
                    params.alignment = alignment;
                const r = await bridge.send('table_create_from_data', params, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_heading', '제목 텍스트를 삽입합니다 (H1~H6). 공문서 순번 체계의 대제목 등에 사용. numbering으로 자동 순번을 붙일 수 있습니다 (예: Ⅰ. 제목, 1. 제목, 가. 제목).', {
            text: z.string().describe('제목 텍스트'),
            level: z.number().int().min(1).max(6).describe('제목 레벨 (1=가장 큰 22pt, 6=가장 작은 10pt)'),
            numbering: z.enum(['roman', 'decimal', 'korean', 'circle', 'paren_decimal', 'paren_korean']).optional().describe('순번 형식: roman(Ⅰ,Ⅱ), decimal(1,2), korean(가,나), circle(①,②), paren_decimal(1),2)), paren_korean(가),나))'),
            number: z.number().int().min(1).max(10).optional().describe('순번 번호 (1~10, 기본 1)'),
        }, async ({ text, level, numbering, number }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = { text, level };
                if (numbering)
                    params.numbering = numbering;
                if (number)
                    params.number = number;
                const r = await bridge.send('insert_heading', params, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_export_docx', '현재 문서를 DOCX(Word) 형식으로 내보냅니다.', { output_path: z.string().describe('DOCX 저장 경로') }, async ({ output_path }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const resolved = path.resolve(output_path);
                const r = await bridge.send('export_format', { path: resolved, format: 'OOXML' }, 120000);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_export_html', '현재 문서를 HTML 형식으로 내보냅니다.', { output_path: z.string().describe('HTML 저장 경로') }, async ({ output_path }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const resolved = path.resolve(output_path);
                const r = await bridge.send('export_format', { path: resolved, format: 'HTML' }, 60000);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_set_background_picture', '문서에 배경 이미지를 설정합니다. 워터마크로 활용 가능합니다.', { file_path: z.string().describe('배경 이미지 파일 경로') }, async ({ file_path }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const resolved = path.resolve(file_path);
                const r = await bridge.send('set_background_picture', { file_path: resolved }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_split_cell', '표에서 현재 셀을 분할합니다.', { table_index: z.number().int().min(0).describe('표 인덱스') }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_split_cell', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_hyperlink', '현재 커서 위치에 하이퍼링크를 삽입합니다.', {
            url: z.string().describe('URL (예: https://example.com)'),
            text: z.string().optional().describe('표시 텍스트 (생략 시 URL)'),
        }, async ({ url, text }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = { url };
                if (text)
                    params.text = text;
                const r = await bridge.send('insert_hyperlink', params, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_footnote', '현재 커서 위치에 각주를 삽입합니다. 학술 문서나 보고서에서 참조 주석을 달 때 사용하세요.', { text: z.string().optional().describe('각주 내용 (생략 시 빈 각주)') }, async ({ text }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_footnote', text ? { text } : {}, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_endnote', '현재 커서 위치에 미주를 삽입합니다.', { text: z.string().optional().describe('미주 내용') }, async ({ text }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_endnote', text ? { text } : {}, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_page_num', '현재 커서 위치에 쪽 번호를 삽입합니다. format으로 형식을 지정할 수 있습니다 (예: - 1 -, (1)).', {
            format: z.enum(['plain', 'dash', 'paren']).optional().describe('페이지 번호 형식: plain(기본), dash(- 1 -), paren((1))'),
        }, async ({ format }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = {};
                if (format)
                    params.format = format;
                const r = await bridge.send('insert_page_num', params);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_date_code', '현재 커서 위치에 오늘 날짜를 자동 삽입합니다.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_date_code', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_formula_sum', '표에서 합계를 자동 계산합니다. 현재 셀 위치에서 한글의 자동 합계 기능을 실행합니다.', { table_index: z.number().int().min(0).describe('표 인덱스') }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_formula_sum', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_formula_avg', '표에서 평균을 자동 계산합니다. 현재 셀 위치에서 한글의 자동 평균 기능을 실행합니다.', { table_index: z.number().int().min(0).describe('표 인덱스') }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_formula_avg', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── Phase B: Quick Win 8개 ──
        server.tool('hwp_table_to_csv', '표 데이터를 CSV 파일로 내보냅니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
            output_path: z.string().describe('CSV 저장 경로'),
        }, async ({ table_index, output_path }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_to_csv', { table_index, output_path: path.resolve(output_path) }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_break_section', '현재 위치에 섹션 나누기를 삽입합니다.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('break_section', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_break_column', '현재 위치에 다단 나누기를 삽입합니다.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('break_column', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_line', '현재 위치에 선(줄)을 삽입합니다.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_line', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_swap_type', '표의 행과 열을 교환합니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
        }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_swap_type', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_auto_num', '자동 번호매기기를 삽입합니다.', {}, async () => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_auto_num', {});
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_insert_memo', '메모 필드를 삽입합니다.', {
            text: z.string().optional().describe('메모 내용'),
        }, async ({ text }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('insert_memo', text ? { text } : {}, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_table_distribute_width', '표 셀 너비를 균등하게 분배합니다.', {
            table_index: z.number().int().min(0).describe('표 인덱스'),
        }, async ({ table_index }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('table_distribute_width', { table_index }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_indent', '현재 커서 위치의 단락을 들여쓰기합니다 (Shift+Tab 효과). 공문서 순번 체계에서 하위 항목 들여쓰기에 사용.', {
            depth: z.number().optional().describe('들여쓰기 깊이 (pt, 기본 10)'),
        }, async ({ depth }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('indent', depth ? { depth } : {}, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_outdent', '현재 커서 위치의 단락 들여쓰기를 줄입니다 (내어쓰기).', {
            depth: z.number().optional().describe('내어쓰기 깊이 (pt, 기본 10)'),
        }, async ({ depth }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('outdent', depth ? { depth } : {}, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_delete_guide_text', '양식의 작성요령/가이드 텍스트(< 작성요령 >, ※ 안내문 등)를 자동 삭제합니다. 공공기관 양식 작성 후 제출 전에 사용.', {
            patterns: z.array(z.string()).optional().describe('삭제할 텍스트 패턴 목록 (기본: < 작성요령 >)'),
        }, async ({ patterns }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = {};
                if (patterns)
                    params.patterns = patterns;
                const r = await bridge.send('delete_guide_text', params, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        server.tool('hwp_toggle_checkbox', '문서의 체크박스를 전환합니다 (□→■, ☐→☑ 등). 양식에서 특정 항목을 체크할 때 사용.', {
            find: z.string().describe('찾을 체크박스 텍스트 (예: "☐ 유")'),
            replace: z.string().describe('바꿀 텍스트 (예: "■ 유")'),
        }, async ({ find, replace }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('toggle_checkbox', { find, replace }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                bridge.setCachedAnalysis(null);
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 표 셀 배경색 ──
        server.tool('hwp_set_cell_color', '표 셀의 배경색을 설정합니다. 간트차트 음영, 헤더행 강조, 데이터 시각화 등에 사용.', {
            table_index: z.number().int().describe('표 인덱스 (0부터, -1=현재 위치한 표)'),
            cells: z.array(z.object({
                tab: z.number().int().describe('셀 탭 인덱스'),
                color: z.string().describe('배경색 (#RRGGBB 형식, 예: "#E8E8E8")'),
            })).describe('배경색을 설정할 셀 목록'),
        }, async ({ table_index, cells }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const r = await bridge.send('set_cell_color', { table_index, cells }, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
        // ── 표 테두리 스타일 ──
        server.tool('hwp_set_table_border', '표의 테두리 스타일을 설정합니다. 표 전체 또는 특정 셀의 테두리를 변경할 수 있습니다.', {
            table_index: z.number().int().describe('표 인덱스 (0부터)'),
            cells: z.array(z.object({
                tab: z.number().int().describe('셀 탭 인덱스'),
            })).optional().describe('특정 셀만 적용 (생략 시 표 전체)'),
            style: z.object({
                line_type: z.number().int().min(0).max(5).optional().describe('선 종류: 0=없음, 1=실선, 2=파선, 3=점선, 4=1점쇄선, 5=2점쇄선'),
                line_width: z.number().optional().describe('선 두께 (pt 단위)'),
                color: z.string().optional().describe('테두리 색상 (#RRGGBB, 예: "#003366")'),
                edges: z.array(z.enum(['left', 'right', 'top', 'bottom'])).optional().describe('적용할 방향 (생략 시 전체)'),
            }).optional().describe('테두리 스타일'),
        }, async ({ table_index, cells, style }) => {
            if (!bridge.getCurrentDocument())
                return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
            try {
                await bridge.ensureRunning();
                const params = { table_index };
                if (cells)
                    params.cells = cells;
                if (style)
                    params.style = style;
                const r = await bridge.send('set_table_border', params, FILL_TIMEOUT);
                if (!r.success)
                    return { content: [{ type: 'text', text: JSON.stringify({ error: r.error }) }], isError: true };
                return { content: [{ type: 'text', text: JSON.stringify(r.data) }] };
            }
            catch (err) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
            }
        });
    } // end standard+ tools
}
