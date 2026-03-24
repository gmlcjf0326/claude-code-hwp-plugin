/**
 * Document lifecycle tools: list, open, close, save
 */
import { z } from 'zod';
import fs from 'node:fs';
import path from 'node:path';
const HWP_EXTENSIONS = new Set(['.hwp', '.hwpx']);
const SKIP_DIRS = new Set(['node_modules', '.git', '__pycache__', 'dist', '.next']);
const MAX_SCAN_DEPTH = 10;
function formatKoreanDate(date) {
    const y = date.getFullYear();
    const m = date.getMonth() + 1;
    const d = date.getDate();
    const h = String(date.getHours()).padStart(2, '0');
    const min = String(date.getMinutes()).padStart(2, '0');
    return `${y}년 ${m}월 ${d}일 ${h}:${min}`;
}
function listHwpFiles(directory, recursive) {
    const results = [];
    const visited = new Set();
    function scan(dir, depth) {
        if (depth > MAX_SCAN_DEPTH)
            return;
        let realDir;
        try {
            realDir = fs.realpathSync(dir);
        }
        catch {
            return;
        }
        if (visited.has(realDir))
            return;
        visited.add(realDir);
        let entries;
        try {
            entries = fs.readdirSync(dir, { withFileTypes: true });
        }
        catch {
            return;
        }
        for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);
            if (entry.isDirectory() || entry.isSymbolicLink()) {
                try {
                    const stat = fs.statSync(fullPath);
                    if (stat.isDirectory()) {
                        if (recursive && !SKIP_DIRS.has(entry.name) && !entry.name.startsWith('.')) {
                            scan(fullPath, depth + 1);
                        }
                        continue;
                    }
                }
                catch {
                    continue;
                }
            }
            if (entry.isDirectory())
                continue;
            const ext = path.extname(entry.name).toLowerCase();
            if (!HWP_EXTENSIONS.has(ext))
                continue;
            try {
                const stat = fs.statSync(fullPath);
                results.push({
                    path: fullPath,
                    name: entry.name,
                    size: stat.size,
                    modifiedAt: stat.mtime.toISOString(),
                    modifiedAtKR: formatKoreanDate(stat.mtime),
                });
            }
            catch {
                // skip inaccessible files
            }
        }
    }
    scan(directory, 0);
    return results;
}
export function registerDocumentTools(server, bridge) {
    // ── 환경 진단 도구 (Python/pyhwpx/한글 없이도 동작) ──
    server.tool('hwp_check_setup', '사용 환경을 진단합니다. Python, pyhwpx, 한글 프로그램의 설치 여부를 확인하고 미설치 항목의 설치 방법을 안내합니다. 처음 사용하거나 에러 발생 시 이 도구를 먼저 호출하세요.', {}, async () => {
        try {
            const prereq = await bridge.checkPrerequisites();
            const items = [];
            if (!prereq.os.ok) {
                items.push(`❌ OS: ${prereq.os.error}`);
            }
            if (prereq.python.found) {
                let pyInfo = `✅ Python ${prereq.python.version} (${prereq.python.path || 'unknown'})`;
                if (prereq.python.guide)
                    pyInfo += `\n   ⚠️ ${prereq.python.guide}`;
                items.push(pyInfo);
            }
            else {
                items.push(`❌ Python 미설치\n   ${prereq.python.guide}`);
            }
            if (prereq.pyhwpx.found) {
                items.push(`✅ pyhwpx ${prereq.pyhwpx.version || ''}`);
            }
            else if (prereq.python.found) {
                items.push(`❌ pyhwpx 미설치\n   ${prereq.pyhwpx.guide}`);
            }
            if (prereq.hwp.found) {
                items.push('✅ 한글(HWP) 프로그램 설치됨');
            }
            else if (prereq.pyhwpx.found) {
                items.push(`❌ 한글(HWP) 미설치\n   ${prereq.hwp.guide}`);
            }
            // 한글 실행 여부
            if (prereq.hwp.found) {
                if (prereq.hwpRunning) {
                    items.push('✅ 한글(HWP) 실행 중');
                }
                else {
                    items.push('⚠️ 한글(HWP)이 실행되지 않았습니다. 문서 작업 전 한글을 먼저 실행하세요.');
                }
            }
            const allReady = prereq.ok && prereq.hwpRunning;
            return { content: [{ type: 'text', text: JSON.stringify({
                            status: allReady ? 'ready' : prereq.ok ? 'hwp_not_running' : 'not_ready',
                            message: allReady
                                ? '모든 요구사항이 충족되었습니다. HWP 도구를 사용할 수 있습니다.'
                                : prereq.ok
                                    ? '한글(HWP) 프로그램을 실행한 후 다시 시도하세요.'
                                    : '아래 항목을 설치한 후 다시 시도하세요.',
                            details: prereq,
                            summary: items.join('\n'),
                        }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_list_files', '디렉토리 내 HWP/HWPX 파일 목록을 반환합니다. Python/한글 프로그램 없이도 사용 가능합니다. 문서 작업 전 파일 위치를 확인할 때 먼저 호출하세요.', {
        directory: z.string().optional().describe('탐색할 디렉토리 경로 (기본: 현재 디렉토리)'),
        recursive: z.boolean().optional().describe('하위 디렉토리 재귀 탐색 여부 (기본: false)'),
    }, async ({ directory, recursive }) => {
        const dir = directory ? path.resolve(directory) : process.cwd();
        if (!fs.existsSync(dir)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `디렉토리를 찾을 수 없습니다: ${dir}` }) }], isError: true };
        }
        const files = listHwpFiles(dir, recursive ?? false);
        return {
            content: [{ type: 'text', text: JSON.stringify({ directory: dir, files, total: files.length }) }],
        };
    });
    server.tool('hwp_open_document', '지정된 경로의 HWP/HWPX 파일을 열어 편집 준비합니다. 이미 열린 문서가 있으면 자동으로 닫고 새 문서를 엽니다. 문서를 열면 hwp_analyze_document로 구조를 파악하세요.', {
        file_path: z.string().describe('HWP/HWPX 파일의 절대 또는 상대 경로'),
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
            // P1 #9: 이미 열린 문서가 있으면 먼저 닫기
            if (bridge.getCurrentDocument()) {
                try {
                    await bridge.send('close_document', {});
                }
                catch { /* Python이 죽었을 수 있으므로 무시 */ }
                bridge.setCurrentDocument(null);
                bridge.setCachedAnalysis(null);
            }
            const response = await bridge.send('open_document', { file_path: resolved }, 60000);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            bridge.setCurrentDocument(resolved);
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_close_document', '현재 열린 HWP 문서를 닫습니다. 다른 문서를 열기 전이나 작업 완료 후 호출하세요.', async () => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({
                            error: '열린 문서가 없습니다.',
                            hint: 'Python 프로세스가 재시작되면 열린 문서 상태가 초기화됩니다.',
                        }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('close_document', {});
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            bridge.setCurrentDocument(null);
            bridge.setCachedAnalysis(null);
            return { content: [{ type: 'text', text: JSON.stringify({ status: 'ok', message: '문서를 닫았습니다.' }) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    server.tool('hwp_save_document', '현재 열린 문서를 지정된 경로와 형식으로 저장합니다. 편집 작업 후 반드시 호출하여 변경사항을 저장하세요.', {
        path: z.string().describe('저장할 파일 경로'),
        format: z.enum(['hwp', 'hwpx', 'pdf', 'docx']).optional().describe('저장 형식 (생략 시 경로 확장자에서 추론, 기본: hwp)'),
    }, async ({ path: savePath, format }) => {
        if (!bridge.getCurrentDocument()) {
            return {
                content: [{ type: 'text', text: JSON.stringify({
                            error: '열린 문서가 없습니다. hwp_open_document로 문서를 열어주세요.',
                            hint: 'Python 프로세스가 재시작되면 열린 문서 상태가 초기화됩니다.',
                        }) }],
                isError: true,
            };
        }
        const resolved = path.resolve(savePath);
        const dir = path.dirname(resolved);
        if (!fs.existsSync(dir)) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: `저장 디렉토리가 존재하지 않습니다: ${dir}` }) }], isError: true };
        }
        // P1 #6: 경로가 디렉토리인지 확인
        try {
            if (fs.existsSync(resolved) && fs.statSync(resolved).isDirectory()) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: `저장 경로가 디렉토리입니다. 파일 이름을 포함한 경로를 지정하세요: ${resolved}` }) }], isError: true };
            }
        }
        catch { /* stat 실패는 무시 — 파일이 아직 없을 수 있음 */ }
        // P1 #6: format 미지정 시 확장자에서 추론
        let saveFormat = format;
        if (!saveFormat) {
            const ext = path.extname(resolved).toLowerCase().replace('.', '');
            if (['hwp', 'hwpx', 'pdf', 'docx'].includes(ext)) {
                saveFormat = ext;
            }
            else {
                saveFormat = 'hwp';
            }
        }
        // P1 #6: 확장자가 format과 불일치 시 자동 추가
        let finalPath = resolved;
        const currentExt = path.extname(resolved).toLowerCase().replace('.', '');
        if (currentExt !== saveFormat) {
            finalPath = `${resolved}.${saveFormat}`;
        }
        try {
            await bridge.ensureRunning();
            const response = await bridge.send('save_as', {
                path: finalPath,
                format: saveFormat,
            });
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
    // ── PDF 전용 내보내기 ──
    server.tool('hwp_export_pdf', '현재 문서를 PDF로 내보냅니다. "PDF로 변환해줘" 요청에 사용하세요.', {
        output_path: z.string().describe('PDF 저장 경로 (예: C:/output/문서.pdf)'),
    }, async ({ output_path }) => {
        if (!bridge.getCurrentDocument()) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: '열린 문서가 없습니다.' }) }], isError: true };
        }
        try {
            await bridge.ensureRunning();
            const resolved = path.resolve(output_path);
            const response = await bridge.send('save_as', { path: resolved, format: 'pdf' }, 60000);
            if (!response.success) {
                return { content: [{ type: 'text', text: JSON.stringify({ error: response.error }) }], isError: true };
            }
            return { content: [{ type: 'text', text: JSON.stringify(response.data) }] };
        }
        catch (err) {
            return { content: [{ type: 'text', text: JSON.stringify({ error: err.message }) }], isError: true };
        }
    });
}
