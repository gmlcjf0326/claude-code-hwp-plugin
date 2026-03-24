/**
 * HWP Python Bridge for MCP Server
 * Adapted from electron/services/hwp-bridge.ts — no Electron dependencies.
 * All logging via console.error() to protect stdout (MCP JSON-RPC).
 */
import { spawn, execFile } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const execFileAsync = promisify(execFile);
export class HwpBridge {
    process = null;
    requestId = 0;
    pending = new Map();
    buffer = '';
    MAX_BUFFER_SIZE = 10 * 1024 * 1024; // 10MB
    startPromise = null;
    pythonRunning = false;
    currentDocumentPath = null;
    currentDocumentFormat = null;
    lastAnalysis = null;
    lastError = null;
    startTime = Date.now();
    getPythonScriptDir() {
        // npm 패키지: dist/hwp-bridge.js → ../python (패키지 내 python/)
        const npmPath = path.resolve(__dirname, '../python');
        if (fs.existsSync(path.join(npmPath, 'hwp_service.py')))
            return npmPath;
        // 개발 환경: mcp-server/dist/hwp-bridge.js → ../../python (프로젝트 루트)
        return path.resolve(__dirname, '../../python');
    }
    findPython() {
        return process.env.PYTHON_PATH || 'python';
    }
    async ensureRunning() {
        if (this.process && this.pythonRunning)
            return;
        // 동시 재시작 방지: 이미 시작 중이면 기다림
        if (this.startPromise)
            return this.startPromise;
        // Clean up previous process
        if (this.process) {
            this.process.kill();
            this.process = null;
        }
        this.pythonRunning = false;
        this.currentDocumentPath = null;
        this.currentDocumentFormat = null;
        this.lastAnalysis = null;
        this.lastError = null;
        this.startPromise = this.start();
        try {
            await this.startPromise;
        }
        finally {
            this.startPromise = null;
        }
    }
    async start() {
        const pythonExe = this.findPython();
        const scriptPath = path.join(this.getPythonScriptDir(), 'hwp_service.py');
        console.error(`[HWP MCP Bridge] Starting Python: ${pythonExe} ${scriptPath}`);
        this.process = spawn(pythonExe, [scriptPath], {
            stdio: ['pipe', 'pipe', 'pipe'],
            env: { ...process.env, PYTHONUNBUFFERED: '1' },
        });
        // Python 실행 파일 자체를 찾을 수 없는 경우 (ENOENT)
        this.process.on('error', (err) => {
            if (err.code === 'ENOENT') {
                this.lastError = 'Python을 찾을 수 없습니다. Python 3.8+을 설치하고 PATH에 추가하세요.\n→ https://www.python.org/downloads/ (설치 시 "Add to PATH" 체크)';
            }
            else {
                this.lastError = `Python 프로세스 시작 실패: ${err.message}`;
            }
            console.error('[HWP MCP Bridge] Spawn error:', err.message);
            this.pythonRunning = false;
            this.rejectAllPending(new Error(this.lastError));
        });
        this.process.stdout?.on('data', (chunk) => {
            this.buffer += chunk.toString('utf-8');
            // BUG-6 fix: 버퍼 초과 시 가장 오래된 대기 요청만 거부 (전체가 아닌 개별)
            if (this.buffer.length > this.MAX_BUFFER_SIZE) {
                console.error(`[HWP MCP Bridge] Buffer exceeded ${this.MAX_BUFFER_SIZE / 1024 / 1024}MB — truncating oldest pending`);
                // 버퍼를 비우되, 마지막 줄바꿈 이후 부분은 보존 (진행 중인 응답)
                const lastNewline = this.buffer.lastIndexOf('\n');
                this.buffer = lastNewline >= 0 ? this.buffer.slice(lastNewline + 1) : '';
                // 가장 오래된 요청 하나만 거부
                const oldestId = this.pending.keys().next().value;
                if (oldestId) {
                    const oldest = this.pending.get(oldestId);
                    if (oldest) {
                        clearTimeout(oldest.timer);
                        oldest.reject(new Error('응답 크기가 너무 큽니다.'));
                        this.pending.delete(oldestId);
                    }
                }
                return;
            }
            this.processBuffer();
        });
        this.process.stderr?.on('data', (chunk) => {
            const text = chunk.toString('utf-8').trim();
            console.error('[HWP MCP Bridge][stderr]', text);
            if (text.includes('ModuleNotFoundError')) {
                this.lastError = 'Python 모듈을 찾을 수 없습니다. pip install pyhwpx pywin32 를 실행해주세요.';
            }
            else if (text.includes('COM class not registered') || text.includes('CoInitialize')) {
                this.lastError = '한글(HWP) 프로그램이 설치되어 있어야 합니다. 한글이 설치되어 있는지 확인하세요.';
            }
            else if (text.includes('RPC') || text.includes('사용할 수 없습니다')) {
                this.lastError = 'RPC 서버를 사용할 수 없습니다. 한글 프로그램이 실행 중인지 확인하세요.';
            }
            else if (text.includes('SyntaxError') || text.includes('ImportError')) {
                this.lastError = `Python 오류가 발생했습니다: ${text.split('\n').pop()}`;
            }
        });
        this.process.on('exit', (code) => {
            console.error(`[HWP MCP Bridge] Python exited with code ${code}`);
            this.rejectAllPending(new Error(`Python process exited with code ${code}`));
            this.process = null;
            this.pythonRunning = false;
            this.currentDocumentPath = null;
            this.currentDocumentFormat = null;
            this.lastAnalysis = null;
        });
        // Verify connection (with 2 retries, 5초 간격)
        let lastErr;
        for (let attempt = 1; attempt <= 3; attempt++) {
            try {
                // ping 시 Python에서 Hwp() COM 초기화 — 한글 프로그램 시작까지 최대 90초
                const response = await this.send('ping', {}, 90000);
                if (response.success) {
                    this.pythonRunning = true;
                    console.error('[HWP MCP Bridge] Python bridge connected');
                    return;
                }
            }
            catch (err) {
                lastErr = err;
                console.error(`[HWP MCP Bridge] Ping attempt ${attempt}/3 failed:`, err);
                if (attempt < 3) {
                    console.error('[HWP MCP Bridge] Retrying in 5 seconds...');
                    await new Promise(resolve => setTimeout(resolve, 5000));
                }
            }
        }
        // All 3 attempts failed — 단계별 안내 제공
        const detail = this.lastError || '';
        let hint;
        if (detail.includes('Python을 찾을 수 없습니다')) {
            hint = detail;
        }
        else if (detail.includes('RPC') || detail.includes('사용할 수 없습니다')) {
            hint = '한글 프로그램이 응답하지 않습니다. 한글을 닫고 다시 시도해주세요.';
        }
        else if (detail.includes('ModuleNotFoundError') || detail.includes('pyhwpx')) {
            hint = 'pyhwpx가 설치되지 않았습니다. 터미널에서 실행: pip install pyhwpx';
        }
        else if (detail.includes('COM') || detail.includes('한글')) {
            hint = '한글(HWP) 프로그램이 설치되지 않았거나 COM 등록이 실패했습니다.';
        }
        else {
            hint = 'hwp_check_setup 도구로 환경을 진단해보세요.\n  확인사항: 1) Python 3.8+ 설치  2) pip install pyhwpx  3) 한글 프로그램 설치';
        }
        throw new Error(`HWP MCP 시작 실패: ${hint}`);
    }
    processBuffer() {
        let newlineIndex;
        while ((newlineIndex = this.buffer.indexOf('\n')) !== -1) {
            const line = this.buffer.slice(0, newlineIndex).trim();
            this.buffer = this.buffer.slice(newlineIndex + 1);
            if (!line)
                continue;
            try {
                const response = JSON.parse(line);
                const pending = this.pending.get(response.id);
                if (pending) {
                    clearTimeout(pending.timer);
                    this.pending.delete(response.id);
                    pending.resolve(response);
                }
            }
            catch {
                console.error('[HWP MCP Bridge] Invalid JSON from Python:', line);
            }
        }
    }
    async send(method, params, timeoutMs = 30000) {
        // start() 내부의 ping 호출 시에는 이미 process가 있으므로 재시작 하지 않음
        if (!this.process || !this.process.stdin?.writable) {
            if (method !== 'ping') {
                this.lastError = null;
                await this.ensureRunning();
            }
        }
        if (!this.process?.stdin?.writable) {
            const detail = this.lastError || 'Python과 pyhwpx가 설치되어 있는지 확인하세요.';
            throw new Error(`Python 프로세스를 시작할 수 없습니다. ${detail}`);
        }
        const id = `req_${++this.requestId}`;
        const request = { id, method, params };
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => {
                this.pending.delete(id);
                const detail = this.lastError ? ` 원인: ${this.lastError}` : '';
                reject(new Error(`요청 시간 초과 (${timeoutMs / 1000}초): ${method}.${detail}`));
            }, timeoutMs);
            this.pending.set(id, { resolve, reject, timer });
            this.process.stdin.write(JSON.stringify(request) + '\n', (err) => {
                if (err) {
                    clearTimeout(timer);
                    this.pending.delete(id);
                    reject(new Error(`Failed to write to Python stdin: ${err.message}`));
                }
            });
        });
    }
    async shutdown() {
        if (!this.process)
            return;
        try {
            await this.send('shutdown', {}, 5000);
        }
        catch {
            // ignore timeout on shutdown
        }
        this.process.kill();
        this.process = null;
        this.pythonRunning = false;
        this.rejectAllPending(new Error('Bridge shut down'));
    }
    rejectAllPending(error) {
        for (const [id, pending] of this.pending) {
            clearTimeout(pending.timer);
            pending.reject(error);
            this.pending.delete(id);
        }
    }
    // ── 사전 요구사항 체크 (Python/pyhwpx/한글) ──
    async checkPrerequisites() {
        const result = {
            ok: false,
            python: { found: false },
            pyhwpx: { found: false },
            hwp: { found: false },
            hwpRunning: false,
            os: { ok: process.platform === 'win32', platform: process.platform },
        };
        if (!result.os.ok) {
            result.os.error = 'HWP MCP는 Windows 전용입니다. 한글(HWP)은 Windows COM API를 사용합니다.';
            return result;
        }
        const pythonExe = this.findPython();
        // 1) Python 체크 — 경로 + 버전 + Microsoft Store 감지
        try {
            const { stdout } = await execFileAsync(pythonExe, ['-c', 'import sys; print(sys.version.split()[0]); print(sys.executable)'], { timeout: 5000 });
            const lines = stdout.trim().split(/\r?\n/);
            const ver = lines[0];
            const exePath = lines[1] || '';
            const isStorePython = exePath.includes('WindowsApps');
            result.python = {
                found: true,
                version: ver,
                path: exePath,
                guide: isStorePython
                    ? `Microsoft Store Python 감지 (${exePath}). pyhwpx가 인식되지 않을 수 있습니다.\n→ python.org에서 재설치를 권장합니다: https://www.python.org/downloads/`
                    : undefined,
            };
        }
        catch {
            result.python = {
                found: false,
                guide: 'Python 3.8+ 설치 필요\n→ https://www.python.org/downloads/ 에서 설치\n→ 설치 시 "Add Python to PATH" 반드시 체크\n→ Microsoft Store 버전이 아닌 python.org 공식 버전 권장\n→ 설치 후 터미널 재시작',
            };
            return result;
        }
        // 2) pyhwpx 체크
        try {
            const { stdout } = await execFileAsync(pythonExe, ['-c', 'import pyhwpx; print(getattr(pyhwpx, "__version__", "installed"))'], { timeout: 5000 });
            result.pyhwpx = { found: true, version: stdout.trim() };
        }
        catch {
            result.pyhwpx = {
                found: false,
                guide: `pyhwpx 패키지 설치 필요\n→ 감지된 Python: ${result.python.path || pythonExe}\n→ 터미널에서 실행: pip install pyhwpx pywin32`,
            };
            return result;
        }
        // 3) 한글(HWP) COM 등록 체크
        try {
            await execFileAsync(pythonExe, ['-c', 'import win32com.client; o = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject"); o.XHwpDocuments.Close(False); del o'], { timeout: 15000 });
            result.hwp = { found: true };
        }
        catch (err) {
            const msg = err.message || '';
            if (msg.includes('COM class not registered') || msg.includes('gencache')) {
                result.hwp = {
                    found: false,
                    guide: '한글(HWP) 프로그램이 설치되지 않았습니다.\n→ 한컴오피스 한글 설치 필요 (한글 2014 이상)\n→ 설치 후 한글을 한번 실행하여 초기 설정 완료',
                };
            }
            else {
                result.hwp = { found: true };
            }
        }
        // 4) 한글 프로세스 실행 여부 체크
        try {
            const { stdout } = await execFileAsync('tasklist', ['/FI', 'IMAGENAME eq Hwp.exe', '/NH'], { timeout: 5000 });
            result.hwpRunning = stdout.includes('Hwp.exe');
        }
        catch {
            result.hwpRunning = false;
        }
        result.ok = result.python.found && result.pyhwpx.found && result.hwp.found;
        return result;
    }
    // State management
    getState() {
        return {
            pythonRunning: this.pythonRunning,
            currentDocumentPath: this.currentDocumentPath,
            currentDocumentFormat: this.currentDocumentFormat,
            lastAnalysis: this.lastAnalysis,
            uptimeMs: Date.now() - this.startTime,
        };
    }
    setCurrentDocument(filePath) {
        this.currentDocumentPath = filePath;
        if (filePath) {
            const ext = path.extname(filePath).toLowerCase();
            this.currentDocumentFormat = ext === '.hwpx' ? 'HWPX' : 'HWP';
        }
        else {
            this.currentDocumentFormat = null;
        }
    }
    getCurrentDocument() {
        return this.currentDocumentPath;
    }
    getCurrentDocumentFormat() {
        return this.currentDocumentFormat;
    }
    getCachedAnalysis() {
        return this.lastAnalysis;
    }
    setCachedAnalysis(data) {
        this.lastAnalysis = data;
    }
}
