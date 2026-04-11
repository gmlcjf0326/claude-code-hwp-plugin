/**
 * HWP Python Bridge for MCP Server
 * Adapted from electron/services/hwp-bridge.ts — no Electron dependencies.
 * All logging via console.error() to protect stdout (MCP JSON-RPC).
 */
import { spawn, execFile } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
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
    // v0.7.4.7: 세션 내 1회만 first-run auto-install 시도
    firstRunCompleted = false;
    getPythonScriptDir() {
        // npm 패키지: dist/hwp-bridge.js → ../python (패키지 내 python/)
        const npmPath = path.resolve(__dirname, '../python');
        if (fs.existsSync(path.join(npmPath, 'hwp_service.py')))
            return npmPath;
        // 개발 환경: mcp-server/dist/hwp-bridge.js → ../../python (프로젝트 루트)
        return path.resolve(__dirname, '../../python');
    }
    /**
     * v0.7.4.5: Python 실행 파일 탐색.
     * - PYTHON_PATH 환경변수가 있으면 그것 사용 (사용자 override)
     * - Windows: py launcher 로 Python 3.13 LTS 우선 (pywin32 + 3.14 gen_py 호환성 이슈 회피)
     * - POSIX: 기본 python3
     *
     * 반환 형식: { exe, preArgs } — spawn(exe, [...preArgs, ...args])
     */
    findPython() {
        if (process.env.PYTHON_PATH) {
            return { exe: process.env.PYTHON_PATH, preArgs: [] };
        }
        if (process.platform === 'win32') {
            // py launcher 가 설치되어 있으면 -3.13 으로 특정 버전 사용
            // (Python 3.14 + pywin32 311 조합의 gen_py 최초 import 수 시간 hang 회피)
            // py launcher 가 없거나 3.13 미설치 시 ENOENT → 사용자가 PYTHON_PATH 설정
            return { exe: 'py', preArgs: ['-3.13'] };
        }
        return { exe: 'python3', preArgs: [] };
    }
    async ensureRunning() {
        if (this.process && this.pythonRunning)
            return;
        // 동시 재시작 방지: 이미 시작 중이면 기다림
        if (this.startPromise)
            return this.startPromise;
        // v0.7.4.7: 세션 최초 호출 시 의존성 자동 설치 (sentinel 기반 1회만)
        if (!this.firstRunCompleted) {
            this.firstRunCompleted = true;
            try {
                await this.firstRunSetup();
            }
            catch (err) {
                // 자동 설치 실패는 하드 에러로 throw — 사용자에게 명확한 가이드 전달
                throw err;
            }
        }
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
        const { exe: pythonExe, preArgs } = this.findPython();
        const scriptPath = path.join(this.getPythonScriptDir(), 'hwp_service.py');
        console.error(`[HWP MCP Bridge] Starting Python: ${pythonExe} ${preArgs.join(' ')} ${scriptPath}`.trim());
        this.process = spawn(pythonExe, [...preArgs, scriptPath], {
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
    // ── v0.7.4.7: 세션 최초 호출 시 1회 사전 자동 설치 ──
    getSentinelPath() {
        return path.join(os.homedir(), '.hwp_studio_state', 'deps_installed.flag');
    }
    async firstRunSetup() {
        // Opt-out: 사용자가 수동 제어를 원하면 HWP_SKIP_AUTO_INSTALL=1
        if (process.env.HWP_SKIP_AUTO_INSTALL === '1') {
            console.error('[HWP Bridge] HWP_SKIP_AUTO_INSTALL=1 — first-run auto-install 스킵');
            return;
        }
        // 이전에 이미 성공적으로 설치된 적이 있으면 skip
        const sentinelPath = this.getSentinelPath();
        if (fs.existsSync(sentinelPath)) {
            return;
        }
        console.error('[HWP Bridge] First-run: 의존성 상태 확인 중...');
        let prereq;
        try {
            prereq = await this.checkPrerequisites();
        }
        catch (err) {
            console.error(`[HWP Bridge] First-run checkPrerequisites 실패: ${err.message}`);
            return; // non-fatal — 실제 spawn 시 에러가 더 명확함
        }
        // Python 이 없으면 auto-install 불가 — 사용자가 먼저 Python 3.13 설치해야 함
        if (!prereq.python.found) {
            console.error('[HWP Bridge] Python 미설치 — auto-install 스킵. 사용자가 먼저 Python 3.13 설치 필요.');
            return;
        }
        const pyhwpxMissing = !prereq.pyhwpx.found;
        const coreDepsMissing = prereq.deps ? !prereq.deps.allCoreReady : false;
        const needsInstall = pyhwpxMissing || coreDepsMissing;
        if (!needsInstall) {
            // 이미 설치 상태가 정상 — sentinel 기록 후 skip
            this.writeSentinel(sentinelPath, {
                mode: 'already_ready',
                python_version: prereq.python.version,
            });
            return;
        }
        const missingList = [];
        if (pyhwpxMissing)
            missingList.push('pyhwpx');
        if (prereq.deps?.missingCore)
            missingList.push(...prereq.deps.missingCore);
        console.error(`[HWP Bridge] First-run: 미설치 의존성 감지 [${missingList.join(', ')}] — core_only 자동 설치 시작 (~2-5분)...`);
        const result = await this.installDeps({ mode: 'core_only', timeoutMs: 600000 });
        if (result.ok) {
            console.error(`[HWP Bridge] First-run auto-install 완료 (${result.durationSeconds}s)`);
            this.writeSentinel(sentinelPath, {
                mode: 'core_only',
                duration_seconds: result.durationSeconds,
                python_version: prereq.python.version,
                installed: result.installed,
            });
        }
        else {
            console.error(`[HWP Bridge] First-run auto-install 실패: ${result.error}`);
            // Sentinel 기록 안 함 — 다음 세션에서 재시도
            throw new Error(`HWP Studio 첫 실행 의존성 자동 설치 실패.\n` +
                `에러: ${result.error}\n\n` +
                `해결 방법:\n` +
                `  1. 수동 설치: py -3.13 -m pip install -r mcp-server/python/requirements.txt\n` +
                `  2. Python 3.13 LTS 설치 확인: https://www.python.org/downloads/release/python-3130/\n` +
                `  3. 프록시 환경 확인 (pip 가 PyPI 접근 가능해야 함)\n` +
                `  4. 자동 설치를 완전히 스킵하려면 HWP_SKIP_AUTO_INSTALL=1 환경변수 설정\n\n` +
                `stderr 요약: ${(result.stderr || '').slice(-1000)}`);
        }
    }
    writeSentinel(sentinelPath, data) {
        try {
            fs.mkdirSync(path.dirname(sentinelPath), { recursive: true });
            fs.writeFileSync(sentinelPath, JSON.stringify({
                timestamp: new Date().toISOString(),
                version: '0.7.4.7',
                ...data,
            }, null, 2), 'utf8');
        }
        catch (err) {
            console.error(`[HWP Bridge] Sentinel 기록 실패: ${err.message}`);
        }
    }
    // ── v0.7.4.6: 의존성 자동 설치 ──
    async installDeps(opts = {}) {
        const { exe: pythonExe, preArgs } = this.findPython();
        const scriptDir = this.getPythonScriptDir();
        const reqPath = path.join(scriptDir, 'requirements.txt');
        const startedAt = Date.now();
        const mode = opts.mode ?? 'all';
        // paddlepaddle 다운로드가 ~500MB 이므로 기본 20분 timeout
        const timeout = opts.timeoutMs ?? 1200000;
        let args;
        let installed = [];
        if (mode === 'core_only') {
            // paddlepaddle/paddleocr 제외, 버전 고정 (requirements.txt 와 동일)
            installed = [
                // v0.7.4.9: Python 3.13 호환성 위해 일부 strict pin 해제 (>= 로 변경)
                'pyhwpx==1.7.2', 'pywin32==308', 'PyMuPDF==1.24.11',
                'openpyxl==3.1.5', 'python-docx==1.1.2',
                'pdfplumber>=0.11.4', 'Pillow>=10.4.0',
                'opencv-python-headless>=4.10.0.84', 'numpy>=1.26.4',
            ];
            args = [...preArgs, '-m', 'pip', 'install', ...installed];
        }
        else {
            if (!fs.existsSync(reqPath)) {
                return {
                    ok: false,
                    command: '',
                    stdout: '',
                    stderr: '',
                    durationSeconds: 0,
                    error: `requirements.txt not found: ${reqPath}`,
                };
            }
            args = [...preArgs, '-m', 'pip', 'install', '-r', reqPath];
            installed = ['(from requirements.txt)'];
        }
        const command = `${pythonExe} ${args.join(' ')}`;
        try {
            const { stdout, stderr } = await execFileAsync(pythonExe, args, {
                timeout,
                maxBuffer: 20 * 1024 * 1024, // pip 출력이 클 수 있음
            });
            const durationSeconds = Math.round((Date.now() - startedAt) / 1000);
            // 설치 후 재검증
            const verify = await this.checkPrerequisites();
            // v0.7.4.7: 수동 설치 성공 시 sentinel 도 기록 → 다음 ensureRunning() 에서 auto-install 스킵
            if (verify.pyhwpx.found && (mode === 'all' || (verify.deps && verify.deps.allCoreReady))) {
                this.writeSentinel(this.getSentinelPath(), {
                    mode,
                    duration_seconds: durationSeconds,
                    python_version: verify.python.version,
                    trigger: 'manual_install_deps',
                });
            }
            return {
                ok: true,
                command,
                stdout: stdout.slice(-8000), // 출력 끝부분만 (시작 부분은 download 진행 상황이라 덜 유용)
                stderr: stderr.slice(-4000),
                durationSeconds,
                installed,
                verified: verify.deps,
            };
        }
        catch (err) {
            const e = err;
            const durationSeconds = Math.round((Date.now() - startedAt) / 1000);
            const isTimeout = e.code === 'ETIMEDOUT' || (e.message ?? '').includes('timed out');
            return {
                ok: false,
                command,
                stdout: (e.stdout ?? '').slice(-8000),
                stderr: (e.stderr ?? '').slice(-4000),
                durationSeconds,
                error: isTimeout
                    ? `pip install 이 ${Math.round(timeout / 60000)}분 이내에 완료되지 않았습니다. 프록시 환경 또는 패키지 크기(특히 paddlepaddle ~500MB) 때문일 수 있습니다. 더 긴 timeoutMs 로 재시도하거나 수동으로 설치하세요.`
                    : (e.message ?? String(err)),
            };
        }
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
        const { exe: pythonExe, preArgs } = this.findPython();
        // 1) Python 체크 — 경로 + 버전 + Microsoft Store 감지 + 3.14 hang 경고
        try {
            const { stdout } = await execFileAsync(pythonExe, [...preArgs, '-c', 'import sys; print(sys.version.split()[0]); print(sys.executable)'], { timeout: 10000 });
            const lines = stdout.trim().split(/\r?\n/);
            const ver = lines[0];
            const exePath = lines[1] || '';
            const isStorePython = exePath.includes('WindowsApps');
            // v0.7.4.5: Python 3.14+ 에서 pywin32 gen_py 최초 import 가 수 시간 hang 하는 이슈
            const majorMinor = ver.split('.').slice(0, 2).join('.');
            const isUnstablePython = majorMinor === '3.14' || parseFloat(majorMinor) >= 3.15;
            let pythonGuide;
            if (isStorePython) {
                pythonGuide = `Microsoft Store Python 감지 (${exePath}). pyhwpx가 인식되지 않을 수 있습니다.\n→ python.org에서 Python 3.13 LTS 재설치를 권장합니다: https://www.python.org/downloads/`;
            }
            else if (isUnstablePython) {
                pythonGuide = `⚠️ Python ${ver} 감지 — pywin32 gen_py 캐시 최초 생성이 수 시간 hang 할 수 있습니다 (알려진 이슈).\n→ Python 3.13 LTS 권장: https://www.python.org/downloads/release/python-3130/\n→ 또는 PYTHON_PATH 환경변수로 3.13 경로 명시`;
            }
            result.python = {
                found: true,
                version: ver,
                path: exePath,
                guide: pythonGuide,
            };
        }
        catch {
            result.python = {
                found: false,
                guide: 'Python 3.13 LTS 설치 필요\n→ https://www.python.org/downloads/release/python-3130/ 에서 설치\n→ 설치 시 "Add Python to PATH" 반드시 체크\n→ Microsoft Store 버전이 아닌 python.org 공식 버전 권장\n→ 설치 후 터미널 재시작\n→ 또는 PYTHON_PATH 환경변수로 Python 3.13 경로 직접 지정',
            };
            return result;
        }
        // 2) pyhwpx 체크 (v0.7.4.5: timeout 30s 로 확장, timeout 과 ImportError 구분)
        try {
            const { stdout } = await execFileAsync(pythonExe, [...preArgs, '-c', 'import pyhwpx; print(getattr(pyhwpx, "__version__", "installed"))'], { timeout: 30000 });
            result.pyhwpx = { found: true, version: stdout.trim() };
        }
        catch (err) {
            const errMsg = err.message || '';
            const isTimeout = errMsg.includes('ETIMEDOUT') || errMsg.includes('timed out') || errMsg.includes('timeout');
            result.pyhwpx = {
                found: false,
                guide: isTimeout
                    ? `⚠️ pyhwpx import 가 30초 이상 걸렸습니다 — Python ${result.python.version || '?'} 에서 pywin32 gen_py 최초 생성 hang 이슈일 수 있습니다.\n→ Python 3.13 LTS 다운그레이드 권장\n→ 또는 별도 터미널에서 python -c "import pyhwpx" 를 끝까지 대기 후 재시도`
                    : `pyhwpx 패키지 설치 필요\n→ 감지된 Python: ${result.python.path || pythonExe}\n→ 터미널에서 실행: pip install pyhwpx pywin32`,
            };
            return result;
        }
        // 2.5) v0.7.4.6: 확장 의존성 체크 (find_spec 기반 — 실제 import 안 함, 빠름)
        try {
            const depsScript = 'import json,importlib.util as iu;' +
                "print(json.dumps({n:iu.find_spec(m) is not None for n,m in [" +
                "('pdfplumber','pdfplumber')," +
                "('Pillow','PIL')," +
                "('opencv','cv2')," +
                "('numpy','numpy')," +
                "('PyMuPDF','fitz')," +
                "('paddlepaddle','paddle')," +
                "('paddleocr','paddleocr')" +
                "]}))";
            const { stdout } = await execFileAsync(pythonExe, [...preArgs, '-c', depsScript], { timeout: 10000 });
            const d = JSON.parse(stdout.trim());
            const core = {
                pdfplumber: d.pdfplumber,
                Pillow: d.Pillow,
                opencv: d.opencv,
                numpy: d.numpy,
                PyMuPDF: d.PyMuPDF,
            };
            const ocr = {
                paddlepaddle: d.paddlepaddle,
                paddleocr: d.paddleocr,
            };
            const missingCore = Object.entries(core).filter(([, v]) => !v).map(([k]) => k);
            const missingOcr = Object.entries(ocr).filter(([, v]) => !v).map(([k]) => k);
            result.deps = {
                core,
                ocr,
                missingCore,
                missingOcr,
                allCoreReady: missingCore.length === 0,
                allOcrReady: missingOcr.length === 0,
            };
        }
        catch {
            // 의존성 체크 실패는 non-fatal — pyhwpx 자체는 작동할 수 있음
        }
        // 3) 한글(HWP) 설치 체크 — COM Dispatch 대신 파일 존재 + pyhwpx import 확인
        // (COM Dispatch는 빈 한글 문서를 열어버리는 부작용이 있으므로 제거)
        try {
            const { stdout } = await execFileAsync(pythonExe, [...preArgs, '-c',
                'import os; paths = [r"C:\\Program Files\\Hancom", r"C:\\Program Files (x86)\\Hancom"]; ' +
                    'found = any(os.path.isdir(p) for p in paths); ' +
                    'print("installed" if found else "not_found")'
            ], { timeout: 5000 });
            if (stdout.trim() === 'installed') {
                result.hwp = { found: true };
            }
            else {
                // 폴더가 없어도 pyhwpx가 COM을 찾을 수 있으므로 pyhwpx import 성공이면 OK
                result.hwp = { found: true, guide: '한컴 설치 폴더를 찾을 수 없지만, pyhwpx가 설치되어 있으므로 동작할 수 있습니다.' };
            }
        }
        catch {
            result.hwp = {
                found: false,
                guide: '한글(HWP) 프로그램이 설치되지 않았습니다.\n→ 한컴오피스 한글 설치 필요 (한글 2014 이상)\n→ 설치 후 한글을 한번 실행하여 초기 설정 완료',
            };
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
