interface HwpResponse {
    id: string;
    success: boolean;
    data?: unknown;
    error?: string;
}
export interface BridgeState {
    pythonRunning: boolean;
    currentDocumentPath: string | null;
    currentDocumentFormat: string | null;
    lastAnalysis: unknown | null;
    uptimeMs: number;
}
export interface PrerequisiteResult {
    ok: boolean;
    python: {
        found: boolean;
        version?: string;
        path?: string;
        error?: string;
        guide?: string;
    };
    pyhwpx: {
        found: boolean;
        version?: string;
        error?: string;
        guide?: string;
    };
    hwp: {
        found: boolean;
        error?: string;
        guide?: string;
    };
    hwpRunning: boolean;
    os: {
        ok: boolean;
        platform: string;
        error?: string;
    };
    deps?: {
        core: {
            pdfplumber: boolean;
            Pillow: boolean;
            opencv: boolean;
            numpy: boolean;
            PyMuPDF: boolean;
        };
        ocr: {
            paddlepaddle: boolean;
            paddleocr: boolean;
        };
        missingCore: string[];
        missingOcr: string[];
        allCoreReady: boolean;
        allOcrReady: boolean;
    };
}
export interface InstallDepsResult {
    ok: boolean;
    command: string;
    stdout: string;
    stderr: string;
    durationSeconds: number;
    error?: string;
    installed?: string[];
    verified?: PrerequisiteResult['deps'];
}
export declare class HwpBridge {
    private process;
    private requestId;
    private pending;
    private buffer;
    private readonly MAX_BUFFER_SIZE;
    private startPromise;
    private pythonRunning;
    private currentDocumentPath;
    private currentDocumentFormat;
    private lastAnalysis;
    private lastError;
    private startTime;
    private firstRunCompleted;
    private getPythonScriptDir;
    /**
     * v0.7.4.5: Python 실행 파일 탐색.
     * - PYTHON_PATH 환경변수가 있으면 그것 사용 (사용자 override)
     * - Windows: py launcher 로 Python 3.13 LTS 우선 (pywin32 + 3.14 gen_py 호환성 이슈 회피)
     * - POSIX: 기본 python3
     *
     * 반환 형식: { exe, preArgs } — spawn(exe, [...preArgs, ...args])
     */
    private findPython;
    ensureRunning(): Promise<void>;
    private start;
    private processBuffer;
    send(method: string, params: Record<string, unknown>, timeoutMs?: number): Promise<HwpResponse>;
    private getSentinelPath;
    private firstRunSetup;
    private writeSentinel;
    installDeps(opts?: {
        mode?: 'all' | 'core_only';
        timeoutMs?: number;
    }): Promise<InstallDepsResult>;
    shutdown(): Promise<void>;
    private rejectAllPending;
    checkPrerequisites(): Promise<PrerequisiteResult>;
    getState(): BridgeState;
    setCurrentDocument(filePath: string | null): void;
    getCurrentDocument(): string | null;
    getCurrentDocumentFormat(): string | null;
    getCachedAnalysis(): unknown | null;
    setCachedAnalysis(data: unknown): void;
}
export {};
