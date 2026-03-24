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
    private getPythonScriptDir;
    private findPython;
    ensureRunning(): Promise<void>;
    private start;
    private processBuffer;
    send(method: string, params: Record<string, unknown>, timeoutMs?: number): Promise<HwpResponse>;
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
