/**
 * MCP Resources: document status, analysis, text, tables
 */
import { ResourceTemplate } from '@modelcontextprotocol/sdk/server/mcp.js';
export function registerResources(server, bridge) {
    // Server status
    server.resource('status', 'hwp://status', {
        description: 'HWP MCP 서버 상태 (Python 실행 여부, 열린 문서 정보)',
        mimeType: 'application/json',
    }, async (uri) => {
        const state = bridge.getState();
        const status = {
            python_running: state.pythonRunning,
            document_open: state.currentDocumentPath !== null,
            current_file: state.currentDocumentPath,
            current_format: state.currentDocumentFormat,
            server_uptime_seconds: Math.round(state.uptimeMs / 1000),
        };
        return {
            contents: [{ uri: uri.href, text: JSON.stringify(status), mimeType: 'application/json' }],
        };
    });
    // Cached analysis
    server.resource('current-analysis', 'hwp://current/analysis', {
        description: '현재 문서의 캐시된 분석 결과 (표, 필드, 텍스트 포함)',
        mimeType: 'application/json',
    }, async (uri) => {
        const analysis = bridge.getCachedAnalysis();
        if (!analysis) {
            return {
                contents: [{
                        uri: uri.href,
                        text: JSON.stringify({ error: '분석된 문서가 없습니다. hwp_analyze_document를 먼저 실행하세요.' }),
                        mimeType: 'application/json',
                    }],
            };
        }
        return {
            contents: [{ uri: uri.href, text: JSON.stringify(analysis), mimeType: 'application/json' }],
        };
    });
    // Current document text
    server.resource('current-text', 'hwp://current/text', {
        description: '현재 문서의 본문 텍스트',
        mimeType: 'text/plain',
    }, async (uri) => {
        const analysis = bridge.getCachedAnalysis();
        if (!analysis) {
            return {
                contents: [{
                        uri: uri.href,
                        text: '분석된 문서가 없습니다. hwp_analyze_document를 먼저 실행하세요.',
                        mimeType: 'text/plain',
                    }],
            };
        }
        return {
            contents: [{ uri: uri.href, text: analysis.full_text || '', mimeType: 'text/plain' }],
        };
    });
    // Table resource template
    server.resource('table', new ResourceTemplate('hwp://tables/{index}', {
        list: async () => {
            const analysis = bridge.getCachedAnalysis();
            if (!analysis || !analysis.tables) {
                return { resources: [] };
            }
            return {
                resources: analysis.tables.map((_, i) => ({
                    uri: `hwp://tables/${i}`,
                    name: `표 ${i}`,
                    description: `문서의 ${i}번째 표 데이터`,
                    mimeType: 'application/json',
                })),
            };
        },
    }), {
        description: '특정 표 데이터 (인덱스로 접근)',
        mimeType: 'application/json',
    }, async (uri, { index }) => {
        const analysis = bridge.getCachedAnalysis();
        if (!analysis || !analysis.tables) {
            return {
                contents: [{
                        uri: uri.href,
                        text: JSON.stringify({ error: '분석된 문서가 없습니다.' }),
                        mimeType: 'application/json',
                    }],
            };
        }
        const idx = parseInt(String(index), 10);
        if (isNaN(idx) || idx < 0 || idx >= analysis.tables.length) {
            return {
                contents: [{
                        uri: uri.href,
                        text: JSON.stringify({ error: `표 인덱스 ${index}이 범위를 벗어났습니다.` }),
                        mimeType: 'application/json',
                    }],
            };
        }
        return {
            contents: [{
                    uri: uri.href,
                    text: JSON.stringify(analysis.tables[idx]),
                    mimeType: 'application/json',
                }],
        };
    });
}
