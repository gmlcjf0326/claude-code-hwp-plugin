import { registerDocumentTools } from './tools/document-tools.js';
import { registerAnalysisTools } from './tools/analysis-tools.js';
import { registerEditingTools } from './tools/editing-tools.js';
import { registerCompositeTools } from './tools/composite-tools.js';
import { registerResources } from './resources/document-resources.js';
import { registerPrompts } from './prompts/hwp-prompts.js';
/**
 * MCP 서버 도구 등록.
 *
 * toolset별 도구 구성 (총 85개+, 2026-03-23 최종):
 * - minimal (15개): 문서관리 5 + 분석 7 + 편집 3. 토큰 절약.
 * - standard (85개+): 전체. 문서(5) + 분석(16) + 편집(51) + 복합(20). (기본값)
 * - full: standard와 동일. 향후 확장 시 분리.
 */
export function setupServer(server, bridge, toolset = 'standard') {
    // 문서 관리 + 분석: 모든 toolset에 포함
    registerDocumentTools(server, bridge);
    registerAnalysisTools(server, bridge, toolset);
    // 편집 도구: minimal에서는 핵심만
    registerEditingTools(server, bridge, toolset);
    // 복합 도구: minimal에서는 제외 (standard 이상)
    if (toolset !== 'minimal') {
        registerCompositeTools(server, bridge);
    }
    // 리소스 + 프롬프트: 모든 toolset에 포함
    registerResources(server, bridge);
    registerPrompts(server);
    console.error(`[HWP MCP] 도구셋: ${toolset}`);
}
