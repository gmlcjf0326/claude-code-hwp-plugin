import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { HwpBridge } from './hwp-bridge.js';
export type Toolset = 'minimal' | 'standard' | 'full';
/**
 * MCP 서버 도구 등록.
 *
 * toolset별 도구 구성 (총 85개+, 2026-03-23 최종):
 * - minimal (15개): 문서관리 5 + 분석 7 + 편집 3. 토큰 절약.
 * - standard (85개+): 전체. 문서(5) + 분석(16) + 편집(51) + 복합(20). (기본값)
 * - full: standard와 동일. 향후 확장 시 분리.
 */
export declare function setupServer(server: McpServer, bridge: HwpBridge, toolset?: Toolset): void;
