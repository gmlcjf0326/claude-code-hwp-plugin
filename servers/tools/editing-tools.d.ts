import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { HwpBridge } from '../hwp-bridge.js';
import type { Toolset } from '../server.js';
export declare function registerEditingTools(server: McpServer, bridge: HwpBridge, toolset?: Toolset): void;
