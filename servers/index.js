#!/usr/bin/env node
/**
 * HWP Studio MCP Server — Entry Point
 * stdio transport for Claude Code integration.
 * WARNING: console.log() is forbidden — stdout is MCP JSON-RPC only.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { HwpBridge } from './hwp-bridge.js';
import { setupServer } from './server.js';
// --toolset 파라미터 파싱 (minimal | standard | full, 기본: standard)
const toolsetArg = process.argv.find(a => a.startsWith('--toolset'));
const toolset = toolsetArg?.split('=')[1]
    ?? process.argv[process.argv.indexOf('--toolset') + 1]
    ?? 'standard';
const validToolsets = ['minimal', 'standard', 'full'];
const resolvedToolset = validToolsets.includes(toolset)
    ? toolset
    : 'standard';
const bridge = new HwpBridge();
const server = new McpServer({
    name: 'claude-code-hwp-mcp',
    version: '0.3.1',
});
setupServer(server, bridge, resolvedToolset);
const transport = new StdioServerTransport();
await server.connect(transport);
console.error(`[HWP MCP] 서버 시작됨 (toolset: ${resolvedToolset}) — Claude Code에서 HWP 도구 사용 가능`);
// 시작 시 환경 자동 체크 (비동기, 서버 시작을 차단하지 않음)
bridge.checkPrerequisites().then(prereq => {
    if (!prereq.ok) {
        console.error('[HWP MCP] ⚠️ 환경 설정 필요:');
        if (!prereq.os.ok)
            console.error(`  ❌ ${prereq.os.error}`);
        if (!prereq.python.found)
            console.error('  ❌ Python 미설치 → https://www.python.org/downloads/ (PATH 추가 필수)');
        else if (!prereq.pyhwpx.found)
            console.error('  ❌ pyhwpx 미설치 → pip install pyhwpx');
        else if (!prereq.hwp.found)
            console.error('  ❌ 한글(HWP) 미설치 → 한컴오피스 설치 필요');
        console.error('  💡 자세한 진단: hwp_check_setup 도구를 호출하세요');
    }
    else {
        console.error(`[HWP MCP] ✅ 환경 준비 완료 (Python ${prereq.python.version})`);
    }
}).catch(() => { });
process.on('SIGINT', async () => {
    await bridge.shutdown();
    process.exit(0);
});
process.on('SIGTERM', async () => {
    await bridge.shutdown();
    process.exit(0);
});
