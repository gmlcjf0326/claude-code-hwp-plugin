import { build } from 'esbuild';
import { readFileSync, cpSync } from 'fs';
import { execSync } from 'child_process';

const pkg = JSON.parse(readFileSync('./package.json', 'utf8'));

// v0.6.6: mcp-server tsc 자동화 — v0.5.x 수동 워크플로우(tsc → cp dist/tools → bundle)를 단일 build.mjs로 자동화.
// 이전엔 사용자가 mcp-server에서 npm run build 후 수동으로 cp 했어야 함 → drift 발생.
try {
  execSync('npm run build', { cwd: '../mcp-server', stdio: 'inherit' });
  console.log('✅ mcp-server tsc completed (../mcp-server/dist/)');
} catch (e) {
  console.error('⚠️  mcp-server tsc failed:', e.message);
  console.error('   build will continue with existing servers/ files');
}

// v0.6.6: mcp-server/dist → claude-code-hwp-plugin/servers 미러 (TypeScript 컴파일 결과물)
// dist는 src와 1:1 매핑 구조 (rootDir/outDir). servers/bundle.mjs는 dist에 없으므로 안 덮어씀.
try {
  cpSync('../mcp-server/dist', './servers', { recursive: true });
  console.log('✅ servers mirrored from ../mcp-server/dist (TypeScript)');
} catch (e) {
  console.error('⚠️  servers mirror failed:', e.message);
}

// B1 (v0.6.6): mcp-server/python → claude-code-hwp-plugin/python 자동 미러
// 양쪽 미러 드리프트 방지. 빌드 시점에 매번 동기화.
try {
  cpSync('../mcp-server/python', './python', { recursive: true });
  console.log('✅ Python mirror synced from ../mcp-server/python');
} catch (e) {
  console.error('⚠️  Python mirror sync failed:', e.message);
  console.error('   build will continue with existing python/ files');
}

await build({
  entryPoints: ['servers/index.js'],
  bundle: true,
  platform: 'node',
  target: 'node18',
  format: 'esm',
  outfile: 'servers/bundle.mjs',
  // shebang 제거 — node로 직접 실행되므로 불필요
  // Node built-ins are external (not bundled)
  external: [
    'fs', 'path', 'os', 'child_process', 'crypto', 'url', 'util',
    'stream', 'events', 'buffer', 'net', 'http', 'https', 'tls',
    'assert', 'zlib', 'querystring', 'string_decoder', 'worker_threads',
    'node:*',
  ],
  minify: false, // 디버그 용이하도록
  sourcemap: false,
});

console.log('✅ Bundle created: servers/bundle.mjs');
