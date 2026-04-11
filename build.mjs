import { build } from 'esbuild';
import { readFileSync, cpSync, rmSync } from 'fs';
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
// Phase 0 (2026-04-11): __pycache__ 제외 — stale .pyc 가 cache 로 복사되는 문제 방지
// v0.7.9 post-refactor fix (2026-04-11): cpSync 는 additive (소스에서 삭제된 파일이 destination 에 남음)
//   → Phase 0-9 에서 hwp_editor.py → hwp_editor/ 분할 후 stale hwp_editor.py 가 mirror 에 남아 Python import 충돌 발생.
//   → 사전 rmSync 로 destination 전체 삭제 후 cpSync 실행.
try {
  try {
    rmSync('./python', { recursive: true, force: true });
  } catch (e) {
    // Windows 에서 일부 .pyc 가 lock 된 경우 무시 (filter 가 어차피 제외)
    console.error('[INFO] python/ rmSync partial:', e.message);
  }
  cpSync('../mcp-server/python', './python', {
    recursive: true,
    filter: (src) => !src.includes('__pycache__'),
  });
  console.log('✅ Python mirror synced from ../mcp-server/python (clean rebuild, no __pycache__)');
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
