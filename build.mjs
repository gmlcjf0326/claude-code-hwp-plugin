import { build } from 'esbuild';
import { readFileSync } from 'fs';

const pkg = JSON.parse(readFileSync('./package.json', 'utf8'));

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
