/**
 * 构建后把源码指纹写进 dist/.build-fingerprint。
 * 检查脚本靠它判断"这份产物是不是这份源码构建的"。
 */
import { writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { sourceFingerprint } from './source-fingerprint.mjs';

const { hash, fileCount } = sourceFingerprint('.');
if (!existsSync('dist')) mkdirSync('dist', { recursive: true });
writeFileSync(
  'dist/.build-fingerprint',
  JSON.stringify({ hash, fileCount, builtAt: new Date().toISOString() }, null, 2) + '\n',
);
console.log(`[fingerprint] ${hash.slice(0, 12)}… (${fileCount} 个源文件)`);
