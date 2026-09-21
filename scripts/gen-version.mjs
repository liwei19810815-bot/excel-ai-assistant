/**
 * 构建期生成 public/version.json。
 * 前端启动后轮询该文件（no-cache），发现版本与自身构建版本不一致时提示用户刷新，
 * 实现「部署即生效、用户无感更新」。缓存头配置见 docs/DEPLOY.md。
 */
import { writeFileSync, mkdirSync, readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));

const payload = {
  version: pkg.version,
  buildTime: new Date().toISOString(),
  // sidecar 握手时低于此版本则禁用 sidecar 类工具并提示更新
  minSidecarVersion: '0.1.0',
};

mkdirSync(resolve(root, 'public'), { recursive: true });
writeFileSync(resolve(root, 'public/version.json'), JSON.stringify(payload, null, 2) + '\n');
console.log(`[gen-version] ${payload.version} @ ${payload.buildTime}`);
