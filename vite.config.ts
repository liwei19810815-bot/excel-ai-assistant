import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { resolve } from 'node:path';
import { readFileSync } from 'node:fs';

const pkg = JSON.parse(readFileSync(resolve(__dirname, 'package.json'), 'utf8')) as {
  version: string;
};

/**
 * Office 加载项必须走 HTTPS，且桌面版 Excel 会拒绝自签名证书。
 * 开发时用 office-addin-dev-certs 生成本机受信任证书（`npm run certs`）。
 * 证书不存在时降级为 http，便于纯浏览器调试 UI。
 */
function devCerts() {
  try {
    const home = process.env.USERPROFILE || process.env.HOME || '';
    const dir = resolve(home, '.office-addin-dev-certs');
    return {
      key: readFileSync(resolve(dir, 'localhost.key')),
      cert: readFileSync(resolve(dir, 'localhost.crt')),
    };
  } catch {
    return undefined;
  }
}

export default defineConfig({
  plugins: [react(), tailwindcss()],
  // 前端据此与 version.json 比对，发现不一致则提示用户刷新
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  resolve: {
    alias: { '@': resolve(__dirname, 'src') },
  },
  server: {
    port: 3000,
    https: devCerts(),
  },
  build: {
    outDir: 'dist',
    // 内容哈希文件名 —— 配合入口 HTML 的 no-cache，实现用户无感热更新。
    // 缓存头配置见 docs/DEPLOY.md。
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'src/taskpane/index.html'),
        commands: resolve(__dirname, 'src/commands/commands.html'),
      },
      output: {
        entryFileNames: 'assets/[name].[hash].js',
        chunkFileNames: 'assets/[name].[hash].js',
        assetFileNames: 'assets/[name].[hash][extname]',
      },
    },
  },
});
