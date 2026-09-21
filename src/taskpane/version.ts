/**
 * 版本协商 —— 支撑「部署即生效、用户无感更新」。
 *
 * 前端资源带内容哈希可长缓存，但入口 HTML 必须 no-cache，否则 WebView2
 * 会把旧版本钉死。这里额外轮询 version.json 作为兜底：即便用户一直开着
 * 任务窗格不刷新，也能收到更新提示。缓存头配置见 docs/DEPLOY.md。
 */

export const APP_VERSION = __APP_VERSION__;

let notified = false;
const listeners = new Set<(v: string) => void>();

export function onUpdateAvailable(fn: (v: string) => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export async function checkForUpdate(manual = false): Promise<void> {
  try {
    const res = await fetch(`${import.meta.env.BASE_URL}version.json`, {
      cache: 'no-cache',
    });
    if (!res.ok) return;
    const { version } = (await res.json()) as { version: string };

    if (version !== APP_VERSION && (!notified || manual)) {
      notified = true;
      listeners.forEach((fn) => fn(version));
    }
  } catch {
    // 离线或文件缺失时静默忽略，不打扰用户
  }
}

/** 启动后每 10 分钟查一次，长时间开着面板的用户也能拿到更新 */
export function startUpdatePolling(): void {
  void checkForUpdate();
  setInterval(() => void checkForUpdate(), 10 * 60 * 1000);
}
