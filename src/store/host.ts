/**
 * 当前运行在哪个 Office 宿主里。
 *
 * ============================================================
 * 【为什么需要这个】
 * ============================================================
 * 同一个任务窗格 bundle 现在要同时给 Excel 和 PowerPoint 用（manifest
 * 里加了两个 Host 块），但两边能调的工具完全不同——Excel 工具调
 * Excel.run，PPT 工具调 PowerPoint.run，装错宿主直接报错。模型看到
 * 的工具列表必须按实际宿主过滤，这就需要一个"现在到底在哪个宿主里"
 * 的全局状态，和 sidecar.ts 的 SidecarStatus 是同一个套路。
 *
 * 【默认必须是 'unknown'，不能猜】。启动探测完成之前，任何读到这个值
 * 的代码都应该按"两边工具都不确定能不能用"处理，而不是随便猜一个——
 * 猜错了模型会在一个装错的工具上尝试，报错还看不出原因。
 */

export type HostKind = 'excel' | 'powerpoint' | 'unknown';

/**
 * 读 Office.context.host，对照 Office.HostType 判断当前宿主。
 *
 * 【必须兜异常】：Office.js 未加载完成、或者在无 Office 环境（单元测试、
 * 早期 main.tsx 还没跑完 Office.onReady）时读这个会抛，不能让探测本身
 * 把启动流程带崩。
 */
export function detectHost(): HostKind {
  try {
    const host = Office?.context?.host;
    if (host === Office.HostType.Excel) return 'excel';
    if (host === Office.HostType.PowerPoint) return 'powerpoint';
    return 'unknown';
  } catch {
    return 'unknown';
  }
}

let cachedHost: HostKind = 'unknown';

export function setHost(h: HostKind): void {
  cachedHost = h;
}

export function getHost(): HostKind {
  return cachedHost;
}
