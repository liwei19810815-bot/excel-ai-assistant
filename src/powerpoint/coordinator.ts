/**
 * 所有 PowerPoint.run 调用的统一入口，串行执行。
 *
 * 和 src/excel/coordinator.ts 是同一个理由：Office.js 的 RequestContext
 * 不是线程安全的，并发的 PowerPoint.run 会互相打断 context.sync()。
 */

let tail: Promise<unknown> = Promise.resolve();

export function runPowerPoint<T>(fn: (ctx: PowerPoint.RequestContext) => Promise<T>): Promise<T> {
  const next = tail.then(
    () => PowerPoint.run(fn),
    () => PowerPoint.run(fn), // 前一个任务失败不应阻塞后续
  );
  tail = next.catch(() => undefined);
  return next;
}

/**
 * 运行时探测 PowerPointApi 能力。
 * manifest 只声明 1.1 基线，高版本能力在调用点按需检查。
 */
export function isSupported(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported('PowerPointApi', version);
  } catch {
    return false;
  }
}

/** 能力不足时抛出可读错误，由工具层转成给模型看的提示 */
export function requireApi(version: string, feature: string): void {
  if (!isSupported(version)) {
    throw new Error(
      `当前 PowerPoint 版本不支持「${feature}」（需要 PowerPointApi ${version}）。` +
        `请改用其它方式实现，或告知用户升级 PowerPoint。`,
    );
  }
}
