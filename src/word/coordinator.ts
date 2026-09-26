/**
 * 所有 Word.run 调用的统一入口，串行执行。
 *
 * 和 src/excel/coordinator.ts / src/powerpoint/coordinator.ts 是同一个
 * 理由：Office.js 的 RequestContext 不是线程安全的，并发的 Word.run 会
 * 互相打断 context.sync()。
 */

let tail: Promise<unknown> = Promise.resolve();

export function runWord<T>(fn: (ctx: Word.RequestContext) => Promise<T>): Promise<T> {
  const next = tail.then(
    () => Word.run(fn),
    () => Word.run(fn), // 前一个任务失败不应阻塞后续
  );
  tail = next.catch(() => undefined);
  return next;
}

/**
 * 运行时探测 WordApi 能力。
 * manifest 只声明 1.1 基线，高版本能力在调用点按需检查。
 */
export function isSupported(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported('WordApi', version);
  } catch {
    return false;
  }
}

/** 能力不足时抛出可读错误，由工具层转成给模型看的提示 */
export function requireApi(version: string, feature: string): void {
  if (!isSupported(version)) {
    throw new Error(
      `当前 Word 版本不支持「${feature}」（需要 WordApi ${version}）。` +
        `请改用其它方式实现，或告知用户升级 Word。`,
    );
  }
}
