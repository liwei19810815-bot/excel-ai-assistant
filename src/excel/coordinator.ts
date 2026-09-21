/**
 * 所有 Excel.run 调用的统一入口，串行执行。
 *
 * 为什么必须串行：Office.js 的 RequestContext 不是线程安全的，多个 Excel.run
 * 并发时 context.sync() 会互相打断，表现为随机丢写入或报 InvalidRequestContext。
 * Agent 循环里可能连续触发多个工具，故全部经此队列排队。
 */

let tail: Promise<unknown> = Promise.resolve();

export function runExcel<T>(fn: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  const next = tail.then(
    () => Excel.run(fn),
    () => Excel.run(fn), // 前一个任务失败不应阻塞后续
  );
  tail = next.catch(() => undefined);
  return next;
}

/**
 * 运行时探测 ExcelApi 能力。
 * manifest 只声明 1.1 基线，高版本能力在调用点按需检查，
 * 保证 Mac / Web / 旧版桌面 Excel 也能装上并降级可用。
 */
export function isSupported(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported('ExcelApi', version);
  } catch {
    return false;
  }
}

/** 能力不足时抛出可读错误，由工具层转成给模型看的提示 */
export function requireApi(version: string, feature: string): void {
  if (!isSupported(version)) {
    throw new Error(
      `当前 Excel 版本不支持「${feature}」（需要 ExcelApi ${version}）。` +
        `请改用其它方式实现，或告知用户升级 Excel。`,
    );
  }
}
