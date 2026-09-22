import { provision, applyRibbonVisibility } from '../store/provisioning';
import { probeSidecar, setSidecarStatus, type SidecarStatus } from '../store/sidecar';

/**
 * 任务窗格的启动链路。
 *
 * 【为什么单独抽出来】：这段逻辑原先直接写在 main.tsx 里，
 * 而 main.tsx 要操作 DOM、要 import 样式，在无 DOM 的测试环境里跑不起来。
 * 结果就是"provision 之后有没有真的去置灰按钮"这条接线【完全没有测试】——
 * 链路被删掉、顺序写反、传错字段，都不会让任何测试变红。
 *
 * 抽成一个纯函数之后，main.tsx 只负责挂载 + 调这一个函数，
 * 接线本身就能被断言了。
 *
 * 【不抛异常】：启动路径上任何失败都不该把任务窗格搞崩。
 * provision 内部已经把所有失败收敛成 byok + ready，
 * applyRibbonVisibility 自己吞异常。这里再兜一层，返回结果供测试断言。
 */
export interface StartupResult {
  visibility: 0 | 1 | 2;
  mode: 'managed' | 'byok';
  /** 功能区按钮是否真的被更新了（老版本 Office 上会是 false） */
  ribbonUpdated: boolean;
  /** 本机 sidecar 的状态。没装这个组件时 available=false，这是正常情况 */
  sidecar: SidecarStatus;
}

export async function runStartup(): Promise<StartupResult> {
  const r = await provision();

  // 【顺序不能反】：先问到配置，才知道按钮该不该置灰。
  const ribbonUpdated = await applyRibbonVisibility(r.visibility);

  // 探一次本机的 sidecar。
  //
  // 【探不到不是错误】。绝大多数用户没装这个组件，probeSidecar 在
  // 没有令牌时连一个请求都不会发。探不到的唯一后果是
  // sidecar 类工具从工具列表里消失——模型看不到它们，
  // 也就不会承诺做不到的事。
  //
  // 【再兜一层 try】：探测本身出任何意外都不能把启动打断。
  let sidecar: SidecarStatus;
  try {
    sidecar = await probeSidecar();
  } catch {
    sidecar = { available: false, port: 0, version: '', reason: '探测本身出错' };
  }

  // 存下来供工具列表使用。【必须存】——不存的话 ChatPane 永远看到
  // 默认的"不可用"，sidecar 类工具装了也不会出现，而且不会有任何报错。
  setSidecarStatus(sidecar);

  return { visibility: r.visibility, mode: r.mode, ribbonUpdated, sidecar };
}
