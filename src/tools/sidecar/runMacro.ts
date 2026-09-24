import { z } from 'zod';
import { register } from '../registry';
import { probeSidecar, callSidecar, readInjectedSidecarToken } from '../../store/sidecar';

/**
 * 调用工作簿里现成的 VBA 宏。
 *
 * 这是 sidecar 三类能力里唯一一个"调用方决定要跑哪段既有逻辑"的工具，
 * 也因此是风险最高的一个——按名字调宏，本质上就是让远端（哪怕是我们
 * 自己的模型）指挥本机执行一段可以做任何事的代码。
 *
 * ============================================================
 * 【不是"任意宏都能调"，是工作簿作者显式暴露的那些】
 * ============================================================
 * sidecar 那边只认 AI_ 开头的 Sub（比如 AI_导出报表）。这个前缀是
 * 工作簿作者的显式选择——把一个宏改名/包一层壳挂上 AI_ 前缀，
 * 就是在说"这段逻辑我认可给 AI 用"。没挂这个前缀的宏，请求在
 * 真的去调用 Excel 之前就会被 sidecar 拒绝，模型这边猜不出来、
 * 也绕不过去。
 *
 * 参数只接受字符串/数字/布尔，不接受代码或表达式——和 refresh_query
 * 是同一个原则。
 */

/** list_macros 单独注册（read），run_macro 在 powerQuery.ts 的 SIDECAR_TOOL_NAMES 之外一并声明。 */
export const RUN_MACRO_TOOL_NAMES = ['list_macros', 'run_macro'] as const;

const UNAVAILABLE_TEXT =
  '本机助手（sidecar）没有运行，调用工作簿里的宏这类操作做不了。' +
  '其余功能不受影响。如果需要这个能力，请联系 IT 安装工具箱的 sidecar 组件。';

async function withSidecar<T>(fn: (status: Awaited<ReturnType<typeof probeSidecar>>, token: string) => Promise<T>) {
  const token = readInjectedSidecarToken();
  const status = await probeSidecar(token);
  if (!status.available) return null;
  return await fn(status, token);
}

register({
  name: 'list_macros',
  description:
    '列出当前工作簿里暴露给 AI 调用的宏（只有名字以 AI_ 开头的宏才会出现在这里）。' +
    '用户提到"跑一下那个宏""执行 XX 宏"但没说清叫什么时，先用这个看看有哪些。' +
    '这是只读操作，不会执行任何宏。' +
    '注意：这条依赖"信任对 VBA 工程对象模型的访问"，用户机上大概率没开，' +
    '枚不出来不代表真的没有宏——如果返回空列表且 trusted 是 false，' +
    '应该直接问用户宏叫什么，而不是断定工作簿里没有可用的宏。',
  policy: 'read',
  schema: z.object({}),
  summarize: () => '列出工作簿里暴露给 AI 的宏',

  async run() {
    const r = await withSidecar(async (status, token) =>
      callSidecar<{ ok: boolean; macros?: string[]; trusted?: boolean; error?: string }>(
        status,
        '/macros',
        token,
      ),
    );

    if (r === null) return { text: UNAVAILABLE_TEXT };
    if (!r.ok) return { text: `读取宏列表失败：${r.error ?? '未知原因'}` };

    const list = r.macros ?? [];
    if (list.length === 0) {
      if (r.trusted === false) {
        return {
          text: '枚举不出宏列表（这台机器可能没开启"信任对 VBA 工程对象模型的访问"）。' +
            '这不代表工作簿里没有可用的宏——如果知道宏名，直接告诉我就行。',
        };
      }
      return { text: '当前工作簿里没有暴露给 AI 的宏（没有 AI_ 开头的 Sub）。' };
    }
    return { text: `当前工作簿暴露给 AI 的宏共 ${list.length} 个：${list.join('、')}` };
  },
});

register({
  name: 'run_macro',
  description:
    '调用当前工作簿里一个已经写好的宏（只能调用名字以 AI_ 开头的宏）。' +
    '只接受宏名和几个基本参数（字符串/数字/布尔），不接受代码。' +
    '不确定有哪些宏可用时，先用 list_macros 看一下。' +
    '【这个工具风险较高】：宏可以做任何事（改数据、存文件、调用外部程序等），' +
    '而且执行前不知道它具体会做什么，所以每次调用都需要用户明确确认。',

  // 【mutate:structure，一律强制确认】。这条比 refresh_query 更要谨慎——
  // 刷新查询至少知道它"只是拉数据"，而宏可以做任何事，我们完全不知道
  // 它内部会改什么、碰不碰文件系统。唯一的安全边界是"用户点头"。
  policy: 'mutate:structure',
  schema: z.object({
    name: z
      .string()
      .min(1)
      .describe('要调用的宏名，必须以 AI_ 开头，且和工作簿里的完全一致'),
    args: z
      .array(z.union([z.string(), z.number(), z.boolean()]))
      .optional()
      .describe('传给宏的参数，按顺序对应宏的形参；只能是字符串/数字/布尔，不能是代码'),
  }),
  summarize: (a) =>
    a.args && a.args.length > 0
      ? `调用宏：${a.name}（参数：${a.args.map((x) => JSON.stringify(x)).join('、')}）`
      : `调用宏：${a.name}`,

  async run(args) {
    // 【前端也做一次前缀校验】，不是因为信任前端——sidecar 那边才是
    // 真正的防线——而是为了在请求发出去之前就给出清楚的中文提示，
    // 不用等一个 400 从网络那头绕回来才告诉用户"名字不对"。
    if (!/^AI_[A-Za-z0-9_]+$/.test(args.name)) {
      return {
        text: `宏名「${args.name}」不符合要求：只能调用以 AI_ 开头的宏。` +
          '这是工作簿作者显式暴露给 AI 的宏，不是任意宏都能调用。',
      };
    }

    // ============================================================
    // 【两步走：先换确认票，再真正调用——但这不是身份边界】
    // ============================================================
    // 走到这里说明 mutate:structure 的确认弹窗已经让用户点过"允许"了
    // （见 agent/loop.ts：策略确认在 tool.run() 之前完成）。这里先调
    // /confirm-macro 换一张绑定"这个宏名 + 这些参数"的一次性票，
    // /run-macro 必须带着它才会真的执行。
    //
    // 【重要】：/confirm-macro 认的还是同一枚 sidecar 令牌，不是什么
    // "只有走过确认流程才能拿到"的凭据——任何持有该令牌的调用方都能
    // 自己换票（这是两轮独立安全评审才厘清的：纯软件层做不出"证明
    // 用户点过确认"这件事，区分"任务窗格代点的"和"脚本自己发的"本质上
    // 需要浏览器脚本无法伪造的东西，比如原生对话框——团队和用户讨论后
    // 决定不引入原生弹窗，转而明确"持有 sidecar 令牌 = 拥有执行宏的
    // 权限"）。这套两步流程仍然有真实价值：单次性消费挡得住重放捕获的
    // 请求，绑定参数减少"手滑传错参"被悄悄放行的窗口，短时效减少令牌
    // 泄露后的可利用时间——但不要把它当成"确认弹窗被 sidecar 强制执行"
    // 的证据。
    //
    // 只探测一次 sidecar 状态，两次调用共用——不用重复走一遍端口探测。
    const outcome = await withSidecar(async (status, token) => {
      const confirmResult = await callSidecar<{ ok: boolean; confirmToken?: string; error?: string }>(
        status,
        '/confirm-macro',
        token,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: args.name, args: args.args ?? [] }),
        },
      );
      if (!confirmResult.ok || !confirmResult.confirmToken) {
        return { text: describeSidecarError(confirmResult.error, args.name) };
      }

      return callSidecar<{ ok: boolean; name?: string; result?: string; error?: string; message?: string }>(
        status,
        '/run-macro',
        token,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: args.name,
            args: args.args ?? [],
            confirmToken: confirmResult.confirmToken,
          }),
        },
      );
    });

    if (outcome === null) return { text: UNAVAILABLE_TEXT };
    if ('text' in outcome) return outcome;
    const r = outcome;

    // 【失败必须如实说】。这里最坏的结果是把"宏没跑起来"报成"跑完了"——
    // 用户以为数据已经处理好了，其实什么都没发生。
    if (!r.ok) {
      if (r.error === 'macro_failed') {
        return { text: `宏「${args.name}」执行失败：${r.message ?? '未知原因'}` };
      }
      return { text: describeSidecarError(r.error, args.name) };
    }

    const resultText = r.result ? `，返回：${r.result}` : '';
    return { text: `已调用宏「${r.name ?? args.name}」${resultText}` };
  },
});

/** run-macro / confirm-macro 共用的错误码 → 中文提示。 */
function describeSidecarError(error: string | undefined, macroName: string): string {
  switch (error) {
    case 'macro_not_allowed':
      return `宏名「${macroName}」不符合要求：只能调用以 AI_ 开头的宏。`;
    case 'excel_not_running':
      return 'Excel 没有在运行，无法调用宏。';
    case 'no_workbook':
      return '当前没有打开的工作簿，无法调用宏。';
    case 'timeout':
      return '调用超时了。Excel 可能正弹着对话框，请切过去看一眼。';
    case 'confirm_required':
    case 'confirm_invalid':
      return '确认已失效（可能是重复调用或参数发生了变化），请重新发起这个操作。';
    case 'missing_name':
      return '没有指定宏名，无法调用。';
    case 'bad_args':
      return '传给宏的参数格式不对：只能是字符串/数字/布尔组成的数组。';
    case 'com_failed':
      return '和 Excel 通信时出错了，可以重试一次；如果一直失败，请检查 Excel 是否正常。';
    case 'no_result':
      return '没有拿到执行结果，无法确认宏是否跑成功了，请去 Excel 里核实一下。';
    default:
      return `调用失败：${error ?? '未知原因'}`;
  }
}
