import { z } from 'zod';
import { register } from '../registry';
import { probeSidecar, callSidecar, readInjectedSidecarToken } from '../../store/sidecar';

/**
 * Power Query：Office.js 【完全做不到】的一类事。
 *
 * 任务窗格跑在浏览器沙箱里，碰不到 WorkbookQuery / WorkbookConnection，
 * 所以"帮我刷新一下数据源"这种最常见的诉求原本只能回答"做不到"。
 * 这两个工具把它交给本机的 sidecar 去做。
 *
 * ============================================================
 * 【只暴露具体动作，绝不接受代码】
 * ============================================================
 * 这里没有、也不会有"传一段 M 表达式/VBA 进去执行"的口子。
 * 那种通用口子会把风险从"刷新某个查询"放大成"任意代码执行"，
 * 而 sidecar 监听在 127.0.0.1 上，本机任何一个网页都够得着它。
 * 要新增能力就在 sidecar 那边加一条具体路由，在这里加一个具体工具。
 *
 * ============================================================
 * 【sidecar 不在的时候这两个工具根本不该出现在列表里】
 * ============================================================
 * ChatPane 会按启动时探测的结果把它们从工具列表里摘掉，
 * 模型看不到就不会承诺做不到的事。
 * 这里的 sidecar 不可用分支是【兜底】：探测之后进程挂了也可能走到，
 * 那时要给一句人话，而不是抛一个连接错误。
 */

/** sidecar 类工具的名字。ChatPane 用它决定摘掉哪些。 */
export const SIDECAR_TOOL_NAMES = ['list_queries', 'refresh_query'] as const;

const UNAVAILABLE_TEXT =
  '本机助手（sidecar）没有运行，刷新 Power Query 这类操作做不了。' +
  '其余功能不受影响。如果需要这个能力，请联系 IT 安装工具箱的 sidecar 组件。';

async function withSidecar<T>(fn: (status: Awaited<ReturnType<typeof probeSidecar>>, token: string) => Promise<T>) {
  const token = readInjectedSidecarToken();
  const status = await probeSidecar(token);
  if (!status.available) return null;
  return await fn(status, token);
}

register({
  name: 'list_queries',
  description:
    '列出当前工作簿里的所有 Power Query 查询名称。' +
    '在用户提到"数据源""查询""刷新数据"但没说清是哪一个时，先用这个看看有哪些。' +
    '这是只读操作，不会改动任何数据。',
  policy: 'read',
  schema: z.object({}),
  summarize: () => '列出工作簿里的 Power Query 查询',

  async run() {
    const r = await withSidecar(async (status, token) =>
      callSidecar<{ ok: boolean; queries?: string[]; error?: string }>(
        status,
        '/queries',
        token,
      ),
    );

    if (r === null) return { text: UNAVAILABLE_TEXT };
    if (!r.ok) return { text: `读取查询列表失败：${r.error ?? '未知原因'}` };

    const list = r.queries ?? [];
    if (list.length === 0) return { text: '当前工作簿里没有 Power Query 查询。' };
    return { text: `当前工作簿共有 ${list.length} 个查询：${list.join('、')}` };
  },
});

register({
  name: 'refresh_query',
  description:
    '刷新当前工作簿里指定名称的 Power Query 查询（重新从数据源拉取数据）。' +
    '只接受查询的名称，不接受任何代码或 M 表达式。' +
    '如果不确定查询叫什么，先用 list_queries 看一下。',

  // 【定成 mutate:structure，一律强制确认】。
  // 刷新会把表里的数据整片换掉，而且【工具箱的撤销框架覆盖不到它】——
  // 那套快照机制是给 VBA 命令用的，sidecar 这边的改动它不知道。
  // 换句话说这个动作事实上不可撤销，所以必须让用户先点头。
  policy: 'mutate:structure',
  schema: z.object({
    name: z.string().min(1).describe('要刷新的查询名称，必须和工作簿里的完全一致'),
  }),
  summarize: (a) => `刷新 Power Query 查询：${a.name}`,

  async run(args) {
    const r = await withSidecar(async (status, token) =>
      callSidecar<{ ok: boolean; refreshed?: string; error?: string; name?: string }>(
        status,
        '/refresh-query',
        token,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: args.name }),
        },
      ),
    );

    if (r === null) return { text: UNAVAILABLE_TEXT };

    // 【失败必须如实说】。这里最坏的结果是把"没刷新"报成"刷新好了"——
    // 用户以为数据是新的，拿着旧数据往下做。
    if (!r.ok) {
      if (r.error === 'query_not_found') {
        return { text: `工作簿里没有名为「${args.name}」的查询，没有刷新任何东西。可以用 list_queries 看看实际有哪些。` };
      }
      if (r.error === 'excel_not_running') {
        return { text: 'Excel 没有在运行，无法刷新。' };
      }
      if (r.error === 'no_workbook') {
        return { text: '当前没有打开的工作簿，无法刷新。' };
      }
      if (r.error === 'timeout') {
        return { text: '刷新超时了。Excel 可能正弹着对话框，请切过去看一眼。' };
      }
      return { text: `刷新失败：${r.error ?? '未知原因'}` };
    }

    return { text: `已刷新查询「${r.refreshed ?? args.name}」。` };
  },
});
