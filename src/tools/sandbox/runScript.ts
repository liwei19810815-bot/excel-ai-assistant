import { z } from 'zod';
import { register } from '../registry';
import { runExcel } from '../../excel/coordinator';

/**
 * 长尾需求兜底：让模型现写一段 Office.js 代码执行。
 *
 * 安全设计（这是全套工具里风险最高的一个，所以约束最严）：
 *  1. 强制人工确认 —— 用户看到完整源码后才决定是否执行，UI 会高亮代码
 *  2. 受限作用域 —— 只注入 context / Excel / console，不给 fetch / XMLHttpRequest / import
 *  3. 超时中断 —— 默认 30 秒，防止死循环卡住任务窗格
 *  4. 静态检查 —— 拦截明显的外部通信与动态求值
 *
 * 注意：这是纵深防御而非沙箱隔离。new Function 仍在同一 realm 内，
 * 恶意代码理论上可绕过静态检查。真正的边界是「用户确认」这一步，
 * 因此确认弹窗必须完整展示源码，不能折叠或省略。
 */

const SCRIPT_TIMEOUT_MS = 30_000;

/** 明显的越界行为，命中直接拒绝执行 */
const FORBIDDEN = [
  { re: /\bfetch\s*\(/, why: '网络请求' },
  { re: /XMLHttpRequest/, why: '网络请求' },
  { re: /\bWebSocket\b/, why: '网络连接' },
  { re: /\bimport\s*\(/, why: '动态导入' },
  { re: /\beval\s*\(/, why: '动态求值' },
  { re: /\blocalStorage\b|\bsessionStorage\b/, why: '访问本地存储' },
  { re: /\bdocument\b|\bwindow\b/, why: '访问宿主页面' },
];

register({
  name: 'run_script',
  description:
    '执行一段自定义 Office.js 代码，用于结构化工具覆盖不到的长尾需求' +
    '（如按条件批量改值、复杂的跨表计算、不规则区域处理）。' +
    '代码体内可直接使用 `context`（Excel.RequestContext）和 `Excel` 全局对象，支持 await。' +
    '必须自行调用 `await context.sync()` 提交改动。' +
    '用 `return` 返回一个可序列化的结果，它会回传给你。' +
    '禁止网络请求、动态导入、访问页面 DOM。' +
    '【重要】此工具每次执行前都需要用户确认，且改动不会自动创建快照，' +
    '因此优先使用结构化工具，仅在确实无法表达时才用它。',
  policy: 'mutate:structure',
  schema: z.object({
    code: z
      .string()
      .describe(
        '要执行的 JavaScript 代码体（不要包裹 function 或 Excel.run）。' +
          '示例：const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1:A10");' +
          ' r.load("values"); await context.sync(); return r.values;',
      ),
    purpose: z
      .string()
      .describe('用一句中文说明这段代码要做什么，会展示给用户确认'),
  }),
  summarize: (a) => `执行自定义脚本：${a.purpose}`,

  async run(args, ctx) {
    const { code, purpose } = args;

    for (const { re, why } of FORBIDDEN) {
      if (re.test(code)) {
        throw new Error(
          `脚本被拒绝：检测到${why}相关代码（匹配 ${re.source}）。` +
            `run_script 只允许操作工作簿，请改用结构化工具或调整代码。`,
        );
      }
    }

    // 唯一真正的安全边界：用户看过源码后确认
    const approved = await ctx.confirm({
      title: '执行自定义脚本',
      detail:
        `AI 想要执行一段自定义代码：${purpose}\n\n` +
        `此操作不会自动创建快照，撤销需要手动恢复。请确认代码符合预期后再允许。`,
      code,
    });
    if (!approved) {
      return { text: '用户拒绝执行该脚本。请改用结构化工具，或调整方案后再试。' };
    }

    const result = await runExcel(async (excelCtx) => {
      const body = `"use strict";\nreturn (async () => {\n${code}\n})();`;
      // 只把这三个标识符引入作用域，其余全局仍可见但已被静态检查拦截
      const fn = new Function('context', 'Excel', 'console', body) as (
        c: Excel.RequestContext,
        e: typeof Excel,
        l: Console,
      ) => Promise<unknown>;

      const exec = fn(excelCtx, Excel, console);
      const timeout = new Promise<never>((_, reject) =>
        setTimeout(
          () => reject(new Error(`脚本执行超过 ${SCRIPT_TIMEOUT_MS / 1000} 秒已中断。`)),
          SCRIPT_TIMEOUT_MS,
        ),
      );

      const value = await Promise.race([exec, timeout]);
      // 脚本可能改了东西却忘了 sync，这里补一次
      await excelCtx.sync();
      return value;
    });

    return { text: `脚本执行完成。返回值：\n${stringify(result)}` };
  },
});

function stringify(v: unknown): string {
  if (v === undefined) return '(无返回值)';
  try {
    return JSON.stringify(v, null, 2) ?? String(v);
  } catch {
    return String(v);
  }
}
