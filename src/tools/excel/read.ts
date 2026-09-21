import { z } from 'zod';
import { register } from '../registry';
import { runExcel } from '../../excel/coordinator';
import { buildBlueprint, renderBlueprint } from '../../excel/blueprint';
import { toMarkdown, toCsv, sampleRows, type CellValue } from '../../excel/format';

/** 单次读取的单元格上限，超出则采样降级 */
const MAX_CELLS = 2000;

register({
  name: 'get_workbook_overview',
  description:
    '获取工作簿结构蓝图：所有工作表名称、数据区域、行列数、推断的表头，以及 Excel 表和图表清单。' +
    '不返回单元格数据本身。每轮对话已自动注入该信息，仅在你做完结构性改动后需要重新确认时调用。',
  policy: 'read',
  schema: z.object({}),
  summarize: () => '读取工作簿结构',
  async run() {
    const bp = await buildBlueprint();
    return { text: renderBlueprint(bp) };
  },
});

register({
  name: 'get_selection',
  description: '获取用户当前选中的区域地址及其内容。',
  policy: 'read',
  schema: z.object({}),
  summarize: () => '读取当前选区',
  async run() {
    const result = await runExcel(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(['address', 'values', 'rowCount', 'columnCount']);
      await ctx.sync();
      return {
        address: range.address,
        values: range.values as CellValue[][],
        cells: range.rowCount * range.columnCount,
      };
    });

    if (result.cells > MAX_CELLS) {
      return {
        text:
          `当前选区 ${result.address} 共 ${result.cells} 个单元格，超出单次读取上限。\n` +
          `请用 read_range 分块读取。`,
      };
    }

    return {
      text: `当前选区：${result.address}\n\n${toMarkdown(result.values, result.address)}`,
    };
  },
});

register({
  name: 'read_range',
  description:
    '读取指定区域的内容。address 可以是 "A1:D20"（默认当前表）或 "Sheet2!A1:D20"。' +
    '不确定数据范围时可传 "used"，表示读取该表的已用区域。' +
    'mode=markdown 返回带行列号的表格（默认，便于你精确引用单元格）；' +
    'csv 更紧凑适合大块数据；detailed 额外返回公式与数字格式。',
  policy: 'read',
  schema: z.object({
    address: z
      .string()
      .describe('区域地址，如 "A1:D20"、"Sheet2!A1:D20"，或 "used" 表示整个已用区域'),
    sheet: z.string().optional().describe('工作表名。address 中已含表名时可省略'),
    mode: z.enum(['markdown', 'csv', 'detailed']).optional().default('markdown'),
  }),
  summarize: (a) => `读取 ${a.sheet ? `${a.sheet}!` : ''}${a.address}`,
  async run(args) {
    const { address, sheet, mode } = args;

    const data = await runExcel(async (ctx) => {
      const ws = resolveSheet(ctx, sheet, address);
      const range =
        address.toLowerCase() === 'used'
          ? ws.getUsedRange()
          : ws.getRange(stripSheet(address));

      range.load(['address', 'values', 'rowCount', 'columnCount']);
      if (mode === 'detailed') range.load(['formulas', 'numberFormat']);
      await ctx.sync();

      return {
        address: range.address,
        values: range.values as CellValue[][],
        formulas: mode === 'detailed' ? (range.formulas as string[][]) : null,
        numberFormat: mode === 'detailed' ? (range.numberFormat as string[][]) : null,
        cells: range.rowCount * range.columnCount,
      };
    });

    // 超限时首尾采样，并明确告诉模型被截断了、可以分块再读
    const perRow = data.values[0]?.length || 1;
    const maxRows = Math.max(1, Math.floor(MAX_CELLS / perRow));
    const { rows, omitted, tailStart } = sampleRows(data.values, maxRows);

    let body: string;
    if (mode === 'csv') {
      body = toCsv(rows);
    } else if (mode === 'detailed' && data.formulas) {
      const f = sampleRows(data.formulas, maxRows).rows;
      const n = sampleRows(data.numberFormat!, maxRows).rows;
      body = [
        '值：',
        toMarkdown(rows, data.address),
        '',
        '公式：',
        toMarkdown(f as CellValue[][], data.address),
        '',
        '数字格式：',
        toMarkdown(n as CellValue[][], data.address),
      ].join('\n');
    } else {
      body = toMarkdown(rows, data.address);
    }

    const note = omitted
      ? `\n\n> 该区域共 ${data.values.length} 行，已省略中间 ${omitted} 行（展示首段与第 ${tailStart + 1} 行起的尾段）。如需完整数据请按行区间分块读取。`
      : '';

    return { text: `区域 ${data.address}（${data.cells} 个单元格）\n\n${body}${note}` };
  },
});

/** address 里带 "Sheet!" 时优先用它，否则用 sheet 参数，再否则当前表 */
export function resolveSheet(
  ctx: Excel.RequestContext,
  sheet: string | undefined,
  address?: string,
): Excel.Worksheet {
  const inAddress = address?.includes('!') ? address.split('!')[0].replace(/^'|'$/g, '') : undefined;
  const name = inAddress || sheet;
  return name
    ? ctx.workbook.worksheets.getItem(name)
    : ctx.workbook.worksheets.getActiveWorksheet();
}

export function stripSheet(address: string): string {
  return address.includes('!') ? address.split('!')[1] : address;
}
