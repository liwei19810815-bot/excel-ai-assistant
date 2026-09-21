import { z } from 'zod';
import { register } from '../registry';
import { runExcel } from '../../excel/coordinator';
import { capture } from '../../excel/checkpoint';
import { colName, parseAddress, type CellValue } from '../../excel/format';
import { resolveSheet, stripSheet } from './read';

const cellSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

register({
  name: 'write_cells',
  description:
    '向指定区域批量写入值或公式。values 是二维数组，外层为行、内层为列，' +
    '尺寸必须与 address 指定的区域完全一致。' +
    '以 "=" 开头的字符串会被当作公式写入（如 "=SUM(B2:B13)"）。' +
    '写入后会自动读回校验并返回实际结果。此操作会创建快照，可撤销。',
  policy: 'mutate:content',
  schema: z.object({
    address: z.string().describe('目标区域，如 "A1:D13" 或 "Sheet2!A1:D13"'),
    sheet: z.string().optional().describe('工作表名，address 中已含表名时可省略'),
    values: z
      .array(z.array(cellSchema))
      .describe('二维数组，[[行1列1, 行1列2], [行2列1, 行2列2]]。尺寸须与 address 一致'),
  }),
  summarize: (a) => `写入 ${a.sheet ? `${a.sheet}!` : ''}${a.address}`,
  async run(args) {
    const { address, sheet, values } = args;
    if (!values.length || !values[0]?.length) {
      throw new Error('values 不能为空。');
    }

    const addr = stripSheet(address);
    const rows = values.length;
    const cols = values[0].length;
    if (values.some((r) => r.length !== cols)) {
      throw new Error('values 的每一行列数必须相同。');
    }

    // 先快照后写入，顺序不能反
    const targetSheet = await resolveSheetName(sheet, address);
    const cp = await capture(targetSheet, addr, `write_cells → ${targetSheet}!${addr}`);

    const readback = await runExcel(async (ctx) => {
      const ws = resolveSheet(ctx, sheet, address);
      const range = ws.getRange(addr);
      range.load(['rowCount', 'columnCount', 'address']);
      await ctx.sync();

      if (range.rowCount !== rows || range.columnCount !== cols) {
        throw new Error(
          `尺寸不匹配：区域 ${range.address} 是 ${range.rowCount}行×${range.columnCount}列，` +
            `但 values 是 ${rows}行×${cols}列。请调整 address 或 values。`,
        );
      }

      // 用 formulas 写入可同时处理常量与公式两种情况
      range.formulas = values.map((r) => r.map((v) => (v === null ? '' : v)));
      await ctx.sync();

      const verify = ws.getRange(addr);
      verify.load(['values', 'address']);
      await ctx.sync();
      return { values: verify.values as CellValue[][], address: verify.address };
    });

    // 只回传摘要而非全量，避免大批写入把上下文吃光
    const preview = readback.values
      .slice(0, 5)
      .map((r, i) => `  第${i + 1}行: ${r.slice(0, 8).join(' | ')}`)
      .join('\n');
    const more = readback.values.length > 5 ? `\n  …共 ${readback.values.length} 行` : '';

    return {
      text: `已写入 ${readback.address}（${rows}行×${cols}列）。读回校验：\n${preview}${more}`,
      checkpointId: cp.id,
    };
  },
});

register({
  name: 'format_cells',
  description:
    '设置区域的格式：字体（加粗/斜体/字号/颜色）、填充色、数字格式、对齐、边框、列宽行高。' +
    '只传需要修改的字段，未传的保持原样。此操作会创建快照，可撤销。',
  policy: 'mutate:content',
  schema: z.object({
    address: z.string().describe('目标区域，如 "A1:D1"'),
    sheet: z.string().optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    fontSize: z.number().optional(),
    fontColor: z.string().optional().describe('字体颜色，十六进制如 "#FFFFFF"'),
    fillColor: z.string().optional().describe('单元格填充色，十六进制如 "#1F4E79"'),
    numberFormat: z
      .string()
      .optional()
      .describe('数字格式代码，如 "#,##0.00"、"¥#,##0.00"、"0.0%"、"yyyy-mm-dd"'),
    horizontalAlignment: z.enum(['Left', 'Center', 'Right']).optional(),
    verticalAlignment: z.enum(['Top', 'Center', 'Bottom']).optional(),
    wrapText: z.boolean().optional(),
    borders: z
      .boolean()
      .optional()
      .describe('true 则为区域内所有单元格加细实线边框'),
    autofitColumns: z.boolean().optional().describe('true 则自动调整列宽以适应内容'),
    columnWidth: z.number().optional().describe('列宽（磅）。与 autofitColumns 互斥'),
    rowHeight: z.number().optional().describe('行高（磅）'),
  }),
  summarize: (a) => `设置格式 ${a.sheet ? `${a.sheet}!` : ''}${a.address}`,
  async run(args) {
    const addr = stripSheet(args.address);
    const targetSheet = await resolveSheetName(args.sheet, args.address);
    const cp = await capture(targetSheet, addr, `format_cells → ${targetSheet}!${addr}`);

    const applied: string[] = [];

    await runExcel(async (ctx) => {
      const ws = resolveSheet(ctx, args.sheet, args.address);
      const range = ws.getRange(addr);
      const fmt = range.format;

      if (args.bold !== undefined) { fmt.font.bold = args.bold; applied.push(args.bold ? '加粗' : '取消加粗'); }
      if (args.italic !== undefined) { fmt.font.italic = args.italic; applied.push('斜体'); }
      if (args.fontSize !== undefined) { fmt.font.size = args.fontSize; applied.push(`字号${args.fontSize}`); }
      if (args.fontColor) { fmt.font.color = args.fontColor; applied.push(`字色${args.fontColor}`); }
      if (args.fillColor) { fmt.fill.color = args.fillColor; applied.push(`填充${args.fillColor}`); }
      if (args.numberFormat) {
        // numberFormat 需要与区域同尺寸，用单元素数组由 Office.js 广播
        range.numberFormat = [[args.numberFormat]];
        applied.push(`数字格式 ${args.numberFormat}`);
      }
      if (args.horizontalAlignment) {
        fmt.horizontalAlignment = args.horizontalAlignment as Excel.HorizontalAlignment;
        applied.push(`水平${args.horizontalAlignment}`);
      }
      if (args.verticalAlignment) {
        fmt.verticalAlignment = args.verticalAlignment as Excel.VerticalAlignment;
        applied.push(`垂直${args.verticalAlignment}`);
      }
      if (args.wrapText !== undefined) { fmt.wrapText = args.wrapText; applied.push('自动换行'); }

      if (args.borders) {
        for (const edge of [
          'EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight',
          'InsideHorizontal', 'InsideVertical',
        ] as const) {
          const b = fmt.borders.getItem(edge);
          b.style = Excel.BorderLineStyle.continuous;
          b.weight = Excel.BorderWeight.thin;
        }
        applied.push('边框');
      }

      if (args.columnWidth !== undefined) { fmt.columnWidth = args.columnWidth; applied.push(`列宽${args.columnWidth}`); }
      if (args.rowHeight !== undefined) { fmt.rowHeight = args.rowHeight; applied.push(`行高${args.rowHeight}`); }

      await ctx.sync();

      // autofit 必须在内容与其它格式落定后单独执行
      if (args.autofitColumns) {
        ws.getRange(addr).format.autofitColumns();
        await ctx.sync();
        applied.push('自适应列宽');
      }
    });

    return {
      text: `已对 ${targetSheet}!${addr} 应用格式：${applied.join('、') || '（无改动）'}`,
      checkpointId: cp.id,
    };
  },
});

/** 快照需要确切的表名，这里先解析一次 */
async function resolveSheetName(sheet: string | undefined, address: string): Promise<string> {
  const inAddress = address.includes('!') ? address.split('!')[0].replace(/^'|'$/g, '') : undefined;
  if (inAddress) return inAddress;
  if (sheet) return sheet;
  return runExcel(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load('name');
    await ctx.sync();
    return ws.name;
  });
}

/** 供 create_chart 等工具复用：把二维数据的尺寸换算成结束地址 */
export function endAddress(start: string, rows: number, cols: number): string {
  const { col, row } = parseAddress(start);
  return `${colName(col)}${row + 1}:${colName(col + cols - 1)}${row + rows}`;
}
