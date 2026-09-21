import { z } from 'zod';
import { register } from '../registry';
import { runExcel } from '../../excel/coordinator';
import { capture } from '../../excel/checkpoint';
import { resolveSheet } from './read';

/**
 * 结构性改动 —— policy 为 mutate:structure，Agent 循环会强制弹确认。
 * 删行删表这类操作没有廉价的完整快照（整表快照代价太高），
 * 所以这里对「删除行列」仍抓取受影响区域快照，对「删除工作表」只能靠确认拦截。
 */
register({
  name: 'modify_structure',
  description:
    '修改工作簿结构：插入/删除行列、新建/删除/重命名/移动工作表。' +
    '这是破坏性操作，执行前会要求用户确认。' +
    '删除行列会创建快照可撤销；删除工作表无法撤销，务必先向用户确认意图。',
  policy: 'mutate:structure',
  schema: z.object({
    action: z.enum([
      'insert_rows',
      'delete_rows',
      'insert_columns',
      'delete_columns',
      'add_sheet',
      'delete_sheet',
      'rename_sheet',
      'move_sheet',
    ]),
    sheet: z.string().optional().describe('目标工作表名，省略则用当前表'),
    /** 行列操作用 */
    start: z.number().optional().describe('起始行号（1 基）或列号（1 基，A=1）'),
    count: z.number().optional().default(1).describe('操作的行数/列数'),
    /** 工作表操作用 */
    name: z.string().optional().describe('add_sheet 的新表名，或 rename_sheet 的目标名'),
    position: z.number().optional().describe('move_sheet 的目标位置（0 基）'),
  }),
  summarize: (a) => {
    const label: Record<string, string> = {
      insert_rows: '插入行', delete_rows: '删除行',
      insert_columns: '插入列', delete_columns: '删除列',
      add_sheet: '新建工作表', delete_sheet: '删除工作表',
      rename_sheet: '重命名工作表', move_sheet: '移动工作表',
    };
    const where = a.sheet ? `「${a.sheet}」` : '';
    if (a.action.includes('rows')) return `${label[a.action]}${where} 第${a.start}行起 ${a.count} 行`;
    if (a.action.includes('columns')) return `${label[a.action]}${where} 第${a.start}列起 ${a.count} 列`;
    return `${label[a.action]}${where}${a.name ? ` → ${a.name}` : ''}`;
  },

  async run(args) {
    const { action, sheet, start, count = 1, name, position } = args;

    // 行列删除先抓快照
    let checkpointId: string | undefined;
    if (action === 'delete_rows' || action === 'delete_columns') {
      if (start === undefined) throw new Error(`${action} 需要提供 start。`);
      const addr =
        action === 'delete_rows'
          ? `${start}:${start + count - 1}`
          : `${colLetter(start)}:${colLetter(start + count - 1)}`;
      const sheetName = sheet ?? (await activeSheetName());
      const cp = await capture(sheetName, addr, `${action} → ${sheetName}!${addr}`);
      checkpointId = cp.id;
    }

    const message = await runExcel(async (ctx) => {
      const wb = ctx.workbook;

      switch (action) {
        case 'insert_rows':
        case 'delete_rows': {
          if (start === undefined) throw new Error('需要提供 start（起始行号，1 基）。');
          const ws = resolveSheet(ctx, sheet);
          const range = ws.getRange(`${start}:${start + count - 1}`);
          if (action === 'insert_rows') range.insert(Excel.InsertShiftDirection.down);
          else range.delete(Excel.DeleteShiftDirection.up);
          await ctx.sync();
          return `已${action === 'insert_rows' ? '插入' : '删除'}第 ${start} 行起的 ${count} 行`;
        }

        case 'insert_columns':
        case 'delete_columns': {
          if (start === undefined) throw new Error('需要提供 start（起始列号，1 基，A=1）。');
          const ws = resolveSheet(ctx, sheet);
          const range = ws.getRange(`${colLetter(start)}:${colLetter(start + count - 1)}`);
          if (action === 'insert_columns') range.insert(Excel.InsertShiftDirection.right);
          else range.delete(Excel.DeleteShiftDirection.left);
          await ctx.sync();
          return `已${action === 'insert_columns' ? '插入' : '删除'}第 ${colLetter(start)} 列起的 ${count} 列`;
        }

        case 'add_sheet': {
          const ws = wb.worksheets.add(name);
          ws.activate();
          ws.load('name');
          await ctx.sync();
          return `已新建并激活工作表「${ws.name}」`;
        }

        case 'delete_sheet': {
          const ws = resolveSheet(ctx, sheet);
          ws.load('name');
          await ctx.sync();
          const deleted = ws.name;
          ws.delete();
          await ctx.sync();
          return `已删除工作表「${deleted}」（此操作无法撤销）`;
        }

        case 'rename_sheet': {
          if (!name) throw new Error('rename_sheet 需要提供 name。');
          const ws = resolveSheet(ctx, sheet);
          ws.load('name');
          await ctx.sync();
          const old = ws.name;
          ws.name = name;
          await ctx.sync();
          return `已将工作表「${old}」重命名为「${name}」`;
        }

        case 'move_sheet': {
          if (position === undefined) throw new Error('move_sheet 需要提供 position。');
          const ws = resolveSheet(ctx, sheet);
          ws.position = position;
          ws.load('name');
          await ctx.sync();
          return `已将工作表「${ws.name}」移到位置 ${position}`;
        }

        default:
          throw new Error(`不支持的 action：${action}`);
      }
    });

    return { text: message, checkpointId };
  },
});

function colLetter(index1Based: number): string {
  let n = index1Based;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function activeSheetName(): Promise<string> {
  return runExcel(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load('name');
    await ctx.sync();
    return ws.name;
  });
}
