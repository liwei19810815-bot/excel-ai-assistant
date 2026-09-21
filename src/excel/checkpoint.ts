import { runExcel } from './coordinator';
import type { CellValue } from './format';

/**
 * 写操作前的区域快照，用于撤销。
 *
 * 为什么必须自建：Office.js 的写入**不进入 Excel 原生撤销栈**，用户按 Ctrl+Z
 * 无法回滚插件的改动。这是 Office.js 的既有行为，所以插件必须自己记录快照，
 * 并在 UI 上明确告知用户「用面板里的撤销，不要用 Ctrl+Z」。
 */

export interface Checkpoint {
  id: string;
  /** 给用户看的描述，如 "write_cells → Sheet1!A1:D13" */
  label: string;
  createdAt: number;
  sheetName: string;
  address: string;
  values: CellValue[][];
  formulas: string[][];
  numberFormat: string[][];
}

const MAX_CHECKPOINTS = 20;
const stack: Checkpoint[] = [];

/** 抓取指定区域的当前状态。address 形如 "A1:D13"（不含表名） */
export async function capture(
  sheetName: string,
  address: string,
  label: string,
): Promise<Checkpoint> {
  const snap = await runExcel(async (ctx) => {
    const sheet = sheetName
      ? ctx.workbook.worksheets.getItem(sheetName)
      : ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    const range = sheet.getRange(address);
    range.load(['values', 'formulas', 'numberFormat', 'address']);
    await ctx.sync();
    return {
      name: sheet.name,
      values: range.values as CellValue[][],
      formulas: range.formulas as string[][],
      numberFormat: range.numberFormat as string[][],
    };
  });

  const cp: Checkpoint = {
    id: `cp_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    label,
    createdAt: Date.now(),
    sheetName: snap.name,
    address,
    values: snap.values,
    formulas: snap.formulas,
    numberFormat: snap.numberFormat,
  };

  stack.push(cp);
  if (stack.length > MAX_CHECKPOINTS) stack.shift();
  return cp;
}

/** 回滚到指定快照。恢复公式而非值，避免把公式拍成常量。 */
export async function restore(id: string): Promise<string> {
  const idx = stack.findIndex((c) => c.id === id);
  if (idx === -1) throw new Error('该快照已过期或不存在，无法撤销。');
  const cp = stack[idx];

  await runExcel(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(cp.sheetName);
    const range = sheet.getRange(cp.address);
    range.numberFormat = cp.numberFormat;
    // formulas 对普通值单元格就是值本身，整体写回即可同时覆盖两种情况
    range.formulas = cp.formulas;
    await ctx.sync();
  });

  stack.splice(idx, 1);
  return `已撤销：${cp.label}（${cp.sheetName}!${cp.address}）`;
}

export function latest(): Checkpoint | undefined {
  return stack[stack.length - 1];
}

export function list(): Checkpoint[] {
  return [...stack].reverse();
}
