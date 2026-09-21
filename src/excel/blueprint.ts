import { runExcel, isSupported } from './coordinator';
import { colName, parseAddress, type CellValue } from './format';

/**
 * 工作簿「蓝图」——只含结构与表头，不含全量数据。
 * 每轮对话自动注入给模型，让它无需先调工具就知道工作簿长什么样。
 * 对 5 万行 / 20 sheet 的大工作簿，这份摘要仍然只有几百 token。
 */

export interface SheetInfo {
  name: string;
  isActive: boolean;
  usedRange: string | null;
  rowCount: number;
  colCount: number;
  /** 推断出的表头（首行非空值），最多 30 列 */
  headers: string[];
}

export interface Blueprint {
  sheets: SheetInfo[];
  tables: Array<{ name: string; sheet: string; range: string }>;
  charts: Array<{ name: string; sheet: string }>;
  selection: { sheet: string; address: string; preview: CellValue[][] } | null;
}

export async function buildBlueprint(): Promise<Blueprint> {
  return runExcel(async (ctx) => {
    const wb = ctx.workbook;
    const sheets = wb.worksheets;
    sheets.load('items/name');
    const active = sheets.getActiveWorksheet();
    active.load('name');

    const sel = wb.getSelectedRange();
    sel.load(['address', 'rowCount', 'columnCount']);

    const tables = wb.tables;
    tables.load('items/name');

    await ctx.sync();

    // 逐表取 usedRange。getUsedRangeOrNullObject 避免空表抛错。
    const used = sheets.items.map((s) => {
      const r = s.getUsedRangeOrNullObject();
      r.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
      return { sheet: s, range: r };
    });

    const tableRanges = tables.items.map((t) => {
      const r = t.getRange();
      r.load('address');
      const ws = t.worksheet;
      ws.load('name');
      return { table: t, range: r, ws };
    });

    // 图表按表收集，仅在 1.1+ 可用（基线即支持）
    const chartsBySheet = sheets.items.map((s) => {
      const c = s.charts;
      c.load('items/name');
      return { sheetName: s.name, charts: c };
    });

    await ctx.sync();

    // 表头：各表 usedRange 首行，限 30 列避免宽表撑爆
    const headerRanges = used.map(({ sheet, range }) => {
      if (range.isNullObject) return null;
      const { col, row } = parseAddress(range.address);
      const width = Math.min(range.columnCount, 30);
      const addr = `${colName(col)}${row + 1}:${colName(col + width - 1)}${row + 1}`;
      const hr = sheet.getRange(addr);
      hr.load('values');
      return hr;
    });

    // 选区预览：最多 10 行 × 10 列，够模型判断数据形态
    const selPreview = wb.getSelectedRange().getAbsoluteResizedRange(
      Math.min(sel.rowCount, 10),
      Math.min(sel.columnCount, 10),
    );
    selPreview.load('values');

    await ctx.sync();

    const sheetInfos: SheetInfo[] = used.map(({ sheet, range }, i) => {
      const hr = headerRanges[i];
      const headers = hr
        ? ((hr.values[0] ?? []) as CellValue[]).map((v) => (v == null ? '' : String(v)))
        : [];
      return {
        name: sheet.name,
        isActive: sheet.name === active.name,
        usedRange: range.isNullObject ? null : range.address,
        rowCount: range.isNullObject ? 0 : range.rowCount,
        colCount: range.isNullObject ? 0 : range.columnCount,
        headers,
      };
    });

    return {
      sheets: sheetInfos,
      tables: tableRanges.map(({ table, range, ws }) => ({
        name: table.name,
        sheet: ws.name,
        range: range.address,
      })),
      charts: chartsBySheet.flatMap(({ sheetName, charts }) =>
        charts.items.map((c) => ({ name: c.name, sheet: sheetName })),
      ),
      selection: {
        sheet: active.name,
        address: sel.address,
        preview: selPreview.values as CellValue[][],
      },
    };
  });
}

/** 蓝图 → 紧凑文本，直接拼进每轮的上下文块 */
export function renderBlueprint(bp: Blueprint): string {
  const lines: string[] = ['## 当前工作簿结构'];

  for (const s of bp.sheets) {
    const mark = s.isActive ? '（当前）' : '';
    if (!s.usedRange) {
      lines.push(`- 工作表「${s.name}」${mark}：空表`);
      continue;
    }
    lines.push(
      `- 工作表「${s.name}」${mark}：数据区 ${s.usedRange}，${s.rowCount} 行 × ${s.colCount} 列`,
    );
    if (s.headers.some((h) => h)) {
      lines.push(`  表头：${s.headers.filter(Boolean).join(' | ')}`);
    }
  }

  if (bp.tables.length) {
    lines.push(`- Excel 表：${bp.tables.map((t) => `${t.name}(${t.sheet}!${t.range})`).join('、')}`);
  }
  if (bp.charts.length) {
    lines.push(`- 图表：${bp.charts.map((c) => `${c.name}(${c.sheet})`).join('、')}`);
  }

  if (bp.selection) {
    lines.push('', `## 用户当前选区：${bp.selection.address}`);
    const preview = bp.selection.preview;
    if (preview.length && preview.some((r) => r.some((v) => v !== '' && v != null))) {
      lines.push('选区内容预览（最多 10×10）：');
      lines.push('```');
      lines.push(preview.map((r) => r.map((v) => (v == null ? '' : v)).join('\t')).join('\n'));
      lines.push('```');
    } else {
      lines.push('（选区为空）');
    }
  }

  // 让模型知道自己所处平台的能力边界，避免反复尝试不支持的 API
  const caps: string[] = [];
  if (!isSupported('1.8')) caps.push('透视表创建（需 1.8）');
  if (!isSupported('1.9')) caps.push('部分图表增强（需 1.9）');
  if (caps.length) {
    lines.push('', `> 当前 Excel 版本不支持：${caps.join('、')}。遇到相关需求请说明并提供替代方案。`);
  }

  return lines.join('\n');
}
