/**
 * 区域数据 → 文本序列化。
 * 给模型看的表示要紧凑（省 token）又要带地址（便于模型精确引用单元格）。
 */

export type CellValue = string | number | boolean | null;

/** 0 → A, 25 → Z, 26 → AA */
export function colName(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 解析 "Sheet1!B3:D10" / "B3" → 起始行列（0 基） */
export function parseAddress(address: string): { col: number; row: number } {
  const ref = address.includes('!') ? address.split('!')[1] : address;
  const m = /^\$?([A-Z]+)\$?(\d+)/i.exec(ref.split(':')[0]);
  if (!m) return { col: 0, row: 0 };
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: parseInt(m[2], 10) - 1 };
}

/**
 * 转成带行列号的 markdown 表格。
 * 行列号让模型能直接说出"把 C5 改成…"，而不用自己数格子。
 */
export function toMarkdown(values: CellValue[][], startAddress: string): string {
  if (!values.length) return '（空区域）';
  const { col: c0, row: r0 } = parseAddress(startAddress);
  const width = values[0].length;

  const header = ['', ...Array.from({ length: width }, (_, i) => colName(c0 + i))];
  const sep = header.map(() => '---');
  const body = values.map((row, i) => [
    String(r0 + i + 1),
    ...row.map(cellText),
  ]);

  return [header, sep, ...body].map((r) => `| ${r.join(' | ')} |`).join('\n');
}

export function toCsv(values: CellValue[][]): string {
  return values
    .map((row) =>
      row
        .map((v) => {
          const s = cellText(v);
          return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        })
        .join(','),
    )
    .join('\n');
}

function cellText(v: CellValue): string {
  if (v === null || v === undefined || v === '') return '';
  return String(v).replace(/\|/g, '\\|').replace(/\n/g, ' ');
}

/**
 * 大区域降级：保留首尾若干行 + 中间省略标记。
 * 避免一次 read_range 就把上下文撑爆，同时让模型知道数据被截断了、可以分块再读。
 */
export function sampleRows<T>(
  rows: T[],
  maxRows: number,
): { rows: T[]; omitted: number; headIndex: number[]; tailStart: number } {
  if (rows.length <= maxRows) {
    return { rows, omitted: 0, headIndex: rows.map((_, i) => i), tailStart: rows.length };
  }
  const head = Math.ceil(maxRows * 0.7);
  const tail = maxRows - head;
  return {
    rows: [...rows.slice(0, head), ...rows.slice(rows.length - tail)],
    omitted: rows.length - maxRows,
    headIndex: rows.slice(0, head).map((_, i) => i),
    tailStart: rows.length - tail,
  };
}
