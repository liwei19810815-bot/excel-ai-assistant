import { describe, it, expect } from 'vitest';
import { colName, parseAddress, toMarkdown, toCsv, sampleRows } from './format';

describe('列名换算', () => {
  it('处理单字母与进位', () => {
    expect(colName(0)).toBe('A');
    expect(colName(25)).toBe('Z');
    expect(colName(26)).toBe('AA');
    expect(colName(701)).toBe('ZZ');
    expect(colName(702)).toBe('AAA');
  });
});

describe('地址解析', () => {
  it('解析区域、单格、带表名与绝对引用', () => {
    expect(parseAddress('A1')).toEqual({ col: 0, row: 0 });
    expect(parseAddress('B3:D10')).toEqual({ col: 1, row: 2 });
    expect(parseAddress('Sheet1!C5')).toEqual({ col: 2, row: 4 });
    expect(parseAddress('$AA$100')).toEqual({ col: 26, row: 99 });
  });
});

describe('markdown 序列化', () => {
  it('带上真实行列号，让模型能精确引用单元格', () => {
    const md = toMarkdown(
      [
        ['姓名', '金额'],
        ['张三', 100],
      ],
      'B3:C4',
    );
    // 起始 B3 → 列头应是 B、C，行号应是 3、4
    expect(md).toContain('| B | C |');
    expect(md).toContain('| 3 | 姓名 | 金额 |');
    expect(md).toContain('| 4 | 张三 | 100 |');
  });

  it('转义竖线，避免破坏表格结构', () => {
    expect(toMarkdown([['a|b']], 'A1')).toContain('a\\|b');
  });

  it('空区域给出明确提示而不是空字符串', () => {
    expect(toMarkdown([], 'A1')).toBe('（空区域）');
  });
});

describe('csv 序列化', () => {
  it('对含逗号/引号/换行的值加引号并转义', () => {
    expect(toCsv([['a,b', 'he said "hi"']])).toBe('"a,b","he said ""hi"""');
  });
});

describe('大区域采样', () => {
  it('未超限时原样返回', () => {
    const rows = [[1], [2], [3]];
    const r = sampleRows(rows, 10);
    expect(r.rows).toHaveLength(3);
    expect(r.omitted).toBe(0);
  });

  it('超限时保留首尾并报告省略行数', () => {
    const rows = Array.from({ length: 100 }, (_, i) => [i]);
    const r = sampleRows(rows, 10);
    expect(r.rows).toHaveLength(10);
    expect(r.omitted).toBe(90);
    // 首段保留开头，尾段保留结尾
    expect(r.rows[0]).toEqual([0]);
    expect(r.rows[r.rows.length - 1]).toEqual([99]);
  });
});
