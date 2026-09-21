import { describe, it, expect } from 'vitest';
import './index';
import { all, get, toToolDefs, truncate, MAX_RESULT_CHARS } from './registry';

describe('工具注册表', () => {
  it('注册了阶段一约定的核心工具', () => {
    const names = all().map((t) => t.name).sort();
    expect(names).toEqual([
      'create_chart',
      'format_cells',
      'get_selection',
      'get_workbook_overview',
      'modify_structure',
      'read_range',
      'run_script',
      'write_cells',
    ]);
  });

  it('破坏性工具标记为 mutate:structure，会触发强制确认', () => {
    expect(get('modify_structure')!.policy).toBe('mutate:structure');
    expect(get('run_script')!.policy).toBe('mutate:structure');
  });

  it('只读工具不建快照', () => {
    for (const name of ['read_range', 'get_selection', 'get_workbook_overview']) {
      expect(get(name)!.policy).toBe('read');
    }
  });

  it('导出的 JSON Schema 不含 $ref —— 多数厂商不解析引用', () => {
    const json = JSON.stringify(toToolDefs());
    expect(json).not.toContain('$ref');
  });

  it('每个工具都能导出带 properties 的 object schema', () => {
    for (const def of toToolDefs()) {
      expect(def.schema).toHaveProperty('type', 'object');
      expect(def.description.length).toBeGreaterThan(10);
    }
  });

  it('disabled 集合里的工具会被摘除', () => {
    const defs = toToolDefs(new Set(['run_script']));
    expect(defs.find((d) => d.name === 'run_script')).toBeUndefined();
    expect(defs.length).toBe(all().length - 1);
  });
});

describe('结果截断', () => {
  it('短文本原样返回', () => {
    expect(truncate('hello')).toBe('hello');
  });

  it('超长文本截断并提示可分块读取', () => {
    const out = truncate('x'.repeat(MAX_RESULT_CHARS + 500));
    expect(out.length).toBeLessThan(MAX_RESULT_CHARS + 200);
    expect(out).toContain('已截断');
  });
});
