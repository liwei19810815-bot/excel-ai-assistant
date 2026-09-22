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
      'list_queries',
      'modify_structure',
      'read_range',
      'refresh_query',
      'run_script',
      'write_cells',
    ]);
  });

  it('破坏性工具标记为 mutate:structure，会触发强制确认', () => {
    expect(get('modify_structure')!.policy).toBe('mutate:structure');
    expect(get('run_script')!.policy).toBe('mutate:structure');

    // 刷新 Power Query 会把表里的数据整片换掉，而且【工具箱的撤销框架
    // 覆盖不到它】——那套快照是给 VBA 命令用的，sidecar 的改动它不知道。
    // 事实上不可撤销，所以必须强制确认。
    expect(get('refresh_query')!.policy).toBe('mutate:structure');
  });

  it('sidecar 的只读工具不标成会改表', () => {
    // 标反了比不标更糟：用户以为只是看看，结果被要求确认一个破坏性操作，
    // 或者反过来，一个会改数据的动作悄悄跑掉了。
    expect(get('list_queries')!.policy).toBe('read');
  });

  it('sidecar 类工具能被整体摘掉（sidecar 没跑时就靠这个）', () => {
    const defs = toToolDefs(new Set(['list_queries', 'refresh_query']));
    const names = defs.map((d) => d.name);
    expect(names).not.toContain('list_queries');
    expect(names).not.toContain('refresh_query');
    // 其余工具不受影响——这是"安静降级"的关键
    expect(names).toContain('read_range');
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
