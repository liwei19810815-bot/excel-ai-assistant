import { describe, it, expect } from 'vitest';
import './index';
import { all, get, toToolDefs, truncate, MAX_RESULT_CHARS } from './registry';
import { EXCEL_TOOL_NAMES } from './excel';
import { POWERPOINT_TOOL_NAMES } from './powerpoint';
import { SIDECAR_TOOL_NAMES } from './sidecar/powerQuery';
import { RUN_MACRO_TOOL_NAMES } from './sidecar/runMacro';

describe('工具注册表', () => {
  it('注册了阶段一约定的核心工具', () => {
    const names = all().map((t) => t.name).sort();
    expect(names).toEqual([
      'add_slide',
      'add_text_box',
      'create_chart',
      'format_cells',
      'get_presentation_overview',
      'get_selection',
      'get_workbook_overview',
      'list_macros',
      'list_queries',
      'modify_structure',
      'read_range',
      'read_slide',
      'refresh_query',
      'run_macro',
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

    // 调宏比刷新查询风险更高——宏可以做任何事，执行前完全不知道
    // 它会改什么，唯一的安全边界是用户点头，绝不能标成只读。
    expect(get('run_macro')!.policy).toBe('mutate:structure');
  });

  it('sidecar 的只读工具不标成会改表', () => {
    // 标反了比不标更糟：用户以为只是看看，结果被要求确认一个破坏性操作，
    // 或者反过来，一个会改数据的动作悄悄跑掉了。
    expect(get('list_queries')!.policy).toBe('read');
    expect(get('list_macros')!.policy).toBe('read');
  });

  it('sidecar 类工具能被整体摘掉（sidecar 没跑时就靠这个）', () => {
    const defs = toToolDefs(new Set(['list_queries', 'refresh_query', 'list_macros', 'run_macro']));
    const names = defs.map((d) => d.name);
    expect(names).not.toContain('list_queries');
    expect(names).not.toContain('refresh_query');
    expect(names).not.toContain('list_macros');
    expect(names).not.toContain('run_macro');
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

/**
 * 按宿主过滤用的名字清单（EXCEL_TOOL_NAMES / POWERPOINT_TOOL_NAMES）。
 *
 * 【这几条断言防的是什么】：ChatPane 按这两份名单往 disabled 集合里塞
 * 名字，名单里写错一个字或者漏加一个新工具，表现是"某个工具在错误的
 * 宿主上也能被模型看到、点了才报错"，或者"某个工具在对的宿主上也被
 * 误摘掉"——两种都不会让任何编译或运行时检查变红，只有对着注册表
 * 逐条核对名字才能抓到。
 */
describe('按宿主过滤的工具名单', () => {
  it('EXCEL_TOOL_NAMES 里的每个名字都确实注册过', () => {
    for (const n of EXCEL_TOOL_NAMES) {
      expect(get(n), `EXCEL_TOOL_NAMES 里的 "${n}" 没有对应的已注册工具`).toBeDefined();
    }
  });

  it('POWERPOINT_TOOL_NAMES 里的每个名字都确实注册过', () => {
    for (const n of POWERPOINT_TOOL_NAMES) {
      expect(get(n), `POWERPOINT_TOOL_NAMES 里的 "${n}" 没有对应的已注册工具`).toBeDefined();
    }
  });

  it('Excel 系（含 sidecar）和 PowerPoint 系的名单不重叠', () => {
    const excelSide = new Set<string>([...EXCEL_TOOL_NAMES, ...SIDECAR_TOOL_NAMES, ...RUN_MACRO_TOOL_NAMES]);
    for (const n of POWERPOINT_TOOL_NAMES) {
      expect(excelSide.has(n), `"${n}" 同时出现在 Excel 系和 PowerPoint 系名单里`).toBe(false);
    }
  });

  it('所有已注册工具都被两份名单之一覆盖到——不能有工具"两边都不摘"', () => {
    // 覆盖不到的后果是：这个工具在 PPT 上也会出现在模型可见的工具列表里，
    // 点了才报错，而不是从一开始就静默消失。
    const covered = new Set<string>([
      ...EXCEL_TOOL_NAMES,
      ...SIDECAR_TOOL_NAMES,
      ...RUN_MACRO_TOOL_NAMES,
      ...POWERPOINT_TOOL_NAMES,
      'run_script', // Excel 专属但由 settings.enableRunScript 单独控制，不在按宿主的名单里
    ]);
    for (const t of all()) {
      expect(covered.has(t.name), `"${t.name}" 没有被任何按宿主过滤的名单覆盖到`).toBe(true);
    }
  });

  it('PowerPoint 的写工具都标成 mutate:structure——没有自定义快照系统撑腰，' +
    '不能承诺 mutate:content 那种"自动建快照、可撤销"，唯一的安全边界是强制确认', () => {
    expect(get('add_text_box')!.policy).toBe('mutate:structure');
    expect(get('add_slide')!.policy).toBe('mutate:structure');
  });

  it('PowerPoint 的读工具不标成会改动', () => {
    expect(get('get_presentation_overview')!.policy).toBe('read');
    expect(get('read_slide')!.policy).toBe('read');
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
