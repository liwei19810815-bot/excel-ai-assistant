import { describe, it, expect } from 'vitest';
import { QUICK_ACTIONS, QUICK_ACTION_GROUPS, actionsByGroup } from './quickActions';

/**
 * 预置提示词的完整性。
 *
 * 这些断言守的是"界面上出现一个点了没反应/分组里空无一物"的低级问题——
 * 它们不会让任何功能测试变红，但用户一眼就看得到。
 */
describe('预置提示词', () => {
  it('id 不能重复（重复会让 React key 冲突、点击错位）', () => {
    const ids = QUICK_ACTIONS.map((a) => a.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('每条都必须有 label 和 prompt', () => {
    for (const a of QUICK_ACTIONS) {
      expect(a.label.trim(), `${a.id} 的 label`).not.toBe('');
      expect(a.prompt.trim(), `${a.id} 的 prompt`).not.toBe('');
    }
  });

  it('每条的分组都必须是已声明的分组之一', () => {
    // 写错一个字，那条就会从界面上【静默消失】——分组是按名字过滤的
    for (const a of QUICK_ACTIONS) {
      expect(QUICK_ACTION_GROUPS, `${a.id} 的分组「${a.group}」`).toContain(a.group);
    }
  });

  it('每个分组都至少有一条，不能出现空分组', () => {
    for (const g of QUICK_ACTION_GROUPS) {
      expect(actionsByGroup(g).length, `分组「${g}」`).toBeGreaterThan(0);
    }
  });

  //--------------------------------------------------------------------------
  // 措辞上的硬要求
  //--------------------------------------------------------------------------
  it('prompt 要足够具体，不能是一句含糊的指令', () => {
    // 【含糊的指令会让模型乱猜范围】，结果往往是改了不该改的地方。
    // 20 个字是个下限，够写清"哪一块、做什么"。
    for (const a of QUICK_ACTIONS) {
      expect(a.prompt.length, `${a.id} 的 prompt 太短`).toBeGreaterThan(20);
    }
  });

  it('只读类的条目不能标成会改表', () => {
    // 标反了比不标更糟：用户以为只是看看，结果数据被改了
    const readOnly = ['audit', 'describe', 'explain-formula', 'fix-errors', 'check-total'];
    for (const id of readOnly) {
      const a = QUICK_ACTIONS.find((x) => x.id === id);
      expect(a, `找不到 ${id}`).toBeDefined();
      expect(a!.mutates, `${id} 应该是只读的`).toBe(false);
    }
  });

  it('只读条目的 prompt 里要明确写出"不要改"', () => {
    // 光把 mutates 标成 false 是给界面看的；模型看的是 prompt。
    // prompt 里不写清楚，模型照样会动手。
    for (const a of QUICK_ACTIONS.filter((x) => !x.mutates)) {
      expect(
        /不要改|不改动|先不要|只分析|只解释|不要直接/.test(a.prompt),
        `${a.id} 标成只读，但 prompt 没告诉模型不要改动`,
      ).toBe(true);
    }
  });
});
