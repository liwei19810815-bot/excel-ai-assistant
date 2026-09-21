import { describe, it, expect, beforeEach, vi } from 'vitest';
import { provision, applyRibbonVisibility } from '../store/provisioning';
import { useSettings } from '../store/settings';

/**
 * 启动路径的测试。
 *
 * main.tsx 里那段启动逻辑是：
 *     mount()  →  provision()  →  applyRibbonVisibility(结果)
 *
 * 挂载那一步要 DOM，这里不碰；但【后两步的接线和异常路径】必须有覆盖——
 * 它们决定了"服务端把功能关了，客户端到底会不会真的停用"。
 * 评审指出这条路径此前完全没有测试。
 */

const ORIGIN = 'https://gw.example.com';

function fakeFetch(body: unknown): typeof fetch {
  return (async () => ({ ok: true, status: 200, json: async () => body })) as unknown as typeof fetch;
}

beforeEach(() => {
  useSettings.setState({
    kind: 'openai-compatible',
    baseUrl: '',
    apiKey: '',
    model: '',
    enableRunScript: true,
    provision: { mode: 'byok', user: '', visibility: 1, status: 'pending' },
  });
  delete (globalThis as { Office?: unknown }).Office;
});

describe('启动路径', () => {
  it('走完之后状态一定是 ready（界面才会放行）', async () => {
    const r = await provision({
      user: 'z',
      origin: ORIGIN,
      fetchImpl: fakeFetch({ visibility: 1, mode: 'byok' }),
    });
    await applyRibbonVisibility(r.visibility);
    expect(useSettings.getState().provision.status).toBe('ready');
  });

  it('服务端关掉功能时，启动走完是"停用"而不是"可用"', async () => {
    const r = await provision({
      user: 'z',
      origin: ORIGIN,
      fetchImpl: fakeFetch({ visibility: 0, mode: 'byok' }),
    });
    await applyRibbonVisibility(r.visibility);

    const s = useSettings.getState();
    expect(s.provision.status).toBe('ready');
    expect(s.provision.visibility).toBe(0);
    // ChatPane 用的就是这个判据
    expect(s.provision.visibility !== 1).toBe(true);
  });

  //--------------------------------------------------------------------------
  // applyRibbonVisibility 是【尽力而为】的：
  // 老版本 Office 上没有 RibbonApi，任务窗格照样得能用。
  // 为了一个视觉效果把启动搞崩是不划算的。
  //--------------------------------------------------------------------------
  it('没有 Office 对象时（浏览器里调 UI）不报错，返回 false', async () => {
    expect(await applyRibbonVisibility(0)).toBe(false);
  });

  it('Office 不支持 RibbonApi 时不报错', async () => {
    (globalThis as { Office?: unknown }).Office = {
      ribbon: { requestUpdate: vi.fn() },
      context: { requirements: { isSetSupported: () => false } },
    };
    expect(await applyRibbonVisibility(0)).toBe(false);
  });

  it('requestUpdate 抛异常时不把启动带崩', async () => {
    (globalThis as { Office?: unknown }).Office = {
      ribbon: {
        requestUpdate: () => {
          throw new Error('boom');
        },
      },
      context: { requirements: { isSetSupported: () => true } },
    };
    expect(await applyRibbonVisibility(0)).toBe(false);
  });

  it('支持时按可见性传 enabled：1 → true，0/2 → false', async () => {
    const calls: unknown[] = [];
    (globalThis as { Office?: unknown }).Office = {
      ribbon: {
        requestUpdate: async (arg: unknown) => {
          calls.push(arg);
        },
      },
      context: { requirements: { isSetSupported: () => true } },
    };

    expect(await applyRibbonVisibility(1)).toBe(true);
    expect(await applyRibbonVisibility(2)).toBe(true);
    expect(await applyRibbonVisibility(0)).toBe(true);

    const enabled = calls.map(
      (c) => (c as { tabs: { groups: { controls: { enabled: boolean }[] }[] }[] })
        .tabs[0].groups[0].controls[0].enabled,
    );
    expect(enabled).toEqual([true, false, false]);
  });
});
