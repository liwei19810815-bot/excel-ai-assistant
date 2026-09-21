import { describe, it, expect, beforeEach, vi } from 'vitest';
import { provision, applyRibbonVisibility } from '../store/provisioning';
import { runStartup } from './startup';
import { useSettings } from '../store/settings';

/**
 * 启动路径的测试。
 *
 * main.tsx 里那段启动逻辑是：
 *     mount()  →  runStartup()
 * 而 runStartup() = provision() → applyRibbonVisibility(结果)
 *
 * 挂载那一步要 DOM，这里不碰；但【链路本身】必须有覆盖——
 * 它决定了"服务端把功能关了，客户端到底会不会真的停用"。
 *
 * 【第一版测试是分别调这两个函数的】，那只能证明两个函数各自能用，
 * 证明不了 main.tsx 真的把它们串起来了：链路被删、顺序写反、
 * 传错字段都不会变红（评审指出）。所以启动逻辑被抽成 runStartup()，
 * 下面直接断言这条链路。
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

//============================================================================
// 启动链路本身
//
// 上面那些是各个零件；这里断言的是【它们真的被串起来了】。
//============================================================================
describe('runStartup 链路', () => {
  it('问完配置之后，真的按结果去更新了功能区按钮', async () => {
    const seen: boolean[] = [];
    (globalThis as { Office?: unknown }).Office = {
      ribbon: {
        requestUpdate: async (arg: unknown) => {
          seen.push(
            (arg as { tabs: { groups: { controls: { enabled: boolean }[] }[] }[] })
              .tabs[0].groups[0].controls[0].enabled,
          );
        },
      },
      context: { requirements: { isSetSupported: () => true } },
    };

    // 没有 location 时 provision 读不到身份 → byok + visibility 1
    const r = await runStartup();

    expect(r.mode).toBe('byok');
    expect(r.visibility).toBe(1);
    // 【关键】：链路走通了，按钮被更新过，且传的是 visibility 对应的值
    expect(r.ribbonUpdated).toBe(true);
    expect(seen).toEqual([true]);
    expect(useSettings.getState().provision.status).toBe('ready');
  });

  it('Office 不可用时链路照样走完，只是按钮没更新', async () => {
    delete (globalThis as { Office?: unknown }).Office;
    const r = await runStartup();
    expect(r.ribbonUpdated).toBe(false);
    // 【不能因为按钮更新不了就不放行】——那会让老版本 Office 上的人用不了
    expect(useSettings.getState().provision.status).toBe('ready');
  });
});
