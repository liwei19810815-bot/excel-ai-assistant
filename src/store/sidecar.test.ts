import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  readInjectedSidecarToken,
  probeSidecar,
  callSidecar,
  getSidecarStatus,
  setSidecarStatus,
  SIDECAR_TOKEN_HEADER,
  SIDECAR_BASE_PORT,
  type SidecarStatus,
} from './sidecar';

const TOKEN = 'a'.repeat(64);

function okHealth(port: number) {
  return {
    ok: true,
    json: async () => ({ ok: true, name: 'excel-toolbox-sidecar', version: '0.1.0', port }),
  } as unknown as Response;
}

describe('读取注入的 sidecar 令牌', () => {
  it('从 ?sidecar= 里读出来', () => {
    expect(readInjectedSidecarToken(`?u=alice&sidecar=${TOKEN}`)).toBe(TOKEN);
  });

  it('没装这个组件时安装器写的是空串，要能正常读成空', () => {
    expect(readInjectedSidecarToken('?u=alice&sidecar=')).toBe('');
  });

  it('参数根本不存在时返回空串，不抛', () => {
    expect(readInjectedSidecarToken('?u=alice')).toBe('');
  });
});

describe('探测 sidecar', () => {
  it('没有令牌时【一个请求都不发】', async () => {
    // 绝大多数用户没装这个组件。白发 5 个注定失败的请求既拖慢启动，
    // 又会在控制台刷一片连接错误，看着像出了故障。
    const fetchImpl = vi.fn();
    const r = await probeSidecar('', { fetchImpl: fetchImpl as unknown as typeof fetch });
    expect(r.available).toBe(false);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('令牌放在请求头里，不放查询串', async () => {
    // 这是整套防护的支点：自定义头强制浏览器预检，远端网页发不出来。
    // 换成 ?token= 就成了简单请求，不触发预检，门就自己拆了。
    const fetchImpl = vi.fn(async () => okHealth(SIDECAR_BASE_PORT));
    await probeSidecar(TOKEN, { fetchImpl: fetchImpl as unknown as typeof fetch });

    const [url, init] = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).not.toContain('token');
    expect(url).not.toContain(TOKEN);
    expect((init.headers as Record<string, string>)[SIDECAR_TOKEN_HEADER]).toBe(TOKEN);
  });

  it('只连回环地址', async () => {
    const fetchImpl = vi.fn(async () => okHealth(SIDECAR_BASE_PORT));
    await probeSidecar(TOKEN, { fetchImpl: fetchImpl as unknown as typeof fetch });
    const [url] = fetchImpl.mock.calls[0] as unknown as [string];
    expect(url.startsWith('http://127.0.0.1:')).toBe(true);
  });

  it('第一个端口就命中时不再往后探', async () => {
    const fetchImpl = vi.fn(async () => okHealth(SIDECAR_BASE_PORT));
    const r = await probeSidecar(TOKEN, { fetchImpl: fetchImpl as unknown as typeof fetch });
    expect(r.available).toBe(true);
    expect(r.port).toBe(SIDECAR_BASE_PORT);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('端口顺延后仍能探到', async () => {
    const hit = SIDECAR_BASE_PORT + 2;
    const fetchImpl = vi.fn(async (url: string) => {
      if (url.includes(`:${hit}/`)) return okHealth(hit);
      throw new Error('connection refused');
    });
    const r = await probeSidecar(TOKEN, { fetchImpl: fetchImpl as unknown as typeof fetch });
    expect(r.available).toBe(true);
    expect(r.port).toBe(hit);
  });

  it('全都连不上就安静地报不可用，不抛异常', async () => {
    const fetchImpl = vi.fn(async () => {
      throw new Error('connection refused');
    });
    const r = await probeSidecar(TOKEN, { fetchImpl: fetchImpl as unknown as typeof fetch });
    expect(r.available).toBe(false);
    expect(r.port).toBe(0);
  });

  //--------------------------------------------------------------------------
  // 认错了人比连不上更糟
  //--------------------------------------------------------------------------
  it('端口上是【别的程序】时不能认成 sidecar', async () => {
    // 这个端口上完全可能跑着别的东西。只要它对任意请求回 200，
    // 不校验身份就会把它当成 sidecar，之后每次调用都失败，
    // 而现象看起来像"sidecar 有问题"，排查方向完全跑偏。
    const fetchImpl = vi.fn(async () => ({
      ok: true,
      json: async () => ({ ok: true, name: 'some-other-service', version: '9' }),
    })) as unknown as typeof fetch;
    const r = await probeSidecar(TOKEN, { fetchImpl });
    expect(r.available).toBe(false);
  });

  it('响应里 ok 不为 true 时不算命中', async () => {
    const fetchImpl = vi.fn(async () => ({
      ok: true,
      json: async () => ({ ok: false, name: 'excel-toolbox-sidecar' }),
    })) as unknown as typeof fetch;
    const r = await probeSidecar(TOKEN, { fetchImpl });
    expect(r.available).toBe(false);
  });

  it('401 不算命中（令牌不对就是不可用）', async () => {
    const fetchImpl = vi.fn(async () => ({
      ok: false,
      status: 401,
      json: async () => ({ error: 'unauthorized' }),
    })) as unknown as typeof fetch;
    const r = await probeSidecar(TOKEN, { fetchImpl });
    expect(r.available).toBe(false);
  });

  it('JSON 坏掉时不抛，只当作没探到', async () => {
    const fetchImpl = vi.fn(async () => ({
      ok: true,
      json: async () => {
        throw new Error('bad json');
      },
    })) as unknown as typeof fetch;
    const r = await probeSidecar(TOKEN, { fetchImpl });
    expect(r.available).toBe(false);
  });
});

describe('缓存的状态', () => {
  beforeEach(() => {
    setSidecarStatus({ available: false, port: 0, version: '', reason: '重置' });
  });

  it('默认是不可用（探测完成之前不能让工具出现在列表里）', () => {
    expect(getSidecarStatus().available).toBe(false);
  });

  it('存进去之后读得出来', () => {
    const s: SidecarStatus = { available: true, port: 8899, version: '0.1.0', reason: '' };
    setSidecarStatus(s);
    expect(getSidecarStatus()).toEqual(s);
  });
});

describe('调用具体动作', () => {
  const up: SidecarStatus = { available: true, port: 8899, version: '0.1.0', reason: '' };

  it('不可用时直接抛，不发请求', async () => {
    const fetchImpl = vi.fn();
    await expect(
      callSidecar(
        { available: false, port: 0, version: '', reason: '' },
        '/queries',
        TOKEN,
        {},
        fetchImpl as unknown as typeof fetch,
      ),
    ).rejects.toThrow();
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('调用时也把令牌放在头里', async () => {
    const fetchImpl = vi.fn(async () => ({
      ok: true,
      json: async () => ({ ok: true }),
    })) as unknown as typeof fetch;
    await callSidecar(up, '/queries', TOKEN, {}, fetchImpl);
    const [url, init] = (fetchImpl as unknown as ReturnType<typeof vi.fn>).mock.calls[0] as [
      string,
      RequestInit,
    ];
    expect(url).toBe('http://127.0.0.1:8899/queries');
    expect((init.headers as Record<string, string>)[SIDECAR_TOKEN_HEADER]).toBe(TOKEN);
  });

  it('非 2xx 要抛出来，不能把失败当成功', async () => {
    const fetchImpl = vi.fn(async () => ({
      ok: false,
      status: 403,
      json: async () => ({}),
    })) as unknown as typeof fetch;
    await expect(callSidecar(up, '/queries', TOKEN, {}, fetchImpl)).rejects.toThrow('403');
  });
});
