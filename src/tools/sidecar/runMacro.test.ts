import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import '../index';
import { get } from '../registry';
import { setSidecarStatus, type SidecarStatus } from '../../store/sidecar';

const up: SidecarStatus = { available: true, port: 8899, version: '0.1.0', reason: '' };
const down: SidecarStatus = { available: false, port: 0, version: '', reason: '' };

// probeSidecar 要求 URL 里有令牌才发请求（见 sidecar.ts），
// 所以每个用例都要注入一个假令牌，否则 fetch 永远不会被调用。
//
// 【测试跑在 Node 环境，没有 window】。readInjectedSidecarToken() 读的是
// 裸标识符 location（sidecar.ts 里 `typeof location === 'undefined' ? ...`
// 就是为了兼容这个环境），所以要用 vi.stubGlobal 定义一个裸的 location，
// 不能挂在 window 上——这里根本没有 window。
function withToken() {
  vi.stubGlobal('location', { search: '?u=test&sidecar=' + 'a'.repeat(64), origin: 'https://gw.example.com' });
}

function mockFetch(handler: (url: string, init?: RequestInit) => Promise<{ ok: boolean; status?: number; json: () => Promise<unknown> }>) {
  const fn = vi.fn(handler);
  vi.stubGlobal('fetch', fn as unknown as typeof fetch);
  return fn;
}

function healthOk() {
  return { ok: true, json: async () => ({ ok: true, name: 'excel-toolbox-sidecar', version: '0.1.0' }) };
}

beforeEach(() => {
  setSidecarStatus(down);
  withToken();
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('list_macros', () => {
  it('sidecar 不在时安静降级，不抛异常', async () => {
    const tool = get('list_macros')!;
    // 不给 /health 一个成功响应，probeSidecar 探测不到就该走"不可用"文案——
    // 不能真的调用 /macros。
    const fetchImpl = mockFetch(async () => ({ ok: false, status: 500, json: async () => ({}) }));
    const r = await tool.run({}, {} as never);
    expect(r.text).toContain('没有运行');
    // 探测会试完端口范围（5 个）才放弃，但每一次都只打 /health，
    // 一次 /macros 都不该调——探测失败就不该再往下走。
    for (const [url] of fetchImpl.mock.calls) {
      expect(url as string).toContain('/health');
    }
  });

  it('trusted=false 时如实说"枚不出来"，不说"没有宏"', async () => {
    setSidecarStatus(up);
    mockFetch(async (url) => {
      if (url.includes('/health')) return healthOk();
      return { ok: true, json: async () => ({ ok: true, macros: [], trusted: false }) };
    });

    const tool = get('list_macros')!;
    const r = await tool.run({}, {} as never);
    // 【这条是关键】：不能把"枚不出来"说成"没有宏"，那会让模型
    // 断定工作簿里没有可用的宏，转而拒绝用户的请求。
    expect(r.text).not.toContain('没有暴露给 AI 的宏');
    expect(r.text).toMatch(/枚举不出|没开启/);
  });

  it('trusted=true 且列表为空时，说的是"确实没有"', async () => {
    setSidecarStatus(up);
    mockFetch(async (url) => {
      if (url.includes('/health')) return healthOk();
      return { ok: true, json: async () => ({ ok: true, macros: [], trusted: true }) };
    });

    const tool = get('list_macros')!;
    const r = await tool.run({}, {} as never);
    expect(r.text).toContain('没有暴露给 AI 的宏');
  });

  it('列出结果里的宏名原样透传', async () => {
    setSidecarStatus(up);
    mockFetch(async (url) => {
      if (url.includes('/health')) return healthOk();
      return { ok: true, json: async () => ({ ok: true, macros: ['AI_导出报表', 'AI_清理数据'], trusted: true }) };
    });

    const tool = get('list_macros')!;
    const r = await tool.run({}, {} as never);
    expect(r.text).toContain('AI_导出报表');
    expect(r.text).toContain('AI_清理数据');
  });
});

describe('run_macro：前缀校验', () => {
  it('没有 AI_ 前缀的宏名，前端直接拒绝，连请求都不发', async () => {
    setSidecarStatus(up);
    const fetchImpl = mockFetch(async () => healthOk());
    const tool = get('run_macro')!;
    const r = await tool.run({ name: 'DeleteAllSheets' }, {} as never);
    expect(r.text).toContain('AI_');
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('前缀藏在中间的也拒绝', async () => {
    setSidecarStatus(up);
    const fetchImpl = mockFetch(async () => healthOk());
    const tool = get('run_macro')!;
    const r = await tool.run({ name: 'NotAI_Foo' }, {} as never);
    expect(r.text).toContain('AI_');
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('大小写不对的前缀拒绝', async () => {
    setSidecarStatus(up);
    const fetchImpl = mockFetch(async () => healthOk());
    const tool = get('run_macro')!;
    const r = await tool.run({ name: 'ai_lowercase' }, {} as never);
    expect(r.text).toContain('AI_');
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('合法前缀的宏名会真的发请求', async () => {
    setSidecarStatus(up);
    const fetchImpl = mockFetch(async (url) => {
      if (url.includes('/health')) return healthOk();
      return { ok: true, json: async () => ({ ok: true, name: 'AI_Test', result: '' }) };
    });
    const tool = get('run_macro')!;
    await tool.run({ name: 'AI_Test' }, {} as never);
    expect(fetchImpl).toHaveBeenCalled();
  });
});

describe('run_macro：请求内容', () => {
  it('参数按原样放进请求体，令牌放请求头', async () => {
    setSidecarStatus(up);
    let capturedBody: string | undefined;
    let capturedHeaders: Record<string, string> | undefined;
    mockFetch(async (url, init) => {
      if (url.includes('/health')) return healthOk();
      capturedBody = init?.body as string;
      capturedHeaders = init?.headers as Record<string, string>;
      return { ok: true, json: async () => ({ ok: true, name: 'AI_Test', result: 'done' }) };
    });

    const tool = get('run_macro')!;
    await tool.run({ name: 'AI_Test', args: ['x', 1, true] }, {} as never);

    expect(JSON.parse(capturedBody!)).toEqual({ name: 'AI_Test', args: ['x', 1, true] });
    expect(capturedHeaders?.['X-Toolbox-Token']).toBeDefined();
  });

  it('没传 args 时发送空数组，不是 undefined', async () => {
    setSidecarStatus(up);
    let capturedBody: string | undefined;
    mockFetch(async (url, init) => {
      if (url.includes('/health')) return healthOk();
      capturedBody = init?.body as string;
      return { ok: true, json: async () => ({ ok: true, name: 'AI_Test' }) };
    });

    const tool = get('run_macro')!;
    await tool.run({ name: 'AI_Test' }, {} as never);
    expect(JSON.parse(capturedBody!).args).toEqual([]);
  });
});

describe('run_macro：失败必须如实说', () => {
  const cases: Array<[string, RegExp]> = [
    ['excel_not_running', /没有在运行/],
    ['no_workbook', /没有打开的工作簿/],
    ['timeout', /超时/],
  ];

  for (const [error, expected] of cases) {
    it(`error=${error} 时不能说成功`, async () => {
      setSidecarStatus(up);
      mockFetch(async (url) => {
        if (url.includes('/health')) return healthOk();
        return { ok: true, json: async () => ({ ok: false, error }) };
      });

      const tool = get('run_macro')!;
      const r = await tool.run({ name: 'AI_Test' }, {} as never);
      expect(r.text).toMatch(expected);
      expect(r.text).not.toContain('已调用');
    });
  }

  it('macro_failed 时带出具体原因，不是笼统的"失败"', async () => {
    setSidecarStatus(up);
    mockFetch(async (url) => {
      if (url.includes('/health')) return healthOk();
      return { ok: true, json: async () => ({ ok: false, error: 'macro_failed', message: '下标越界' }) };
    });

    const tool = get('run_macro')!;
    const r = await tool.run({ name: 'AI_Test' }, {} as never);
    expect(r.text).toContain('下标越界');
    expect(r.text).not.toContain('已调用');
  });

  it('成功时才说"已调用"', async () => {
    setSidecarStatus(up);
    mockFetch(async (url) => {
      if (url.includes('/health')) return healthOk();
      return { ok: true, json: async () => ({ ok: true, name: 'AI_Test', result: '完成' }) };
    });

    const tool = get('run_macro')!;
    const r = await tool.run({ name: 'AI_Test' }, {} as never);
    expect(r.text).toContain('已调用');
    expect(r.text).toContain('完成');
  });
});
