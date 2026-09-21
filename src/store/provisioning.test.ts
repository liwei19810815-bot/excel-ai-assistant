import { describe, it, expect, beforeEach, vi } from 'vitest';
import { readInjectedUser, parseConfig, parseVisibility, fetchAiConfig, provision } from './provisioning';
import { useSettings } from './settings';

/** 造一个只会返回指定内容的 fetch */
function fakeFetch(body: unknown, ok = true, status = 200): typeof fetch {
  return (async () => ({
    ok,
    status,
    json: async () => body,
  })) as unknown as typeof fetch;
}

/** 造一个永远不返回的 fetch，用来验证超时 */
function hangingFetch(): typeof fetch {
  return ((_url: string, init?: RequestInit) =>
    new Promise((_resolve, reject) => {
      init?.signal?.addEventListener('abort', () => {
        const e = new Error('aborted');
        e.name = 'AbortError';
        reject(e);
      });
    })) as unknown as typeof fetch;
}

const ORIGIN = 'https://gw.example.com';

beforeEach(() => {
  useSettings.setState({
    kind: 'openai-compatible',
    baseUrl: '',
    apiKey: '',
    model: '',
    enableRunScript: true,
    provision: { mode: 'byok', user: '', visibility: 1, status: 'ready' },
  });
});

describe('readInjectedUser', () => {
  it('从 ?u= 读出安装时注入的身份', () => {
    expect(readInjectedUser('?u=zhangsan')).toBe('zhangsan');
  });

  it('域账号里的空格和中文经过 URL 编码后也能读回来', () => {
    expect(readInjectedUser('?u=' + encodeURIComponent('张三 li'))).toBe('张三 li');
  });

  it('没有 u 参数时返回空串，不抛异常', () => {
    expect(readInjectedUser('')).toBe('');
    expect(readInjectedUser('?x=1')).toBe('');
  });
});

describe('parseConfig', () => {
  it('managed 必须带 baseUrl 和 model', () => {
    expect(parseConfig({ mode: 'managed', baseUrl: 'http://m/v1', model: 'qwen' })).toMatchObject({
      mode: 'managed',
      baseUrl: 'http://m/v1',
      model: 'qwen',
    });
  });

  it('managed 但缺 baseUrl 或 model 一律判为无效', () => {
    // 拿着一份空配置去撞墙，表现是"AI 一用就报奇怪的错"，排查要绕一大圈
    expect(parseConfig({ mode: 'managed', model: 'qwen' })).toBeNull();
    expect(parseConfig({ mode: 'managed', baseUrl: 'http://m/v1' })).toBeNull();
    expect(parseConfig({ mode: 'managed', baseUrl: '   ', model: 'qwen' })).toBeNull();
  });

  it('byok 原样识别', () => {
    expect(parseConfig({ mode: 'byok' })).toEqual({ mode: 'byok', visibility: 1 });
  });

  it('畸形内容返回 null 而不是抛异常', () => {
    expect(parseConfig(null)).toBeNull();
    expect(parseConfig('byok')).toBeNull();
    expect(parseConfig({ mode: 'whatever' })).toBeNull();
    expect(parseConfig({})).toBeNull();
  });

  it('enableRunScript 只接受布尔值', () => {
    const c = parseConfig({
      mode: 'managed',
      baseUrl: 'http://m/v1',
      model: 'q',
      enableRunScript: 'false',
    });
    // 字符串 'false' 是真值，直接用会把"关掉"变成"打开"
    expect(c && 'enableRunScript' in c ? c.enableRunScript : undefined).toBeUndefined();
  });
});

describe('fetchAiConfig', () => {
  it('命中白名单返回 managed 和完整配置', async () => {
    const r = await fetchAiConfig(
      'zhangsan',
      ORIGIN,
      fakeFetch({ mode: 'managed', baseUrl: 'http://m/v1', model: 'qwen' }),
    );
    expect(r.mode).toBe('managed');
    if (r.mode === 'managed') expect(r.config.model).toBe('qwen');
  });

  it('未命中返回 byok', async () => {
    const r = await fetchAiConfig('lisi', ORIGIN, fakeFetch({ mode: 'byok' }));
    expect(r.mode).toBe('byok');
  });

  it('没有身份时直接 byok，不发请求', async () => {
    const spy = vi.fn();
    const r = await fetchAiConfig('', ORIGIN, spy as unknown as typeof fetch);
    expect(r.mode).toBe('byok');
    expect(spy).not.toHaveBeenCalled();
  });

  //--------------------------------------------------------------------------
  // 【铁律：网关挂了也不能让任务窗格不能用】
  // 下面几条守的是同一件事——任何形式的网关故障都必须安静退回 byok，
  // 用户自己填地址照样能用。绝不能卡住、绝不能抛到界面上。
  //--------------------------------------------------------------------------
  it('网关返回 500 时退回 byok', async () => {
    const r = await fetchAiConfig('z', ORIGIN, fakeFetch({}, false, 500));
    expect(r.mode).toBe('byok');
    if (r.mode === 'byok') expect(r.reason).toContain('500');
  });

  it('网关返回畸形 JSON 时退回 byok', async () => {
    const r = await fetchAiConfig('z', ORIGIN, fakeFetch({ mode: 'managed' }));
    expect(r.mode).toBe('byok');
  });

  it('网关连不上时退回 byok', async () => {
    const boom = (async () => {
      throw new Error('ECONNREFUSED');
    }) as unknown as typeof fetch;
    const r = await fetchAiConfig('z', ORIGIN, boom);
    expect(r.mode).toBe('byok');
  });

  it('网关不响应时会超时退回 byok，不会一直挂着', async () => {
    vi.useFakeTimers();
    const p = fetchAiConfig('z', ORIGIN, hangingFetch());
    await vi.advanceTimersByTimeAsync(5000);
    const r = await p;
    vi.useRealTimers();
    expect(r.mode).toBe('byok');
    if (r.mode === 'byok') expect(r.reason).toBe('网关没响应');
  });
});

describe('provision 把结果落到 settings', () => {
  it('managed 会写入 IT 下发的地址与模型', async () => {
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: fakeFetch({
        mode: 'managed',
        baseUrl: 'http://llm.corp/v1',
        model: 'qwen2.5-32b',
        apiKey: 'sk-it',
        managedBy: 'IT 运维',
      }),
    });
    const s = useSettings.getState();
    expect(s.provision.mode).toBe('managed');
    expect(s.baseUrl).toBe('http://llm.corp/v1');
    expect(s.model).toBe('qwen2.5-32b');
    expect(s.apiKey).toBe('sk-it');
    expect(s.provision.managedBy).toBe('IT 运维');
  });

  it('IT 可以集中关掉 run_script', async () => {
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: fakeFetch({
        mode: 'managed',
        baseUrl: 'http://llm.corp/v1',
        model: 'q',
        enableRunScript: false,
      }),
    });
    expect(useSettings.getState().enableRunScript).toBe(false);
  });

  it('没下发 enableRunScript 时保持用户本地的选择', async () => {
    useSettings.setState({ enableRunScript: false });
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: fakeFetch({ mode: 'managed', baseUrl: 'http://llm.corp/v1', model: 'q' }) });
    expect(useSettings.getState().enableRunScript).toBe(false);
  });

  it('byok 不会动用户已经填好的配置', async () => {
    useSettings.setState({ baseUrl: 'http://my-own/v1', model: 'mine' });
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: fakeFetch({ mode: 'byok' }) });
    const s = useSettings.getState();
    expect(s.provision.mode).toBe('byok');
    expect(s.baseUrl).toBe('http://my-own/v1');
    expect(s.model).toBe('mine');
  });
});

//============================================================================
// 可见性开关（服务端下发）
//
// 0 不可见 / 1 可见可使用 / 2 可见但置灰
//============================================================================
describe('可见性开关', () => {
  it('只接受 0/1/2，别的一律按 1', () => {
    expect(parseVisibility(0)).toBe(0);
    expect(parseVisibility(1)).toBe(1);
    expect(parseVisibility(2)).toBe(2);
    // 【字符串 "0" 是真值】：服务端把数字写成字符串是很容易犯的错，
    // 不校验的话会把"不可见"变成"可用"——开关方向反了，最糟的一种 bug
    expect(parseVisibility('0')).toBe(1);
    expect(parseVisibility(3)).toBe(1);
    expect(parseVisibility(null)).toBe(1);
    expect(parseVisibility(undefined)).toBe(1);
  });

  it('managed 用户的可见性被下发到 settings', async () => {
    await provision({
      user: 'z',
      origin: ORIGIN,
      fetchImpl: fakeFetch({
        visibility: 2,
        mode: 'managed',
        baseUrl: 'http://llm.corp/v1',
        model: 'q',
      }),
    });
    expect(useSettings.getState().provision.visibility).toBe(2);
  });

  it('【不在白名单的人也要受可见性管控】', async () => {
    // 可见性和白名单是两件事：白名单管"用谁的模型"，
    // 可见性管"能不能用"。只对白名单内生效的话，
    // 把功能关掉之后大多数人照样能用，开关等于没有。
    await provision({
      user: 'z',
      origin: ORIGIN,
      fetchImpl: fakeFetch({ visibility: 0, mode: 'byok' }),
    });
    const s = useSettings.getState();
    expect(s.provision.mode).toBe('byok');
    expect(s.provision.visibility).toBe(0);
  });

  it('网关挂掉时按"可用"处理，不是按"停用"', async () => {
    // 这是治理开关不是安全闸。网关一抖动就让全公司用不了 AI，
    // 代价比"多开了一会儿"大得多。真要强管控用 0 在安装侧卡死。
    const boom = (async () => {
      throw new Error('ECONNREFUSED');
    }) as unknown as typeof fetch;
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: boom });
    expect(useSettings.getState().provision.visibility).toBe(1);
  });

  it('没配可见性时默认可用', async () => {
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: fakeFetch({ mode: 'byok' }) });
    expect(useSettings.getState().provision.visibility).toBe(1);
  });
});

//============================================================================
// provision 完成之前不能放行
//
// 【这是 Codex 评审发现的治理开关绕过】：原先界面先挂出来、
// provision 在后台异步跑、默认按"可用"处理——服务端已经关掉了功能，
// 用户却能在那几百毫秒里正常发消息。
//============================================================================
describe('provision 未完成期间不放行', () => {
  it('初始状态是 pending，不是"可用"', () => {
    useSettings.setState({
      provision: { mode: 'byok', user: '', visibility: 1, status: 'pending' },
    });
    expect(useSettings.getState().provision.status).toBe('pending');
  });

  it('provision 期间状态一直是 pending，问完才变 ready', async () => {
    useSettings.setState({
      provision: { mode: 'byok', user: '', visibility: 1, status: 'pending' },
    });

    // 造一个可以人工控制何时返回的 fetch
    let release: (v: unknown) => void = () => {};
    const gate = new Promise((r) => (release = r));
    const slowFetch = (async () => {
      await gate;
      return { ok: true, status: 200, json: async () => ({ visibility: 0, mode: 'byok' }) };
    }) as unknown as typeof fetch;

    const p = provision({ user: 'z', origin: ORIGIN, fetchImpl: slowFetch });

    // 网关还没回话：必须仍然是 pending，界面据此拦住发送
    expect(useSettings.getState().provision.status).toBe('pending');

    release(null);
    await p;

    const s = useSettings.getState();
    expect(s.provision.status).toBe('ready');
    expect(s.provision.visibility).toBe(0);
  });

  it('网关失败也要把状态置为 ready，不能永远卡在 pending', async () => {
    useSettings.setState({
      provision: { mode: 'byok', user: '', visibility: 1, status: 'pending' },
    });
    const boom = (async () => {
      throw new Error('ECONNREFUSED');
    }) as unknown as typeof fetch;
    await provision({ user: 'z', origin: ORIGIN, fetchImpl: boom });
    // 卡在 pending 的话界面就永远不放行了，那是另一种故障
    expect(useSettings.getState().provision.status).toBe('ready');
  });

  it('意外异常也不能把界面卡在 pending', async () => {
    useSettings.setState({
      provision: { mode: 'byok', user: '', visibility: 1, status: 'pending' },
    });
    // 造一个连 json() 都会炸的响应，绕过 fetchAiConfig 的常规兜底
    const weird = (async () => ({
      ok: true,
      status: 200,
      get json() {
        throw new Error('boom');
      },
    })) as unknown as typeof fetch;

    await provision({ user: 'z', origin: ORIGIN, fetchImpl: weird });
    expect(useSettings.getState().provision.status).toBe('ready');
  });
});
