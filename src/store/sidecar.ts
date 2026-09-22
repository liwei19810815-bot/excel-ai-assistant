/**
 * sidecar 伴生进程的客户端。
 *
 * sidecar 是本机上的一个小进程，干的是 Office.js【做不到】的事
 * （刷新 Power Query、数据模型、调用工作簿里的宏）。
 * 它由工具箱的 .xlam 在 Excel 启动时拉起，关 Excel 就跟着退出。
 *
 * ============================================================
 * 【铁律：sidecar 不在 ≠ AI 坏了】
 * ============================================================
 * 大多数用户根本没装这个组件。所以这里【任何失败都必须安静降级】：
 *   · 不弹窗、不报错、不拦着用户用别的功能
 *   · sidecar 类工具直接从工具列表里消失（模型看不到它们，
 *     就不会承诺做不到的事）
 * 和"网关挂了任务窗格照常可用"是同一条铁律。
 *
 * ============================================================
 * 【令牌只能放请求头，不能放查询串】
 * ============================================================
 * 自定义头 X-Toolbox-Token 会强制浏览器先发 OPTIONS 预检，
 * 而预检只对白名单 Origin 放行——远端网页因此【根本发不出】这个请求。
 * 换成 ?token=xxx 就成了简单请求，不触发预检，这道门等于自己拆了。
 * sidecar 那边对查询串里的令牌一律回 400，两边是一致的。
 */

/** 令牌走这个头。改名字要同步改 sidecar 那边。 */
export const SIDECAR_TOKEN_HEADER = 'X-Toolbox-Token';

/** 默认端口，和安装器写进配置的那个一致。 */
export const SIDECAR_BASE_PORT = 8899;

/**
 * 往后探几个端口。
 * sidecar 遇到端口冲突会顺延，所以不能只认一个；
 * 但范围也不能大——每多一个都是一次真实的网络往返，拖慢启动。
 */
export const SIDECAR_PORT_SPAN = 5;

/** 单次探测的超时。本机回环，正常是毫秒级；给到 1 秒已经很宽。 */
const PROBE_TIMEOUT_MS = 1000;

export interface SidecarStatus {
  available: boolean;
  /** 探到的端口；没探到是 0 */
  port: number;
  version: string;
  /** 给日志和"为什么不可用"提示用，不展示给普通用户 */
  reason: string;
}

const UNAVAILABLE = (reason: string): SidecarStatus => ({
  available: false,
  port: 0,
  version: '',
  reason,
});

/**
 * 启动时探到的状态，给 UI 和工具列表用。
 *
 * 【默认必须是"不可用"】。探测完成之前就把 sidecar 类工具放进列表的话，
 * 模型会在一个还不知道能不能用的能力上做承诺——
 * 宁可晚一点出现，也不要承诺了做不到。
 */
let cachedStatus: SidecarStatus = UNAVAILABLE('尚未探测');

export function setSidecarStatus(s: SidecarStatus): void {
  cachedStatus = s;
}

export function getSidecarStatus(): SidecarStatus {
  return cachedStatus;
}

/**
 * 从 URL 里读安装器注入的 sidecar 令牌。
 *
 * 复用 `?u=` 那条已验证可行的路子：Office.js 沙箱读不到本机的任何东西，
 * 令牌只能由安装器在生成 manifest 时拼进 SourceLocation。
 *
 * 【没装 sidecar 组件时安装器写的是空串】——这不是异常，是正常情况。
 */
export function readInjectedSidecarToken(
  search: string = typeof location === 'undefined' ? '' : location.search,
): string {
  try {
    return new URLSearchParams(search).get('sidecar')?.trim() ?? '';
  } catch {
    return '';
  }
}

/**
 * 探一次 sidecar 在不在。
 *
 * 【没有令牌就一个请求都不发】。绝大多数用户没装这个组件，
 * 对他们来说这里应该是零成本——白白发 5 个注定失败的请求，
 * 既拖慢启动，又会在控制台刷一片红色的连接错误，看着像出了故障。
 */
export async function probeSidecar(
  token: string = readInjectedSidecarToken(),
  opts: {
    basePort?: number;
    span?: number;
    fetchImpl?: typeof fetch;
    timeoutMs?: number;
  } = {},
): Promise<SidecarStatus> {
  if (!token) return UNAVAILABLE('没有注入 sidecar 令牌（多半是没装这个组件）');

  const basePort = opts.basePort ?? SIDECAR_BASE_PORT;
  const span = opts.span ?? SIDECAR_PORT_SPAN;
  const fetchImpl = opts.fetchImpl ?? fetch;
  const timeoutMs = opts.timeoutMs ?? PROBE_TIMEOUT_MS;

  for (let port = basePort; port < basePort + span; port++) {
    const res = await probeOne(port, token, fetchImpl, timeoutMs);
    if (res) return res;
  }
  return UNAVAILABLE(`端口 ${basePort}-${basePort + span - 1} 上没找到 sidecar`);
}

async function probeOne(
  port: number,
  token: string,
  fetchImpl: typeof fetch,
  timeoutMs: number,
): Promise<SidecarStatus | null> {
  // 【必须有超时】。端口被别的程序占着但不回包时，fetch 会一直挂着，
  // 启动流程就卡在这儿了——而用户看到的是任务窗格一直转圈。
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), timeoutMs);
  try {
    const resp = await fetchImpl(`http://127.0.0.1:${port}/health`, {
      method: 'GET',
      headers: { [SIDECAR_TOKEN_HEADER]: token },
      signal: ac.signal,
    });
    if (!resp.ok) return null;

    const data = (await resp.json()) as { ok?: boolean; name?: string; version?: string };

    // 【必须认名字】。这个端口上完全可能是【别的程序】——
    // 只要它对任意请求回 200，我们就会把它当成 sidecar，
    // 之后每一次调用都失败，而现象看起来像"sidecar 有问题"。
    if (data?.ok !== true || data?.name !== 'excel-toolbox-sidecar') return null;

    return {
      available: true,
      port,
      version: String(data.version ?? ''),
      reason: '',
    };
  } catch {
    // 连不上、超时、跨域被拦、JSON 坏了——全都只意味着"这个端口不是它"
    return null;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * 调 sidecar 的某个具体动作。
 *
 * 【白名单式】：只允许调已经实现的那几个路径，
 * 绝不做"把一段代码传进去执行"的通用口子——那会把风险从
 * "几个动作"放大成"任意代码执行"。
 */
export async function callSidecar<T>(
  status: SidecarStatus,
  path: string,
  token: string = readInjectedSidecarToken(),
  init: RequestInit = {},
  fetchImpl: typeof fetch = fetch,
): Promise<T> {
  if (!status.available) throw new Error('sidecar 未运行');
  if (!token) throw new Error('没有 sidecar 令牌');

  const resp = await fetchImpl(`http://127.0.0.1:${status.port}${path}`, {
    ...init,
    headers: {
      ...(init.headers ?? {}),
      [SIDECAR_TOKEN_HEADER]: token,
    },
  });

  if (!resp.ok) {
    throw new Error(`sidecar 返回 ${resp.status}`);
  }
  return (await resp.json()) as T;
}
