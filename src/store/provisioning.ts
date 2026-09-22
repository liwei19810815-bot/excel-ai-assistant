/**
 * 白名单分流（provisioning）。
 *
 * 任务窗格启动时问一次网关："当前这个人该用公司配好的模型，还是自己配？"
 *
 * 身份怎么来的：Office.js 的任务窗格是沙箱网页，**读不到 Windows 用户名**
 * （Excel/Word/PPT 的 Office.js 都没有用户身份 API，只有 Outlook 有）。
 * 所以身份是安装时注入的——安装脚本按 %USERNAME% 生成该用户专属的
 * manifest，把 `?u=<用户名>` 放进 SourceLocation。这里只负责读回来。
 *
 * ⚠ **这是分流，不是鉴权。**
 * 用户可以编辑本机的 manifest.xml 把 `?u=` 改成别人的名字。
 * 这是已确认接受的取舍，前提是 IT 配的那个模型本身不怕被多用几个人
 * （没有敏感数据访问权、成本可控）。
 * 一旦模型涉及成本分摊或能读到敏感数据，必须改成网关侧的 Windows 集成
 * 认证（Kerberos/Negotiate），由服务器从票据里解析域账号。
 *
 * 【铁律：网关挂了也不能让任务窗格不能用】
 * 接口超时、返回畸形、网关根本没部署——一律安静退回 byok（用户自己配），
 * 绝不卡住启动、绝不让用户对着一个转圈的界面。
 */

import { useSettings } from './settings';

/** 网关配置接口的超时。宁可早点放弃，也不要让用户干等。 */
const CONFIG_TIMEOUT_MS = 3000;

export type ProvisionMode = 'managed' | 'byok';

/**
 * AI 功能的可见性，由服务端下发。和白名单是两件事：
 *   白名单  = 用公司配好的模型还是自己配
 *   可见性  = 能不能看到、能不能用
 *
 *   0 不可见     —— 安装时就不注册加载项。【这是唯一能真正"看不见"的办法】：
 *                   Office.js 没有隐藏内置选项卡上按钮的能力。
 *                   对已经装上的机器，运行时按"停用"处理，下次跑安装程序时移除。
 *   1 可见可使用 —— 默认
 *   2 可见但置灰 —— 按钮还在，但点开只告诉用户已停用
 *
 * 【读不到就按 1】。这是治理开关不是安全闸：网关挂了就让全公司用不了 AI，
 * 代价比"多开了一会儿"大得多。真要强管控，用 0 在安装侧卡死。
 */
export type Visibility = 0 | 1 | 2;

export const VISIBILITY_HIDDEN: Visibility = 0;
export const VISIBILITY_ENABLED: Visibility = 1;
export const VISIBILITY_DISABLED: Visibility = 2;

export interface ManagedConfig {
  mode: 'managed';
  visibility?: Visibility;
  baseUrl: string;
  model: string;
  apiKey?: string;
  /** IT 可以集中关掉 run_script（模型现写代码执行那个工具） */
  enableRunScript?: boolean;
  /** 显示用：这份配置是谁下发的 */
  managedBy?: string;
}

export interface ByokConfig {
  mode: 'byok';
  visibility?: Visibility;
}

export type AiConfigResponse = ManagedConfig | ByokConfig;

export type ProvisionResult =
  | { mode: 'managed'; config: ManagedConfig; user: string; visibility: Visibility }
  /** reason 只用于设置页提示和排查，不影响功能 */
  | { mode: 'byok'; reason?: string; user?: string; visibility: Visibility };

/** 从 URL 取安装时注入的身份。取不到返回空串。 */
export function readInjectedUser(search: string = typeof location === 'undefined' ? '' : location.search): string {
  try {
    return new URLSearchParams(search).get('u')?.trim() ?? '';
  } catch {
    return '';
  }
}

/**
 * 向网关问一次配置。
 *
 * 【所有失败都收敛成 byok】，调用方不需要写 try/catch。
 */
export async function fetchAiConfig(
  user: string,
  origin: string = typeof location === 'undefined' ? '' : location.origin,
  fetchImpl: typeof fetch = fetch,
): Promise<ProvisionResult> {
  if (!user) {
    return { mode: 'byok', visibility: 1, reason: '没有从 URL 里读到身份（manifest 可能不是安装程序生成的）' };
  }

  // AbortController 做超时：网关不可达时 fetch 默认会等很久，
  // 那段时间用户对着的是一个什么都不能做的界面。
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), CONFIG_TIMEOUT_MS);

  try {
    const url = `${origin}/api/ai-config?u=${encodeURIComponent(user)}`;
    const resp = await fetchImpl(url, { signal: ctl.signal });

    if (!resp.ok) {
      return { mode: 'byok', visibility: 1, reason: `网关返回 ${resp.status}`, user };
    }

    const data = (await resp.json()) as unknown;
    const parsed = parseConfig(data);
    if (!parsed) {
      return { mode: 'byok', visibility: 1, reason: '网关返回的内容看不懂', user };
    }
    if (parsed.mode === 'byok') {
      return { mode: 'byok', visibility: parsed.visibility ?? 1, reason: '不在白名单里', user };
    }
    return { mode: 'managed', config: parsed, user, visibility: parsed.visibility ?? 1 };
  } catch (e) {
    // 超时、断网、CORS、网关没部署……对用户来说都一样：自己配就是了
    const why = (e as Error)?.name === 'AbortError' ? '网关没响应' : '连不上网关';
    return { mode: 'byok', visibility: 1, reason: why, user };
  } finally {
    clearTimeout(timer);
  }
}

/**
 * 校验并收窄网关返回的内容。
 *
 * 【必须严格校验】：这些值会直接拿去发请求。网关配错了、返回了半截 JSON，
 * 直接用下去的表现是"AI 一用就报奇怪的错"，排查要绕一大圈。
 * 在这里拦住，退回 byok 让用户自己填，至少是个能走通的路。
 */
export function parseConfig(data: unknown): AiConfigResponse | null {
  if (!data || typeof data !== 'object') return null;
  const d = data as Record<string, unknown>;

  const visibility = parseVisibility(d.visibility);

  if (d.mode === 'byok') return { mode: 'byok', visibility };

  if (d.mode === 'managed') {
    const baseUrl = typeof d.baseUrl === 'string' ? d.baseUrl.trim() : '';
    const model = typeof d.model === 'string' ? d.model.trim() : '';
    // managed 却没给地址或模型名，等于没配——当作没命中处理，别让用户
    // 拿着一份空配置去撞墙
    if (!baseUrl || !model) return null;

    return {
      mode: 'managed',
      visibility,
      baseUrl,
      model,
      apiKey: typeof d.apiKey === 'string' ? d.apiKey : undefined,
      enableRunScript: typeof d.enableRunScript === 'boolean' ? d.enableRunScript : undefined,
      managedBy: typeof d.managedBy === 'string' ? d.managedBy : undefined,
    };
  }

  return null;
}

/**
 * 【只接受 0/1/2，别的一律按 1】。
 * 服务端把 visibility 写成字符串 "0" 是个很容易犯的错，
 * 而字符串 "0" 是真值——不校验的话会把"不可见"变成"可用"。
 */
export function parseVisibility(v: unknown): Visibility {
  return v === 0 || v === 1 || v === 2 ? v : 1;
}

/**
 * 启动时跑一次：读身份 → 问网关 → 命中就把配置装进 settings。
 *
 * 【managed 的值不落 localStorage】。IT 下发的地址和 key 只活在内存里，
 * 下次打开重新问网关。否则把人从白名单移出去之后，他本机还留着一份
 * 能用的配置，管控就是假的。
 */
export interface ProvisionOptions {
  origin?: string;
  fetchImpl?: typeof fetch;
  /** 身份。不传就从 URL 读——留这个口子是为了能在没有 location 的环境里测。 */
  user?: string;
}

export async function provision(opts: ProvisionOptions = {}): Promise<ProvisionResult> {
  try {
    return await provisionInner(opts);
  } catch (e) {
    // 【任何意外都必须把状态推进到 ready】。卡在 pending 的话，
    // 界面就永远停在"正在获取配置"——用户不是被停用，是彻底用不了。
    // fetchAiConfig 内部已经吞掉了所有网络异常，这里兜的是它之外的意外
    // （比如 store 本身出问题）。
    try {
      useSettings.getState().setProvision({
        mode: 'byok',
        user: '',
        visibility: 1,
        status: 'ready',
        reason: '读取配置时出错，已退回自行配置',
      });
    } catch {
      /* 连 store 都写不了就真没辙了，但至少别把异常抛到界面上 */
    }
    return { mode: 'byok', visibility: 1, reason: (e as Error)?.message };
  }
}

async function provisionInner(opts: ProvisionOptions): Promise<ProvisionResult> {
  const { origin, fetchImpl } = opts;
  const user = opts.user ?? readInjectedUser();
  const result = await fetchAiConfig(user, origin, fetchImpl);
  const s = useSettings.getState();

  if (result.mode === 'byok') {
    s.setProvision({
      mode: 'byok',
      user: result.user ?? '',
      reason: result.reason,
      visibility: result.visibility,
      status: 'ready',
    });
    return result;
  }

  const c = result.config;
  s.setProvision({
    mode: 'managed',
    user: result.user,
    managedBy: c.managedBy,
    runScriptPinned: typeof c.enableRunScript === 'boolean',
    visibility: result.visibility,
    status: 'ready',
  });
  s.set({
    kind: 'openai-compatible',
    baseUrl: c.baseUrl,
    model: c.model,
    apiKey: c.apiKey ?? '',
    // IT 说关就关。没下发这一项时保持用户本地的选择。
    ...(typeof c.enableRunScript === 'boolean' ? { enableRunScript: c.enableRunScript } : {}),
  });

  return result;
}

/**
 * 按可见性把功能区按钮置灰。
 *
 * 【这是尽力而为，不是强制点】，有两条实打实的限制：
 *
 *   1. 没有共享运行时（shared runtime）时，这段代码只有在用户
 *      【点开过任务窗格】之后才会跑。也就是说第一次点开之前，
 *      按钮看起来是正常的。
 *   2. Office.js **没有隐藏**内置选项卡上按钮的能力。
 *      requestUpdate 只能改 enabled，改不了 visible。
 *      真要"看不见"，只能在安装时不注册这个加载项（visibility=0 的做法）。
 *
 * 所以真正的强制落在任务窗格里（ChatPane 直接不渲染聊天界面）。
 * 这里只是让按钮的样子和实际状态对上，免得用户反复去点一个点了没用的按钮。
 *
 * 【失败必须静默】：RibbonApi 在老版本 Office 上不存在，
 * 为了一个视觉效果把任务窗格搞崩是不划算的。
 */
export async function applyRibbonVisibility(visibility: Visibility): Promise<boolean> {
  try {
    const off = (globalThis as { Office?: any }).Office;
    if (!off?.ribbon?.requestUpdate) return false;
    if (!off.context?.requirements?.isSetSupported?.('RibbonApi', '1.1')) return false;

    // 【manifest 里声明了几个按钮，这里就要列几个】。
    // 只置灰 OpenPane 的话，AskSelection 还亮着——用户点它会打开窗格、
    // 看到"已停用"，一个亮着却什么也做不了的按钮比灰掉更让人困惑。
    // 加按钮时别忘了同步这里；check-consistency 有一条断言守着这件事。
    const enabled = visibility === 1;
    await off.ribbon.requestUpdate({
      tabs: [
        {
          id: 'TabHome',
          groups: [
            {
              id: 'ExcelAI.Group',
              controls: [
                { id: 'ExcelAI.OpenPane', enabled },
                { id: 'ExcelAI.AskSelection', enabled },
              ],
            },
          ],
        },
      ],
    });
    return true;
  } catch {
    return false;
  }
}
