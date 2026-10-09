/** 白名单分流：启动时向网关询问当前用户应使用托管模型还是 BYOK。 */

import { useSettings } from './settings';

const CONFIG_TIMEOUT_MS = 3000;

export type ProvisionMode = 'managed' | 'byok';

export interface ManagedConfig {
  mode: 'managed';
  baseUrl: string;
  model: string;
  apiKey?: string;
  enableRunScript?: boolean;
  managedBy?: string;
}

export interface ByokConfig { mode: 'byok' }
export type AiConfigResponse = ManagedConfig | ByokConfig;
export type ProvisionResult =
  | { mode: 'managed'; config: ManagedConfig; user: string }
  | { mode: 'byok'; reason?: string; user?: string };

export function readInjectedUser(search: string = typeof location === 'undefined' ? '' : location.search): string {
  try { return new URLSearchParams(search).get('u')?.trim() ?? ''; } catch { return ''; }
}

export async function fetchAiConfig(
  user: string,
  origin: string = typeof location === 'undefined' ? '' : location.origin,
  fetchImpl: typeof fetch = fetch,
): Promise<ProvisionResult> {
  if (!user) return { mode: 'byok', reason: '没有从 URL 里读到身份（manifest 可能不是安装程序生成的）' };

  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), CONFIG_TIMEOUT_MS);
  try {
    const resp = await fetchImpl(`${origin}/api/ai-config?u=${encodeURIComponent(user)}`, { signal: ctl.signal });
    if (!resp.ok) return { mode: 'byok', reason: `网关返回 ${resp.status}`, user };
    const parsed = parseConfig((await resp.json()) as unknown);
    if (!parsed) return { mode: 'byok', reason: '网关返回的内容看不懂', user };
    if (parsed.mode === 'byok') return { mode: 'byok', reason: '不在白名单里', user };
    return { mode: 'managed', config: parsed, user };
  } catch (e) {
    return { mode: 'byok', reason: (e as Error)?.name === 'AbortError' ? '网关没响应' : '连不上网关', user };
  } finally { clearTimeout(timer); }
}

export function parseConfig(data: unknown): AiConfigResponse | null {
  if (!data || typeof data !== 'object') return null;
  const d = data as Record<string, unknown>;
  if (d.mode === 'byok') return { mode: 'byok' };
  if (d.mode !== 'managed') return null;
  const baseUrl = typeof d.baseUrl === 'string' ? d.baseUrl.trim() : '';
  const model = typeof d.model === 'string' ? d.model.trim() : '';
  if (!baseUrl || !model) return null;
  return {
    mode: 'managed', baseUrl, model,
    apiKey: typeof d.apiKey === 'string' ? d.apiKey : undefined,
    enableRunScript: typeof d.enableRunScript === 'boolean' ? d.enableRunScript : undefined,
    managedBy: typeof d.managedBy === 'string' ? d.managedBy : undefined,
  };
}

export interface ProvisionOptions { origin?: string; fetchImpl?: typeof fetch; user?: string }

export async function provision(opts: ProvisionOptions = {}): Promise<ProvisionResult> {
  try { return await provisionInner(opts); } catch (e) {
    try {
      useSettings.getState().setProvision({ mode: 'byok', user: '', status: 'ready', reason: '读取配置时出错，已退回自行配置' });
    } catch { /* 启动链路不能因为配置异常崩溃 */ }
    return { mode: 'byok', reason: (e as Error)?.message };
  }
}

async function provisionInner(opts: ProvisionOptions): Promise<ProvisionResult> {
  const result = await fetchAiConfig(opts.user ?? readInjectedUser(), opts.origin, opts.fetchImpl);
  const s = useSettings.getState();
  if (result.mode === 'byok') {
    s.setProvision({ mode: 'byok', user: result.user ?? '', reason: result.reason, status: 'ready' });
    return result;
  }
  const c = result.config;
  s.setProvision({ mode: 'managed', user: result.user, managedBy: c.managedBy, runScriptPinned: typeof c.enableRunScript === 'boolean', status: 'ready' });
  s.set({ kind: 'openai-compatible', baseUrl: c.baseUrl, model: c.model, apiKey: c.apiKey ?? '', ...(typeof c.enableRunScript === 'boolean' ? { enableRunScript: c.enableRunScript } : {}) });
  return result;
}
