import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import type { Provider, ProviderConfig } from '../llm/types';
import { createOpenAICompatibleProvider } from '../llm/providers/openaiCompatible';
import { createAnthropicProvider } from '../llm/providers/anthropic';

/**
 * 内网部署时把默认值改成公司的模型网关地址，用户打开即用、无需填 Key。
 * 构建期可用 Vite 的 define 注入，或部署时改这里的常量。
 */
export const PRESETS: Array<{
  id: string;
  label: string;
  kind: ProviderConfig['kind'];
  baseUrl: string;
  model: string;
  hint: string;
}> = [
  {
    id: 'intranet',
    label: '内网模型网关',
    kind: 'openai-compatible',
    baseUrl: 'https://模型网关地址/v1',
    model: '',
    hint: '公司内网部署的 vLLM / Xinference / One-API 等，走 OpenAI 兼容协议',
  },
  {
    id: 'ollama',
    label: 'Ollama（本机）',
    kind: 'openai-compatible',
    baseUrl: 'http://localhost:11434/v1',
    model: 'qwen2.5:14b',
    hint: '本机运行的 Ollama，无需 API Key',
  },
  {
    id: 'vllm',
    label: 'vLLM',
    kind: 'openai-compatible',
    baseUrl: 'http://<服务器IP>:8000/v1',
    model: '',
    hint: 'vLLM 的 OpenAI 兼容端点',
  },
  {
    id: 'openai',
    label: 'OpenAI',
    kind: 'openai-compatible',
    baseUrl: 'https://api.openai.com/v1',
    model: 'gpt-4o',
    hint: '需要出网权限',
  },
  {
    id: 'anthropic',
    label: 'Anthropic',
    kind: 'anthropic',
    baseUrl: 'https://api.anthropic.com/v1',
    model: 'claude-sonnet-5',
    hint: '需要出网权限',
  },
  {
    id: 'deepseek',
    label: 'DeepSeek',
    kind: 'openai-compatible',
    baseUrl: 'https://api.deepseek.com/v1',
    model: 'deepseek-chat',
    hint: '需要出网权限',
  },
];

/** 白名单分流的结果。详见 store/provisioning.ts */
export interface ProvisionState {
  /** managed = 用 IT 配好的模型；byok = 用户自己配 */
  mode: 'managed' | 'byok';
  /** 安装时注入的身份（`?u=`）。空串表示没读到。 */
  user: string;
  /** 走到 byok 的原因，只用于提示和排查 */
  reason?: string;
  /** 这份配置是谁下发的，显示用 */
  managedBy?: string;
  /**
   * IT 是否明确下发了 enableRunScript。
   * 下发了就不许用户在界面上改回来——否则"IT 能集中关掉"是句空话。
   */
  runScriptPinned?: boolean;
  /**
   * 服务端下发的可见性：0 不可见 / 1 可见可使用 / 2 可见但置灰。
   * 读不到时按 1，详见 provisioning.ts 的说明。
   */
  visibility: 0 | 1 | 2;
  /**
   * 有没有问过网关。
   *
   * 【必须有这个状态】。原先界面先挂出来、provision 在后台异步跑，
   * 默认 visibility=1 —— 那就留出了一个"开关还没生效、用户已经能聊天"
   * 的窗口，治理开关可以被绕过（Codex 评审发现）。
   * pending 期间一律按不可用处理，问完才放行。
   */
  status: 'pending' | 'ready';
}

export interface SettingsState extends ProviderConfig {
  systemAddition: string;
  enableRunScript: boolean;
  /** 已知的模型列表，来自「测试连接」时拉取 */
  knownModels: string[];
  /**
   * 启动时由 provision() 写入。
   *
   * 【初始是 pending，不等于"可用"】。界面必须等 status 变成 ready
   * 才放行——否则服务端已经关掉功能的用户能在那个窗口里正常使用。
   */
  provision: ProvisionState;
  set(patch: Partial<SettingsState>): void;
  setProvision(p: ProvisionState): void;
  applyPreset(id: string): void;
}

export const useSettings = create<SettingsState>()(
  persist(
    (set) => ({
      kind: 'openai-compatible',
      baseUrl: '',
      apiKey: '',
      model: '',
      temperature: 0.2,
      maxTokens: 4096,
      systemAddition: '',
      enableRunScript: true,
      knownModels: [],
      provision: { mode: 'byok', user: '', visibility: 1, status: 'pending' },
      set: (patch) => set(patch),
      setProvision: (p) => set({ provision: p }),
      applyPreset: (id) => {
        const p = PRESETS.find((x) => x.id === id);
        if (p) set({ kind: p.kind, baseUrl: p.baseUrl, model: p.model, knownModels: [] });
      },
    }),
    {
      name: 'excel-ai.settings',
      // API Key 存 localStorage（任务窗格的隔离 origin 内）。
      // Sidecar 就绪后应迁移到 OS keychain，设置页已标注当前存储位置。

      /**
       * 【managed 模式下的配置一律不落盘】。
       *
       * IT 下发的地址、模型和 key 只活在内存里，每次打开重新问网关。
       * 否则把某个人从白名单里移出去之后，他本机还留着一份能用的配置，
       * 管控就是假的——而且那份 key 会一直躺在 localStorage 里。
       *
       * provision 本身也不存：它是每次启动现问出来的结果，
       * 存下来只会在网关不可达时给出一个过期的"你在白名单里"。
       */
      partialize: (s) => {
        const { provision, ...rest } = s;
        if (provision.mode === 'managed') {
          const { baseUrl, apiKey, model, ...safe } = rest;
          return safe;
        }
        return rest;
      },
    },
  ),
);

export function createProvider(cfg: ProviderConfig): Provider {
  return cfg.kind === 'anthropic'
    ? createAnthropicProvider(cfg)
    : createOpenAICompatibleProvider(cfg);
}

/**
 * 测试连接：不仅要能连通，还要验证端点**真的支持 function calling**。
 * 内网自建网关常见的坑是模型不支持工具调用，等到实际对话才发现就太晚了。
 */
export async function testConnection(
  cfg: ProviderConfig,
): Promise<{ ok: boolean; message: string; models: string[] }> {
  if (!cfg.baseUrl.trim()) return { ok: false, message: '请先填写接口地址。', models: [] };
  if (!cfg.model.trim()) return { ok: false, message: '请先填写模型名称。', models: [] };

  const provider = createProvider(cfg);
  const models = await provider.listModels().catch(() => []);

  try {
    let sawToolCall = false;
    let sawText = false;

    for await (const d of provider.chat({
      model: cfg.model,
      temperature: 0,
      maxTokens: 256,
      messages: [
        { role: 'user', content: '请调用 ping 工具，参数 msg 填 "hello"。' },
      ],
      tools: [
        {
          name: 'ping',
          description: '连通性自检工具，收到调用即视为成功',
          schema: {
            type: 'object',
            properties: { msg: { type: 'string' } },
            required: ['msg'],
          },
        },
      ],
    })) {
      if (d.type === 'tool_call_start') sawToolCall = true;
      if (d.type === 'text') sawText = true;
      if (d.type === 'error') return { ok: false, message: d.message, models };
    }

    if (sawToolCall) {
      return {
        ok: true,
        message: `连接成功，且该模型支持工具调用。${models.length ? `可用模型 ${models.length} 个。` : ''}`,
        models,
      };
    }
    if (sawText) {
      return {
        ok: false,
        message:
          '能连通，但模型没有发起工具调用 —— 该端点或模型可能不支持 function calling。' +
          '本插件依赖工具调用才能操作 Excel，请换一个支持的模型。',
        models,
      };
    }
    return { ok: false, message: '连接成功但没有收到任何响应内容，请检查模型名称。', models };
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    // CORS 在浏览器里只表现为 TypeError，需要特别提示，否则用户会一头雾水
    const hint = /Failed to fetch|NetworkError|TypeError/.test(msg)
      ? '\n可能原因：地址不可达，或服务端未开启 CORS（需允许任务窗格的来源）。'
      : '';
    return { ok: false, message: `${msg}${hint}`, models };
  }
}
