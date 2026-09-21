import type { ChatRequest, Delta, Provider, ProviderConfig } from '../types';
import { readSSE, throwIfNotOk } from '../sse';

/**
 * OpenAI Chat Completions 协议适配。
 * 内网部署的 vLLM / Ollama / Xinference / One-API 以及 DeepSeek、通义、智谱
 * 全都兼容这套协议，所以换厂商只需改 baseUrl + model，不动代码。
 */
export function createOpenAICompatibleProvider(cfg: ProviderConfig): Provider {
  const base = cfg.baseUrl.replace(/\/+$/, '');

  return {
    id: 'openai-compatible',

    async listModels() {
      try {
        const res = await fetch(`${base}/models`, { headers: authHeaders(cfg) });
        if (!res.ok) return [];
        const json = (await res.json()) as { data?: Array<{ id: string }> };
        return (json.data ?? []).map((m) => m.id).sort();
      } catch {
        // 内网部署常见：/models 未实现。不视为错误，设置页降级为手填。
        return [];
      }
    },

    async *chat(req: ChatRequest): AsyncIterable<Delta> {
      const res = await fetch(`${base}/chat/completions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders(cfg) },
        signal: req.signal,
        body: JSON.stringify({
          model: req.model,
          messages: req.messages.map(toWireMessage),
          tools: req.tools.length
            ? req.tools.map((t) => ({
                type: 'function',
                function: { name: t.name, description: t.description, parameters: t.schema },
              }))
            : undefined,
          tool_choice: req.tools.length ? 'auto' : undefined,
          temperature: req.temperature,
          max_tokens: req.maxTokens,
          stream: true,
        }),
      });

      await throwIfNotOk(res, 'OpenAI 兼容接口');

      // 流式分片里 tool_call 的 name 只在第一片出现，需按 index 累积
      const started = new Set<number>();
      let finish: 'stop' | 'tool_calls' | 'length' = 'stop';

      for await (const payload of readSSE(res, req.signal)) {
        if (payload === '[DONE]') break;

        let chunk: WireChunk;
        try {
          chunk = JSON.parse(payload);
        } catch {
          continue; // 心跳或非 JSON 注释行
        }

        const choice = chunk.choices?.[0];
        if (!choice) continue;

        const text = choice.delta?.content;
        if (text) yield { type: 'text', text };

        for (const tc of choice.delta?.tool_calls ?? []) {
          const idx = tc.index ?? 0;
          if (!started.has(idx)) {
            started.add(idx);
            yield {
              type: 'tool_call_start',
              id: tc.id ?? `call_${idx}`,
              name: tc.function?.name ?? '',
            };
          }
          if (tc.function?.arguments) {
            yield {
              type: 'tool_call_args',
              id: tc.id ?? `call_${idx}`,
              argsChunk: tc.function.arguments,
            };
          }
        }

        if (choice.finish_reason === 'tool_calls') finish = 'tool_calls';
        else if (choice.finish_reason === 'length') finish = 'length';
      }

      yield { type: 'done', reason: req.signal?.aborted ? 'aborted' : finish };
    },
  };
}

function authHeaders(cfg: ProviderConfig): Record<string, string> {
  // 内网自建网关常不校验 key，此时留空即可
  return cfg.apiKey ? { Authorization: `Bearer ${cfg.apiKey}` } : {};
}

function toWireMessage(m: import('../types').ChatMessage) {
  if (m.role === 'tool') {
    return { role: 'tool', content: m.content, tool_call_id: m.toolCallId };
  }
  if (m.role === 'assistant' && m.toolCalls?.length) {
    return {
      role: 'assistant',
      content: m.content || null,
      tool_calls: m.toolCalls.map((c) => ({
        id: c.id,
        type: 'function',
        function: { name: c.name, arguments: c.args },
      })),
    };
  }
  return { role: m.role, content: m.content };
}

interface WireChunk {
  choices?: Array<{
    delta?: {
      content?: string;
      tool_calls?: Array<{
        index?: number;
        id?: string;
        function?: { name?: string; arguments?: string };
      }>;
    };
    finish_reason?: string | null;
  }>;
}
