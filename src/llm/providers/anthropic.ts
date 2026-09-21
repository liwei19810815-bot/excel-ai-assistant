import type { ChatMessage, ChatRequest, Delta, Provider, ProviderConfig } from '../types';
import { readSSE, throwIfNotOk } from '../sse';

/**
 * Anthropic Messages API 适配。
 * 与 OpenAI 协议的三处关键差异：
 *  1. system 是顶层参数，不在 messages 里
 *  2. tool_result 以 user 消息的 content block 形式回灌
 *  3. 工具参数走 input_json_delta 流式分片
 */
export function createAnthropicProvider(cfg: ProviderConfig): Provider {
  const base = cfg.baseUrl.replace(/\/+$/, '');

  return {
    id: 'anthropic',

    async listModels() {
      try {
        const res = await fetch(`${base}/models`, { headers: headers(cfg) });
        if (!res.ok) return [];
        const json = (await res.json()) as { data?: Array<{ id: string }> };
        return (json.data ?? []).map((m) => m.id).sort();
      } catch {
        return [];
      }
    },

    async *chat(req: ChatRequest): AsyncIterable<Delta> {
      const system = req.messages
        .filter((m) => m.role === 'system')
        .map((m) => m.content)
        .join('\n\n');

      const res = await fetch(`${base}/messages`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...headers(cfg) },
        signal: req.signal,
        body: JSON.stringify({
          model: req.model,
          system: system || undefined,
          messages: toWireMessages(req.messages.filter((m) => m.role !== 'system')),
          tools: req.tools.length
            ? req.tools.map((t) => ({
                name: t.name,
                description: t.description,
                input_schema: t.schema,
              }))
            : undefined,
          temperature: req.temperature,
          max_tokens: req.maxTokens ?? 4096,
          stream: true,
        }),
      });

      await throwIfNotOk(res, 'Anthropic 接口');

      // content_block_start 给出 tool id/name，后续 input_json_delta 按 block index 累积
      const blockIds = new Map<number, string>();
      let finish: 'stop' | 'tool_calls' | 'length' = 'stop';

      for await (const payload of readSSE(res, req.signal)) {
        let evt: WireEvent;
        try {
          evt = JSON.parse(payload);
        } catch {
          continue;
        }

        switch (evt.type) {
          case 'content_block_start':
            if (evt.content_block?.type === 'tool_use' && evt.index !== undefined) {
              const id = evt.content_block.id ?? `call_${evt.index}`;
              blockIds.set(evt.index, id);
              yield { type: 'tool_call_start', id, name: evt.content_block.name ?? '' };
            }
            break;

          case 'content_block_delta':
            if (evt.delta?.type === 'text_delta' && evt.delta.text) {
              yield { type: 'text', text: evt.delta.text };
            } else if (evt.delta?.type === 'input_json_delta' && evt.index !== undefined) {
              const id = blockIds.get(evt.index);
              if (id) {
                yield { type: 'tool_call_args', id, argsChunk: evt.delta.partial_json ?? '' };
              }
            }
            break;

          case 'message_delta':
            if (evt.delta?.stop_reason === 'tool_use') finish = 'tool_calls';
            else if (evt.delta?.stop_reason === 'max_tokens') finish = 'length';
            break;

          case 'error':
            yield { type: 'error', message: evt.error?.message ?? 'Anthropic 返回未知错误' };
            return;
        }
      }

      yield { type: 'done', reason: req.signal?.aborted ? 'aborted' : finish };
    },
  };
}

function headers(cfg: ProviderConfig): Record<string, string> {
  return {
    'anthropic-version': '2023-06-01',
    // 浏览器环境（任务窗格）直连必须显式开启，否则被 SDK 侧 CORS 策略拒绝
    'anthropic-dangerous-direct-browser-access': 'true',
    ...(cfg.apiKey ? { 'x-api-key': cfg.apiKey } : {}),
  };
}

/** 把内部消息翻成 Anthropic 的 content block 结构 */
function toWireMessages(messages: ChatMessage[]) {
  return messages.map((m) => {
    if (m.role === 'tool') {
      return {
        role: 'user' as const,
        content: [
          { type: 'tool_result', tool_use_id: m.toolCallId, content: m.content },
        ],
      };
    }

    if (m.role === 'assistant' && m.toolCalls?.length) {
      const blocks: unknown[] = [];
      if (m.content) blocks.push({ type: 'text', text: m.content });
      for (const c of m.toolCalls) {
        blocks.push({
          type: 'tool_use',
          id: c.id,
          name: c.name,
          input: safeParse(c.args),
        });
      }
      return { role: 'assistant' as const, content: blocks };
    }

    return { role: m.role as 'user' | 'assistant', content: m.content };
  });
}

function safeParse(s: string): unknown {
  try {
    return JSON.parse(s || '{}');
  } catch {
    return {};
  }
}

interface WireEvent {
  type: string;
  index?: number;
  content_block?: { type?: string; id?: string; name?: string };
  delta?: {
    type?: string;
    text?: string;
    partial_json?: string;
    stop_reason?: string;
  };
  error?: { message?: string };
}
