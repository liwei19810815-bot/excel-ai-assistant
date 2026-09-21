import { describe, it, expect, vi, afterEach } from 'vitest';
import { createOpenAICompatibleProvider } from './openaiCompatible';
import type { Delta } from '../types';

/** 把若干 SSE 行拼成一个可读的 Response，用于回放录制的流式响应 */
function sseResponse(lines: string[]): Response {
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      const enc = new TextEncoder();
      for (const l of lines) controller.enqueue(enc.encode(`data: ${l}\n\n`));
      controller.close();
    },
  });
  return new Response(body, { status: 200 });
}

const cfg = {
  kind: 'openai-compatible' as const,
  baseUrl: 'https://gateway.test/v1',
  apiKey: 'k',
  model: 'm',
  temperature: 0,
  maxTokens: 100,
};

async function collect(stream: AsyncIterable<Delta>): Promise<Delta[]> {
  const out: Delta[] = [];
  for await (const d of stream) out.push(d);
  return out;
}

afterEach(() => vi.unstubAllGlobals());

describe('OpenAI 兼容 provider', () => {
  it('解析流式文本增量', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        sseResponse([
          JSON.stringify({ choices: [{ delta: { content: '你好' } }] }),
          JSON.stringify({ choices: [{ delta: { content: '世界' } }] }),
          JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] }),
          '[DONE]',
        ]),
      ),
    );

    const deltas = await collect(
      createOpenAICompatibleProvider(cfg).chat({ messages: [], tools: [], model: 'm' }),
    );

    expect(deltas.filter((d) => d.type === 'text').map((d) => (d as never)['text'])).toEqual([
      '你好',
      '世界',
    ]);
    expect(deltas.at(-1)).toEqual({ type: 'done', reason: 'stop' });
  });

  it('按 index 累积分片的工具调用参数', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        sseResponse([
          JSON.stringify({
            choices: [
              {
                delta: {
                  tool_calls: [
                    { index: 0, id: 'call_1', function: { name: 'write_cells', arguments: '{"ad' } },
                  ],
                },
              },
            ],
          }),
          JSON.stringify({
            choices: [
              { delta: { tool_calls: [{ index: 0, id: 'call_1', function: { arguments: 'dress":"A1"}' } }] } },
            ],
          }),
          JSON.stringify({ choices: [{ delta: {}, finish_reason: 'tool_calls' }] }),
          '[DONE]',
        ]),
      ),
    );

    const deltas = await collect(
      createOpenAICompatibleProvider(cfg).chat({ messages: [], tools: [], model: 'm' }),
    );

    const start = deltas.find((d) => d.type === 'tool_call_start');
    expect(start).toMatchObject({ id: 'call_1', name: 'write_cells' });

    const args = deltas
      .filter((d) => d.type === 'tool_call_args')
      .map((d) => (d as never)['argsChunk'])
      .join('');
    expect(JSON.parse(args)).toEqual({ address: 'A1' });

    expect(deltas.at(-1)).toEqual({ type: 'done', reason: 'tool_calls' });
  });

  it('忽略非 JSON 的心跳行而不是崩溃', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        sseResponse([
          ': keep-alive',
          JSON.stringify({ choices: [{ delta: { content: 'ok' } }] }),
          '[DONE]',
        ]),
      ),
    );

    const deltas = await collect(
      createOpenAICompatibleProvider(cfg).chat({ messages: [], tools: [], model: 'm' }),
    );
    expect(deltas.some((d) => d.type === 'text')).toBe(true);
  });

  it('非 2xx 时抛出带状态码的可读错误', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response('model not found', { status: 404 })),
    );

    await expect(
      collect(createOpenAICompatibleProvider(cfg).chat({ messages: [], tools: [], model: 'm' })),
    ).rejects.toThrow(/404/);
  });

  it('apiKey 为空时不发送 Authorization —— 内网网关常不校验', async () => {
    const fetchMock = vi.fn((_url: string, _init: RequestInit) =>
      Promise.resolve(sseResponse(['[DONE]'])),
    );
    vi.stubGlobal('fetch', fetchMock);

    await collect(
      createOpenAICompatibleProvider({ ...cfg, apiKey: '' }).chat({
        messages: [],
        tools: [],
        model: 'm',
      }),
    );

    const headers = fetchMock.mock.calls[0][1].headers as Record<string, string>;
    expect(headers.Authorization).toBeUndefined();
    expect(headers['Content-Type']).toBe('application/json');
  });
});
