/**
 * 最小 SSE 解析器：把 fetch 的字节流切成一行行 `data:` 载荷。
 * OpenAI 兼容端点与 Anthropic 都用 SSE，故两个 provider 共用。
 */
export async function* readSSE(
  res: Response,
  signal?: AbortSignal,
): AsyncGenerator<string> {
  if (!res.body) throw new Error('响应没有 body，无法流式读取');

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';

  try {
    while (true) {
      if (signal?.aborted) return;
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });

      // SSE 事件以空行分隔；按行处理即可覆盖两家格式
      let idx: number;
      while ((idx = buffer.indexOf('\n')) !== -1) {
        const line = buffer.slice(0, idx).trim();
        buffer = buffer.slice(idx + 1);
        if (line.startsWith('data:')) {
          yield line.slice(5).trim();
        }
      }
    }
  } finally {
    reader.releaseLock();
  }
}

/** 把非 2xx 响应转成带上下文的错误，便于 UI 显示可读信息 */
export async function throwIfNotOk(res: Response, label: string): Promise<void> {
  if (res.ok) return;
  let detail = '';
  try {
    detail = (await res.text()).slice(0, 500);
  } catch {
    /* 读不出正文就算了 */
  }
  throw new Error(`${label} 请求失败（HTTP ${res.status}）${detail ? `：${detail}` : ''}`);
}
