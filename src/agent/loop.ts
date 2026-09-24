import type { ChatMessage, Provider, ToolCall } from '../llm/types';
import { get as getTool, toToolDefs, truncate, type ConfirmPayload } from '../tools';
import { buildBlueprint, renderBlueprint } from '../excel/blueprint';
import { buildPptBlueprint, renderPptBlueprint } from '../powerpoint/blueprint';
import { getHost } from '../store/host';
import { buildSystemPrompt } from './systemPrompt';

/** 单轮对话内最多的「模型→工具→模型」往返次数，防止死循环 */
const MAX_TURNS = 25;

export interface ToolCallRecord {
  id: string;
  name: string;
  summary: string;
  status: 'running' | 'done' | 'error' | 'rejected';
  detail?: string;
  checkpointId?: string;
}

/** Agent 运行期向 UI 推送的事件 */
export type AgentEvent =
  | { type: 'assistant_text'; text: string }
  | { type: 'tool_start'; record: ToolCallRecord }
  | { type: 'tool_end'; record: ToolCallRecord }
  | { type: 'turn_end' }
  | { type: 'error'; message: string };

export interface RunOptions {
  provider: Provider;
  model: string;
  temperature: number;
  maxTokens: number;
  systemAddition?: string;
  /** sidecar 未运行等情况下需要摘除的工具 */
  disabledTools?: ReadonlySet<string>;
  /** 用户已勾选「本会话始终允许」的工具 */
  alwaysAllow: Set<string>;
  confirm(payload: ConfirmPayload & { toolName: string }): Promise<'allow' | 'always' | 'reject'>;
  emit(event: AgentEvent): void;
  signal: AbortSignal;
}

/**
 * Agent 主循环。
 * history 会被原地追加（assistant / tool 消息），调用方持有同一个数组即可保留会话。
 */
export async function runAgent(history: ChatMessage[], opts: RunOptions): Promise<void> {
  const tools = toToolDefs(opts.disabledTools);

  for (let turn = 0; turn < MAX_TURNS; turn++) {
    if (opts.signal.aborted) return;

    // 每轮重新注入蓝图 —— 上一轮的工具可能改了结构，旧快照会误导模型
    const contextBlock = await safeContextBlock();

    const messages: ChatMessage[] = [
      { role: 'system', content: buildSystemPrompt(opts.systemAddition) },
      ...history,
      { role: 'system', content: contextBlock },
    ];

    let text = '';
    const calls = new Map<string, ToolCall>();
    const order: string[] = [];

    try {
      for await (const delta of opts.provider.chat({
        messages,
        tools,
        model: opts.model,
        temperature: opts.temperature,
        maxTokens: opts.maxTokens,
        signal: opts.signal,
      })) {
        switch (delta.type) {
          case 'text':
            text += delta.text;
            opts.emit({ type: 'assistant_text', text: delta.text });
            break;
          case 'tool_call_start':
            calls.set(delta.id, { id: delta.id, name: delta.name, args: '' });
            order.push(delta.id);
            break;
          case 'tool_call_args': {
            const c = calls.get(delta.id);
            if (c) c.args += delta.argsChunk;
            break;
          }
          case 'error':
            opts.emit({ type: 'error', message: delta.message });
            return;
          case 'done':
            if (delta.reason === 'aborted') return;
            break;
        }
      }
    } catch (e) {
      if (opts.signal.aborted) return;
      opts.emit({ type: 'error', message: describeError(e) });
      return;
    }

    const toolCalls = order.map((id) => calls.get(id)!).filter(Boolean);

    history.push({
      role: 'assistant',
      content: text,
      toolCalls: toolCalls.length ? toolCalls : undefined,
    });

    // 模型没有要调工具 → 本轮结束
    if (!toolCalls.length) {
      opts.emit({ type: 'turn_end' });
      return;
    }

    for (const call of toolCalls) {
      if (opts.signal.aborted) return;
      await executeCall(call, history, opts);
    }
  }

  opts.emit({
    type: 'error',
    message: `已达到单轮 ${MAX_TURNS} 次工具调用上限，已停止。请拆分任务后重试。`,
  });
}

async function executeCall(
  call: ToolCall,
  history: ChatMessage[],
  opts: RunOptions,
): Promise<void> {
  const tool = getTool(call.name);

  if (!tool) {
    pushToolResult(history, call.id, `错误：不存在名为 ${call.name} 的工具。`);
    return;
  }

  // 模型可能生成不完整或带包裹的 JSON
  let args: unknown;
  try {
    args = JSON.parse(call.args || '{}');
  } catch {
    pushToolResult(
      history,
      call.id,
      `错误：参数不是合法 JSON。收到的内容：${call.args.slice(0, 200)}`,
    );
    return;
  }

  const parsed = tool.schema.safeParse(args);
  if (!parsed.success) {
    pushToolResult(
      history,
      call.id,
      `错误：参数校验失败。${parsed.error.issues
        .map((i) => `${i.path.join('.') || '(根)'}: ${i.message}`)
        .join('；')}`,
    );
    return;
  }

  const summary = safeSummarize(tool, parsed.data);
  const record: ToolCallRecord = {
    id: call.id,
    name: call.name,
    summary,
    status: 'running',
  };
  opts.emit({ type: 'tool_start', record });

  // 结构性改动默认强制确认，除非用户本会话勾选了「始终允许」
  if (tool.policy === 'mutate:structure' && !opts.alwaysAllow.has(tool.name)) {
    const decision = await opts.confirm({
      toolName: tool.name,
      title: summary,
      detail: `AI 想要执行「${tool.name}」，这是一个可能改变工作簿结构的操作。`,
    });
    if (decision === 'always') opts.alwaysAllow.add(tool.name);
    if (decision === 'reject') {
      const done = { ...record, status: 'rejected' as const, detail: '用户拒绝了该操作' };
      opts.emit({ type: 'tool_end', record: done });
      pushToolResult(history, call.id, '用户拒绝执行该操作。请换一种方式，或先与用户确认意图。');
      return;
    }
  }

  try {
    const result = await tool.run(parsed.data, {
      signal: opts.signal,
      // 工具内部的二次确认（如 run_script 的代码预览）走同一通道
      confirm: async (p) => {
        if (opts.alwaysAllow.has(`${tool.name}:inner`)) return true;
        const d = await opts.confirm({ ...p, toolName: tool.name });
        if (d === 'always') opts.alwaysAllow.add(`${tool.name}:inner`);
        return d !== 'reject';
      },
    });

    const done: ToolCallRecord = {
      ...record,
      status: 'done',
      detail: result.text,
      checkpointId: result.checkpointId,
    };
    opts.emit({ type: 'tool_end', record: done });
    pushToolResult(history, call.id, result.text);
  } catch (e) {
    const message = describeError(e);
    opts.emit({ type: 'tool_end', record: { ...record, status: 'error', detail: message } });
    // 错误回灌给模型，让它有机会自我修正而不是直接中断
    pushToolResult(history, call.id, `执行失败：${message}`);
  }
}

function pushToolResult(history: ChatMessage[], toolCallId: string, text: string): void {
  history.push({ role: 'tool', toolCallId, content: truncate(text) });
}

function safeSummarize(tool: { summarize: (a: never) => string }, args: unknown): string {
  try {
    return tool.summarize(args as never);
  } catch {
    return '执行操作';
  }
}

/**
 * 蓝图取不到时不应中断整轮对话，降级为提示即可。
 *
 * 按探测到的宿主选对应的蓝图——host=unknown（比如浏览器里裸调、
 * 或者 PowerPoint/Excel 都不是的宿主）时按 Excel 的老行为兜底，
 * 不是因为 Excel 更"默认"，是因为这段代码历史上只服务过 Excel，
 * 贸然改成"不认识就不读"会让所有现有部署在升级后突然看不到蓝图。
 */
async function safeContextBlock(): Promise<string> {
  const host = getHost();
  if (host === 'powerpoint') {
    try {
      return renderPptBlueprint(await buildPptBlueprint());
    } catch (e) {
      return `（无法读取演示文稿结构：${describeError(e)}。需要时请主动调用工具读取。）`;
    }
  }
  try {
    return renderBlueprint(await buildBlueprint());
  } catch (e) {
    return `（无法读取工作簿结构：${describeError(e)}。需要时请主动调用工具读取。）`;
  }
}

export function describeError(e: unknown): string {
  if (e instanceof Error) {
    // Office.js 的错误对象带 code/debugInfo，展开后更可诊断
    const info = (e as Error & { debugInfo?: { errorLocation?: string } }).debugInfo;
    const where = info?.errorLocation ? `（位置：${info.errorLocation}）` : '';
    return `${e.message}${where}`;
  }
  return String(e);
}
