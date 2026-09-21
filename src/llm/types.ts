/**
 * 各家大模型的统一内部接口。
 * provider 适配层负责把这里的类型翻译成各家的请求/响应格式，
 * 上层（agent/loop.ts）只认这套类型，因此新增厂商不影响 Agent 逻辑。
 */

export type Role = 'system' | 'user' | 'assistant' | 'tool';

export interface ToolCall {
  id: string;
  name: string;
  /** 原始 JSON 字符串。流式过程中可能是不完整的片段，累积完成后再 parse。 */
  args: string;
}

export interface ChatMessage {
  role: Role;
  content: string;
  /** role === 'assistant' 时可能携带 */
  toolCalls?: ToolCall[];
  /** role === 'tool' 时必填，对应 ToolCall.id */
  toolCallId?: string;
}

/** 供模型选择的工具声明，schema 为标准 JSON Schema */
export interface ToolDef {
  name: string;
  description: string;
  schema: Record<string, unknown>;
}

/** 流式增量事件 */
export type Delta =
  | { type: 'text'; text: string }
  | { type: 'tool_call_start'; id: string; name: string }
  | { type: 'tool_call_args'; id: string; argsChunk: string }
  | { type: 'done'; reason: 'stop' | 'tool_calls' | 'length' | 'aborted' }
  | { type: 'error'; message: string };

export interface ChatRequest {
  messages: ChatMessage[];
  tools: ToolDef[];
  model: string;
  temperature?: number;
  maxTokens?: number;
  signal?: AbortSignal;
}

export interface Provider {
  readonly id: string;
  chat(req: ChatRequest): AsyncIterable<Delta>;
  /** 拉取可用模型列表，用于设置页下拉。不支持则返回 [] */
  listModels(): Promise<string[]>;
}

export interface ProviderConfig {
  kind: 'openai-compatible' | 'anthropic';
  baseUrl: string;
  apiKey: string;
  model: string;
  temperature: number;
  maxTokens: number;
}
