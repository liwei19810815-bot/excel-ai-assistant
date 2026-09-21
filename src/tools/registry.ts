import type { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import type { ToolDef } from '../llm/types';

/**
 * 工具执行策略 —— 决定是否建快照、是否需要用户确认。
 *  read            只读，直接执行
 *  mutate:content  改值/公式/格式 → 自动快照，可撤销
 *  mutate:structure 增删行列/工作表等破坏性操作 → 快照 + 默认强制确认
 */
export type ToolPolicy = 'read' | 'mutate:content' | 'mutate:structure';

export interface ToolContext {
  /** 工具内部可请求额外确认（如 run_script 的代码预览） */
  confirm(payload: ConfirmPayload): Promise<boolean>;
  signal?: AbortSignal;
}

export interface ConfirmPayload {
  title: string;
  detail: string;
  /** 代码类确认额外带上源码，UI 用等宽字体渲染 */
  code?: string;
}

export interface ToolResult {
  /** 回灌给模型的文本 */
  text: string;
  /** 建立的快照 id，UI 据此渲染「撤销」按钮 */
  checkpointId?: string;
}

export interface Tool<S extends z.ZodTypeAny = z.ZodTypeAny> {
  name: string;
  description: string;
  policy: ToolPolicy;
  schema: S;
  /** 给用户看的一句话摘要，渲染在工具调用卡片标题上 */
  summarize(args: z.infer<S>): string;
  run(args: z.infer<S>, ctx: ToolContext): Promise<ToolResult>;
}

const tools = new Map<string, Tool>();

/** 泛型参数让 run/summarize 的 args 能从 schema 推导出具体类型 */
export function register<S extends z.ZodTypeAny>(tool: Tool<S>): void {
  if (tools.has(tool.name)) throw new Error(`工具重复注册：${tool.name}`);
  tools.set(tool.name, tool as unknown as Tool);
}

export function get(name: string): Tool | undefined {
  return tools.get(name);
}

export function all(): Tool[] {
  return [...tools.values()];
}

/** 导出为模型可见的工具声明。disabled 里的名字会被摘除（如 sidecar 未运行时）。 */
export function toToolDefs(disabled: ReadonlySet<string> = new Set()): ToolDef[] {
  return all()
    .filter((t) => !disabled.has(t.name))
    .map((t) => ({
      name: t.name,
      description: t.description,
      schema: zodToJsonSchema(t.schema, {
        $refStrategy: 'none', // 多数厂商不解析 $ref，内联展开更稳
        target: 'openApi3',
      }) as Record<string, unknown>,
    }));
}

/** 单条工具结果回灌上限，防止一次 read_range 撑爆上下文 */
export const MAX_RESULT_CHARS = 8000;

export function truncate(text: string): string {
  if (text.length <= MAX_RESULT_CHARS) return text;
  return (
    text.slice(0, MAX_RESULT_CHARS) +
    `\n\n…（结果过长已截断，原始长度 ${text.length} 字符。如需完整内容请分块读取。）`
  );
}
