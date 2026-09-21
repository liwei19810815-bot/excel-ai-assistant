import { create } from 'zustand';
import type { ChatMessage } from '../llm/types';
import type { ToolCallRecord } from '../agent/loop';
import type { ConfirmPayload } from '../tools';

/** UI 展示用的消息条目，与发给模型的 ChatMessage 分开维护 */
export interface DisplayMessage {
  id: string;
  role: 'user' | 'assistant' | 'error';
  text: string;
  /** assistant 消息下挂的工具调用卡片 */
  tools: ToolCallRecord[];
}

export interface PendingConfirm extends ConfirmPayload {
  toolName: string;
  resolve(decision: 'allow' | 'always' | 'reject'): void;
}

interface SessionState {
  /** 发给模型的完整历史，Agent 循环直接原地追加 */
  history: ChatMessage[];
  messages: DisplayMessage[];
  running: boolean;
  pendingConfirm: PendingConfirm | null;
  abortController: AbortController | null;

  pushUser(text: string): void;
  startAssistant(): string;
  appendText(id: string, chunk: string): void;
  upsertTool(id: string, record: ToolCallRecord): void;
  pushError(text: string): void;
  setRunning(v: boolean, ctrl?: AbortController | null): void;
  setConfirm(c: PendingConfirm | null): void;
  reset(): void;
}

export const useSession = create<SessionState>((set) => ({
  history: [],
  messages: [],
  running: false,
  pendingConfirm: null,
  abortController: null,

  pushUser: (text) =>
    set((s) => {
      s.history.push({ role: 'user', content: text });
      return {
        messages: [...s.messages, { id: rid(), role: 'user' as const, text, tools: [] }],
      };
    }),

  startAssistant: () => {
    const id = rid();
    set((s) => ({
      messages: [...s.messages, { id, role: 'assistant' as const, text: '', tools: [] }],
    }));
    return id;
  },

  appendText: (id, chunk) =>
    set((s) => ({
      messages: s.messages.map((m) => (m.id === id ? { ...m, text: m.text + chunk } : m)),
    })),

  upsertTool: (id, record) =>
    set((s) => ({
      messages: s.messages.map((m) => {
        if (m.id !== id) return m;
        const idx = m.tools.findIndex((t) => t.id === record.id);
        const tools =
          idx === -1
            ? [...m.tools, record]
            : m.tools.map((t, i) => (i === idx ? record : t));
        return { ...m, tools };
      }),
    })),

  pushError: (text) =>
    set((s) => ({
      messages: [...s.messages, { id: rid(), role: 'error' as const, text, tools: [] }],
    })),

  setRunning: (running, ctrl) =>
    set((s) => ({ running, abortController: ctrl === undefined ? s.abortController : ctrl })),

  setConfirm: (pendingConfirm) => set({ pendingConfirm }),

  reset: () => set({ history: [], messages: [], running: false, pendingConfirm: null }),
}));

function rid(): string {
  return `m_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}
