import { useState } from 'react';
import type { ToolCallRecord } from '../../agent/loop';
import { restore } from '../../excel/checkpoint';

const STATUS: Record<ToolCallRecord['status'], { icon: string; cls: string; label: string }> = {
  running: { icon: '⋯', cls: 'text-sky-600 bg-sky-50 border-sky-200', label: '执行中' },
  done: { icon: '✓', cls: 'text-emerald-700 bg-emerald-50 border-emerald-200', label: '完成' },
  error: { icon: '!', cls: 'text-red-700 bg-red-50 border-red-200', label: '失败' },
  rejected: { icon: '✕', cls: 'text-neutral-600 bg-neutral-100 border-neutral-300', label: '已拒绝' },
};

export function ToolCallCard({ record }: { record: ToolCallRecord }) {
  const [expanded, setExpanded] = useState(false);
  const [undone, setUndone] = useState(false);
  const [undoError, setUndoError] = useState<string | null>(null);
  const s = STATUS[record.status];

  async function handleUndo() {
    if (!record.checkpointId) return;
    try {
      await restore(record.checkpointId);
      setUndone(true);
      setUndoError(null);
    } catch (e) {
      setUndoError(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div className={`rounded-md border px-2.5 py-1.5 text-xs ${s.cls}`}>
      <div className="flex items-center gap-2">
        <span className="font-mono font-bold">{s.icon}</span>
        <span className="flex-1 truncate font-medium">{record.summary}</span>

        {record.checkpointId && record.status === 'done' && !undone && (
          <button
            onClick={handleUndo}
            className="shrink-0 rounded border border-current px-1.5 py-0.5 opacity-70 hover:opacity-100"
          >
            撤销
          </button>
        )}
        {undone && <span className="shrink-0 opacity-60">已撤销</span>}

        {record.detail && (
          <button
            onClick={() => setExpanded((v) => !v)}
            className="shrink-0 opacity-60 hover:opacity-100"
          >
            {expanded ? '收起' : '详情'}
          </button>
        )}
      </div>

      {undoError && <div className="mt-1 text-red-700">撤销失败：{undoError}</div>}

      {expanded && record.detail && (
        <pre className="mt-1.5 max-h-60 overflow-auto whitespace-pre-wrap break-all rounded bg-white/60 p-2 font-mono text-[11px] leading-relaxed">
          {record.detail}
        </pre>
      )}
    </div>
  );
}
