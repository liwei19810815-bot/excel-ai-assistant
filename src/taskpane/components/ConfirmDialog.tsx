import type { PendingConfirm } from '../../store/session';

/**
 * 破坏性操作与 run_script 的确认弹窗。
 * 对 run_script 而言这是唯一真正的安全边界，所以代码必须完整展示、不折叠。
 */
export function ConfirmDialog({ pending }: { pending: PendingConfirm }) {
  return (
    <div className="absolute inset-0 z-50 flex items-end bg-black/30 backdrop-blur-[1px]">
      <div className="max-h-[85%] w-full overflow-auto rounded-t-lg border-t border-neutral-300 bg-white p-4 shadow-2xl">
        <div className="mb-1 text-sm font-semibold text-neutral-900">{pending.title}</div>
        <div className="mb-3 whitespace-pre-wrap text-xs leading-relaxed text-neutral-600">
          {pending.detail}
        </div>

        {pending.code && (
          <pre className="mb-3 max-h-64 overflow-auto rounded border border-neutral-200 bg-neutral-50 p-2.5 font-mono text-[11px] leading-relaxed text-neutral-800">
            {pending.code}
          </pre>
        )}

        <div className="flex gap-2">
          <button
            onClick={() => pending.resolve('allow')}
            className="flex-1 rounded bg-sky-600 px-3 py-2 text-xs font-medium text-white hover:bg-sky-700"
          >
            允许
          </button>
          <button
            onClick={() => pending.resolve('always')}
            className="rounded border border-neutral-300 px-3 py-2 text-xs text-neutral-700 hover:bg-neutral-50"
            title="本次会话内不再询问该工具"
          >
            始终允许
          </button>
          <button
            onClick={() => pending.resolve('reject')}
            className="rounded border border-neutral-300 px-3 py-2 text-xs text-neutral-700 hover:bg-neutral-50"
          >
            拒绝
          </button>
        </div>
      </div>
    </div>
  );
}
