import { useEffect, useRef, useState } from 'react';
import { useSession } from '../../store/session';
import { useSettings, createProvider } from '../../store/settings';
import { runAgent, describeError } from '../../agent/loop';
import { ToolCallCard } from './ToolCallCard';
import { ConfirmDialog } from './ConfirmDialog';
import { SettingsPage } from './SettingsPage';
import { QuickActions } from './QuickActions';
import { onUpdateAvailable, startUpdatePolling } from '../version';
import { getSidecarStatus } from '../../store/sidecar';
import { getHost } from '../../store/host';
import { SIDECAR_TOOL_NAMES } from '../../tools/sidecar/powerQuery';
import { RUN_MACRO_TOOL_NAMES } from '../../tools/sidecar/runMacro';
import { EXCEL_TOOL_NAMES } from '../../tools/excel';
import { POWERPOINT_TOOL_NAMES } from '../../tools/powerpoint';
import { WORD_TOOL_NAMES } from '../../tools/word';

export function ChatPane() {
  const session = useSession();
  const settings = useSettings();
  const [input, setInput] = useState('');
  const [showQuick, setShowQuick] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [updateVersion, setUpdateVersion] = useState<string | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    startUpdatePolling();
    return onUpdateAvailable(setUpdateVersion);
  }, []);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: 'smooth' });
  }, [session.messages]);

  const configured = Boolean(settings.baseUrl && settings.model);

  // 标题/说明文字按宿主换措辞——host=unknown（浏览器裸调等）时按 Excel
  // 文案兜底，理由同 agent/loop.ts 的 safeContextBlock：这段文案历史上
  // 只服务过 Excel，不认识的宿主突然不认它会让现有部署看着奇怪。
  const host = getHost();
  const title = host === 'powerpoint' ? 'PPT AI 助手' : host === 'word' ? 'Word AI 助手' : 'Excel AI 助手';
  const docWord = host === 'powerpoint' ? '当前演示文稿' : host === 'word' ? '当前文档' : '当前工作簿';

  async function send() {
    const text = input.trim();
    if (!text || session.running) return;
    if (!configured) {
      setShowSettings(true);
      return;
    }

    setInput('');
    session.pushUser(text);

    const ctrl = new AbortController();
    session.setRunning(true, ctrl);
    const msgId = session.startAssistant();

    // 关掉 run_script 时从工具列表中摘除，模型就看不到它
    const disabled = new Set<string>();
    if (!settings.enableRunScript) disabled.add('run_script');

    // sidecar 没跑就把它那几个工具摘掉，模型看不到就不会承诺做不到的事。
    // 【这是安静降级】：不弹窗、不报错，AI 其余功能完全正常。
    if (!getSidecarStatus().available) {
      for (const n of SIDECAR_TOOL_NAMES) disabled.add(n);
      for (const n of RUN_MACRO_TOOL_NAMES) disabled.add(n);
    }

    // 按宿主摘掉不匹配的那一套工具——Excel 工具调 Excel.run，PPT 工具调
    // PowerPoint.run，Word 工具调 Word.run，装错宿主直接报错，模型不该
    // 看到它调不动的工具。sidecar 那两套（Power Query / VBA 宏）也是
    // Excel 专属概念，非 Excel 宿主时一并摘掉，即使 sidecar 本身探测到
    // 可用。复用组件顶层已经算过的 host，不用再探测一次。
    // 【host=unknown 按 excel 兜底】：和标题/文案那处兜底同一个理由，
    // 浏览器裸调等场景之前一直是当 Excel 用的，不能因为加了 Word
    // 分支就悄悄改变这个默认行为。
    const effectiveHost = host === 'unknown' ? 'excel' : host;
    if (effectiveHost !== 'excel') {
      for (const n of EXCEL_TOOL_NAMES) disabled.add(n);
      for (const n of SIDECAR_TOOL_NAMES) disabled.add(n);
      for (const n of RUN_MACRO_TOOL_NAMES) disabled.add(n);
    }
    if (effectiveHost !== 'powerpoint') {
      for (const n of POWERPOINT_TOOL_NAMES) disabled.add(n);
    }
    if (effectiveHost !== 'word') {
      for (const n of WORD_TOOL_NAMES) disabled.add(n);
    }

    try {
      await runAgent(session.history, {
        provider: createProvider(settings),
        model: settings.model,
        temperature: settings.temperature,
        maxTokens: settings.maxTokens,
        systemAddition: settings.systemAddition,
        disabledTools: disabled,
        alwaysAllow: new Set(),
        signal: ctrl.signal,
        confirm: (payload) =>
          new Promise((resolve) => {
            session.setConfirm({
              ...payload,
              resolve: (d) => {
                session.setConfirm(null);
                resolve(d);
              },
            });
          }),
        emit: (e) => {
          switch (e.type) {
            case 'assistant_text':
              session.appendText(msgId, e.text);
              break;
            case 'tool_start':
            case 'tool_end':
              session.upsertTool(msgId, e.record);
              break;
            case 'error':
              session.pushError(e.message);
              break;
          }
        },
      });
    } catch (e) {
      session.pushError(describeError(e));
    } finally {
      session.setRunning(false, null);
    }
  }

  /**
   * 还没问到网关的配置。
   *
   * 【这段不能放行】。原先是界面先挂出来、provision 在后台异步跑，
   * 默认按"可用"处理——那就留出了一个窗口：开关已经在服务端关掉了，
   * 用户却能在这几百毫秒里正常发消息。治理开关被绕过。
   *
   * 所以 pending 期间一律按不可用处理。等待最长就是 provision 的
   * 3 秒超时，网关不可达时也会在 3 秒内落到 byok/可用。
   */
  if (settings.provision.status === 'pending') {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-2 bg-white p-6 text-center">
        <div className="text-sm text-neutral-700">正在获取配置…</div>
        <div className="text-xs text-neutral-500">首次打开需要向服务器确认一次</div>
      </div>
    );
  }

  /**
   * 服务端把 AI 功能关掉了（visibility 0 或 2）。
   *
   * 【这里是真正的强制点】。功能区按钮置灰只是"看起来不能用"——
   * 而且没有共享运行时的话，按钮要等加载项跑起来之后才会变灰，
   * 用户第一次点开之前它是正常的。所以拦截必须落在窗格里：
   * 无论他怎么点进来，看到的都是这一页，聊天界面根本不渲染。
   *
   * 0 和 2 在这里的行为一样。区别在安装侧：0 会让安装程序【不注册】
   * 这个加载项，按钮从一开始就不存在；2 保留按钮只是置灰。
   */
  if (settings.provision.visibility !== 1) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-3 bg-white p-6 text-center">
        <div className="text-sm font-semibold text-neutral-800">AI 助手已停用</div>
        <div className="text-xs leading-relaxed text-neutral-600">
          此功能已被{settings.provision.managedBy || 'IT'}关闭。
          <br />
          如需使用，请联系 IT 运维。
        </div>
        <div className="mt-2 text-[11px] text-neutral-400">
          工具箱的其它功能不受影响，可以正常使用。
        </div>
      </div>
    );
  }

  if (showSettings) return <SettingsPage onClose={() => setShowSettings(false)} />;

  return (
    <div className="relative flex h-full flex-col bg-neutral-50">
      <header className="flex items-center justify-between border-b border-neutral-200 bg-white px-3 py-2">
        <span className="text-sm font-semibold text-neutral-800">{title}</span>
        <div className="flex items-center gap-2 text-xs">
          {session.messages.length > 0 && (
            <button onClick={session.reset} className="text-neutral-500 hover:text-neutral-800">
              新对话
            </button>
          )}
          <button
            onClick={() => setShowQuick((v) => !v)}
            className={showQuick ? 'text-sky-700 underline' : 'text-sky-600 hover:underline'}
          >
            常用操作
          </button>
          <button onClick={() => setShowSettings(true)} className="text-sky-600 hover:underline">
            设置
          </button>
        </div>
      </header>

      {updateVersion && (
        <button
          onClick={() => location.reload()}
          className="border-b border-amber-200 bg-amber-50 px-3 py-1.5 text-left text-xs text-amber-800 hover:bg-amber-100"
        >
          有新版本 {updateVersion} 可用，点击刷新
        </button>
      )}

      {showQuick && session.messages.length > 0 && (
        <div className="max-h-64 overflow-auto border-b border-neutral-200 bg-white p-3">
          <QuickActions
            onPick={(prompt) => {
              setInput(prompt);
              setShowQuick(false);
            }}
          />
        </div>
      )}

      <div ref={scrollRef} className="flex-1 space-y-3 overflow-auto p-3">
        {session.messages.length === 0 && (
          <div className="space-y-3 pt-6 text-center">
            <p className="text-xs text-neutral-500">
              {configured
                ? `用自然语言描述你想做的事，AI 会直接操作${docWord}。`
                : '还没有配置模型，先去设置页填写接口地址和模型名称。'}
            </p>
            {configured && (
              <QuickActions
                onPick={(prompt) => {
                  setInput(prompt);
                  setShowQuick(false);
                }}
              />
            )}
            {!configured && (
              <button
                onClick={() => setShowSettings(true)}
                className="rounded bg-sky-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-sky-700"
              >
                去设置
              </button>
            )}
          </div>
        )}

        {session.messages.map((m) => (
          <div key={m.id} className="space-y-1.5">
            {m.role === 'user' && (
              <div className="ml-6 rounded-lg bg-sky-600 px-3 py-2 text-xs leading-relaxed text-white">
                {m.text}
              </div>
            )}

            {m.role === 'error' && (
              <div className="rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs leading-relaxed text-red-800">
                {m.text}
              </div>
            )}

            {m.role === 'assistant' && (
              <>
                {m.text && (
                  <div className="whitespace-pre-wrap rounded-lg border border-neutral-200 bg-white px-3 py-2 text-xs leading-relaxed text-neutral-800">
                    {m.text}
                  </div>
                )}
                {m.tools.map((t) => (
                  <ToolCallCard key={t.id} record={t} />
                ))}
              </>
            )}
          </div>
        ))}
      </div>

      <div className="border-t border-neutral-200 bg-white p-2">
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              void send();
            }
          }}
          rows={2}
          placeholder={configured ? '描述你想做的事…（Enter 发送，Shift+Enter 换行）' : '请先配置模型'}
          className="w-full resize-none rounded border border-neutral-300 px-2 py-1.5 text-xs outline-none focus:border-sky-500"
        />
        <div className="mt-1.5 flex items-center justify-between">
          <span className="text-[11px] text-neutral-400">
            {host === 'powerpoint' || host === 'word'
              ? '插件的改动大概率也无法用 Ctrl+Z 撤销，改结构前 AI 会先问你'
              : '插件的改动无法用 Ctrl+Z 撤销，请用操作卡片上的「撤销」'}
          </span>
          {session.running ? (
            <button
              onClick={() => session.abortController?.abort()}
              className="rounded border border-neutral-300 px-3 py-1 text-xs hover:bg-neutral-50"
            >
              停止
            </button>
          ) : (
            <button
              onClick={() => void send()}
              disabled={!input.trim()}
              className="rounded bg-sky-600 px-3 py-1 text-xs font-medium text-white hover:bg-sky-700 disabled:opacity-40"
            >
              发送
            </button>
          )}
        </div>
      </div>

      {session.pendingConfirm && <ConfirmDialog pending={session.pendingConfirm} />}
    </div>
  );
}
