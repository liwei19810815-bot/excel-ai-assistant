import { useEffect, useRef, useState } from 'react';
import { useSession } from '../../store/session';
import { useSettings, createProvider } from '../../store/settings';
import { runAgent, describeError } from '../../agent/loop';
import { ToolCallCard } from './ToolCallCard';
import { ConfirmDialog } from './ConfirmDialog';
import { SettingsPage } from './SettingsPage';
import { onUpdateAvailable, startUpdatePolling } from '../version';

const EXAMPLES = [
  '在当前表造一张 12 个月的销售数据表',
  '把标题行加粗、底色深蓝、白字',
  '根据这张表做一个柱状图',
  '这块选中的数据有什么问题？',
];

export function ChatPane() {
  const session = useSession();
  const settings = useSettings();
  const [input, setInput] = useState('');
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
        <span className="text-sm font-semibold text-neutral-800">Excel AI 助手</span>
        <div className="flex items-center gap-2 text-xs">
          {session.messages.length > 0 && (
            <button onClick={session.reset} className="text-neutral-500 hover:text-neutral-800">
              新对话
            </button>
          )}
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

      <div ref={scrollRef} className="flex-1 space-y-3 overflow-auto p-3">
        {session.messages.length === 0 && (
          <div className="space-y-3 pt-6 text-center">
            <p className="text-xs text-neutral-500">
              {configured
                ? '用自然语言描述你想做的事，AI 会直接操作当前工作簿。'
                : '还没有配置模型，先去设置页填写接口地址和模型名称。'}
            </p>
            {configured && (
              <div className="space-y-1.5">
                {EXAMPLES.map((ex) => (
                  <button
                    key={ex}
                    onClick={() => setInput(ex)}
                    className="block w-full rounded border border-neutral-200 bg-white px-2.5 py-1.5 text-left text-xs text-neutral-600 hover:border-sky-300 hover:text-sky-700"
                  >
                    {ex}
                  </button>
                ))}
              </div>
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
            插件的改动无法用 Ctrl+Z 撤销，请用操作卡片上的「撤销」
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
