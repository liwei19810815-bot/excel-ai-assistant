import { useState } from 'react';
import { useSettings, PRESETS, testConnection } from '../../store/settings';
import { APP_VERSION, checkForUpdate } from '../version';

export function SettingsPage({ onClose }: { onClose: () => void }) {
  const s = useSettings();
  // 白名单命中时，连接相关的输入框一律锁住（见下面那块提示）
  const managed = s.provision.mode === 'managed';
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState<{ ok: boolean; message: string } | null>(null);

  async function handleTest() {
    setTesting(true);
    setResult(null);
    const r = await testConnection(s);
    setResult(r);
    if (r.models.length) s.set({ knownModels: r.models });
    setTesting(false);
  }

  function handleExport() {
    // 导出配置供同事导入 —— 内网场景下这是降低配置门槛最实用的一招
    const { kind, baseUrl, model, temperature, maxTokens, systemAddition } = s;
    const blob = new Blob(
      [JSON.stringify({ kind, baseUrl, model, temperature, maxTokens, systemAddition }, null, 2)],
      { type: 'application/json' },
    );
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'excel-ai-config.json';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function handleImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    // 【managed 下不许导入】。把输入框锁住却留着"导入配置"这条路，
    // 等于没锁——用户导一份文件就把 IT 下发的模型覆盖掉了。
    if (managed) {
      setResult({ ok: false, message: '当前模型由 IT 统一配置，不能导入配置覆盖。' });
      return;
    }

    file.text().then((t) => {
      try {
        s.set(JSON.parse(t));
        setResult({ ok: true, message: '配置已导入。' });
      } catch {
        setResult({ ok: false, message: '配置文件格式不正确。' });
      }
    });
  }

  return (
    <div className="flex h-full flex-col bg-white">
      <header className="flex items-center justify-between border-b border-neutral-200 px-3 py-2">
        <span className="text-sm font-semibold">设置</span>
        <button onClick={onClose} className="text-xs text-sky-600 hover:underline">
          返回对话
        </button>
      </header>

      <div className="flex-1 space-y-4 overflow-auto p-3 text-xs">
        {/*
          managed：模型由 IT 统一配置，输入框全部锁住。
          【必须显式告诉用户"这是被管控的"】——否则他改了半天发现改不动，
          或者以为自己的配置生效了、实际用的是别人下发的，都很糟。
        */}
        {managed && (
          <div className="rounded border border-sky-200 bg-sky-50 p-2 text-sky-900">
            <div className="font-medium">模型由 IT 统一配置</div>
            <div className="mt-1 leading-relaxed">
              当前账号（{s.provision.user || '未知'}）在白名单内，
              下面的连接设置由{s.provision.managedBy || 'IT'}下发，不需要也不能在这里改。
              <br />
              配置只保存在内存里，每次打开重新获取。
            </div>
          </div>
        )}

        {!managed && s.provision.reason && (
          <div className="rounded border border-neutral-200 bg-neutral-50 p-2 text-neutral-600">
            需要自行配置模型（{s.provision.reason}）。
            填好下面的接口地址和模型名称即可使用。
          </div>
        )}

        <Field label="快速选择">
          <select
            className={inputCls}
            defaultValue=""
            disabled={managed}
            onChange={(e) => e.target.value && s.applyPreset(e.target.value)}
          >
            <option value="">— 选择一个预设 —</option>
            {PRESETS.map((p) => (
              <option key={p.id} value={p.id}>
                {p.label} — {p.hint}
              </option>
            ))}
          </select>
        </Field>

        <Field label="协议">
          <select
            className={inputCls}
            value={s.kind}
            disabled={managed}
            onChange={(e) => s.set({ kind: e.target.value as typeof s.kind })}
          >
            <option value="openai-compatible">OpenAI 兼容（vLLM / Ollama / DeepSeek / 通义…）</option>
            <option value="anthropic">Anthropic Messages</option>
          </select>
        </Field>

        <Field label="接口地址" hint="需以 /v1 结尾，如 https://gateway.corp/v1">
          <input
            className={inputCls}
            value={s.baseUrl}
            disabled={managed}
            onChange={(e) => s.set({ baseUrl: e.target.value })}
            placeholder="https://模型网关地址/v1"
          />
        </Field>

        <Field label="API Key" hint="内网自建网关通常不校验，可留空">
          <input
            className={inputCls}
            type="password"
            value={s.apiKey}
            disabled={managed}
            onChange={(e) => s.set({ apiKey: e.target.value })}
            placeholder="留空表示不发送鉴权头"
          />
        </Field>

        <Field label="模型">
          {s.knownModels.length ? (
            <select
              className={inputCls}
              value={s.model}
              disabled={managed}
              onChange={(e) => s.set({ model: e.target.value })}
            >
              <option value="">— 选择模型 —</option>
              {s.knownModels.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          ) : (
            <input
              className={inputCls}
              value={s.model}
              disabled={managed}
              onChange={(e) => s.set({ model: e.target.value })}
              placeholder="测试连接后可从列表选择"
            />
          )}
        </Field>

        <div className="grid grid-cols-2 gap-2">
          <Field label="温度">
            <input
              className={inputCls}
              type="number"
              step="0.1"
              min="0"
              max="2"
              value={s.temperature}
              onChange={(e) => s.set({ temperature: Number(e.target.value) })}
            />
          </Field>
          <Field label="最大输出 tokens">
            <input
              className={inputCls}
              type="number"
              value={s.maxTokens}
              onChange={(e) => s.set({ maxTokens: Number(e.target.value) })}
            />
          </Field>
        </div>

        <Field label="追加系统提示词" hint="会拼在内置提示词之后，用于团队约定的格式规范等">
          <textarea
            className={`${inputCls} h-20 resize-none`}
            value={s.systemAddition}
            onChange={(e) => s.set({ systemAddition: e.target.value })}
            placeholder="例如：金额一律保留两位小数，用 ¥ 符号。"
          />
        </Field>

        {/*
          【IT 下发过就不许改回来】。run_script 让模型现写 Office.js 代码执行，
          虽然每次都要人工确认、作用域受限、有静态检查，但它是全套工具里
          风险最高的一个。既然允许 IT 集中关闭，就不能让用户在界面上
          一勾了事——那样"集中管控"只是个摆设。
        */}
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={s.enableRunScript}
            disabled={s.provision.runScriptPinned}
            onChange={(e) => s.set({ enableRunScript: e.target.checked })}
          />
          <span>
            启用自定义脚本工具（run_script）
            <span className="block text-[11px] text-neutral-500">
              关闭后 AI 只能使用结构化工具，更安全但能力受限
            </span>
            {s.provision.runScriptPinned && (
              <span className="block text-[11px] text-sky-700">
                该项由{s.provision.managedBy || 'IT'}统一设定，本机不能更改。
              </span>
            )}
          </span>
        </label>

        <div className="flex gap-2 pt-1">
          <button
            onClick={handleTest}
            disabled={testing}
            className="rounded bg-sky-600 px-3 py-1.5 font-medium text-white hover:bg-sky-700 disabled:opacity-50"
          >
            {testing ? '测试中…' : '测试连接'}
          </button>
          <button onClick={handleExport} className={btnCls}>
            导出配置
          </button>
          <label className={`${btnCls} cursor-pointer`}>
            导入配置
            <input type="file" accept=".json" className="hidden" onChange={handleImport} />
          </label>
        </div>

        {result && (
          <div
            className={`whitespace-pre-wrap rounded border p-2 ${
              result.ok
                ? 'border-emerald-200 bg-emerald-50 text-emerald-800'
                : 'border-red-200 bg-red-50 text-red-800'
            }`}
          >
            {result.message}
          </div>
        )}

        <div className="border-t border-neutral-200 pt-3 text-[11px] text-neutral-500">
          <div className="flex items-center justify-between">
            <span>版本 {APP_VERSION}</span>
            <button
              onClick={() => checkForUpdate(true)}
              className="text-sky-600 hover:underline"
            >
              检查更新
            </button>
          </div>
          <p className="mt-1.5 leading-relaxed">
            API Key 保存在本机浏览器存储中，不会发送到除你配置的接口地址之外的任何地方。
          </p>
        </div>
      </div>
    </div>
  );
}

const inputCls =
  'w-full rounded border border-neutral-300 px-2 py-1.5 text-xs outline-none focus:border-sky-500';
const btnCls = 'rounded border border-neutral-300 px-3 py-1.5 hover:bg-neutral-50';

function Field({
  label,
  hint,
  children,
}: {
  label: string;
  hint?: string;
  children: React.ReactNode;
}) {
  return (
    <div>
      <div className="mb-1 font-medium text-neutral-700">{label}</div>
      {children}
      {hint && <div className="mt-0.5 text-[11px] text-neutral-500">{hint}</div>}
    </div>
  );
}
