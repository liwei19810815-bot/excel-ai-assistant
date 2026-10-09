import { describe, it, expect, beforeEach, vi } from 'vitest';
import { runStartup } from './startup';
import { useSettings } from '../store/settings';
import * as sidecarMod from '../store/sidecar';
import { getSidecarStatus, setSidecarStatus } from '../store/sidecar';
import { getHost, setHost } from '../store/host';

beforeEach(() => {
  useSettings.setState({
    kind: 'openai-compatible', baseUrl: '', apiKey: '', model: '', enableRunScript: true,
    provision: { mode: 'byok', user: '', status: 'pending' },
  });
  delete (globalThis as { Office?: unknown }).Office;
});

describe('runStartup 链路', () => {
  it('完成配置、sidecar 和宿主探测后进入 ready', async () => {
    const r = await runStartup();
    expect(r.mode).toBe('byok');
    expect(r.host).toBe('unknown');
    expect(r.sidecar.available).toBe(false);
    expect(useSettings.getState().provision.status).toBe('ready');
  });
});

describe('启动路径：sidecar 探测的接线', () => {
  beforeEach(() => setSidecarStatus({ available: false, port: 0, version: '', reason: '重置' }));

  it('探到了就要把状态存下来', async () => {
    const status = { available: true, port: 8899, version: '0.1.0', reason: '' };
    const spy = vi.spyOn(sidecarMod, 'probeSidecar').mockResolvedValue(status);
    try {
      const r = await runStartup();
      expect(r.sidecar.available).toBe(true);
      expect(getSidecarStatus()).toEqual(status);
    } finally { spy.mockRestore(); }
  });

  it('探不到时状态是不可用，但启动本身照常完成', async () => {
    const spy = vi.spyOn(sidecarMod, 'probeSidecar').mockResolvedValue({
      available: false, port: 0, version: '', reason: '没找到',
    });
    try {
      const r = await runStartup();
      expect(r.sidecar.available).toBe(false);
      expect(getSidecarStatus().available).toBe(false);
      expect(useSettings.getState().provision.status).toBe('ready');
    } finally { spy.mockRestore(); }
  });

  it('探测抛异常也不能把启动打断', async () => {
    const spy = vi.spyOn(sidecarMod, 'probeSidecar').mockRejectedValue(new Error('boom'));
    try {
      const r = await runStartup();
      expect(r.sidecar.available).toBe(false);
      expect(useSettings.getState().provision.status).toBe('ready');
    } finally { spy.mockRestore(); }
  });
});

describe('启动路径：宿主探测的接线', () => {
  beforeEach(() => setHost('unknown'));

  it('host=Excel 时探出来并存进共享状态', async () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { host: 'excel-host-marker', requirements: { isSetSupported: () => true } },
      HostType: { Excel: 'excel-host-marker', PowerPoint: 'ppt-host-marker' },
    };
    const r = await runStartup();
    expect(r.host).toBe('excel');
    expect(getHost()).toBe('excel');
  });

  it('host=PowerPoint 时探出来并存进共享状态', async () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { host: 'ppt-host-marker', requirements: { isSetSupported: () => true } },
      HostType: { Excel: 'excel-host-marker', PowerPoint: 'ppt-host-marker' },
    };
    const r = await runStartup();
    expect(r.host).toBe('powerpoint');
    expect(getHost()).toBe('powerpoint');
  });

  it('没有 Office 对象时宿主是 unknown', async () => {
    delete (globalThis as { Office?: unknown }).Office;
    const r = await runStartup();
    expect(r.host).toBe('unknown');
    expect(getHost()).toBe('unknown');
  });
});
