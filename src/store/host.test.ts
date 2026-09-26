import { describe, it, expect, beforeEach } from 'vitest';
import { detectHost, setHost, getHost } from './host';

beforeEach(() => {
  setHost('unknown');
  delete (globalThis as { Office?: unknown }).Office;
});

describe('detectHost', () => {
  it('没有 Office 对象时返回 unknown，不抛异常', () => {
    expect(detectHost()).toBe('unknown');
  });

  it('Office.context.host 匹配 Excel 时返回 excel', () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { host: 'x' },
      HostType: { Excel: 'x', PowerPoint: 'y' },
    };
    expect(detectHost()).toBe('excel');
  });

  it('Office.context.host 匹配 PowerPoint 时返回 powerpoint', () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { host: 'y' },
      HostType: { Excel: 'x', PowerPoint: 'y' },
    };
    expect(detectHost()).toBe('powerpoint');
  });

  it('Office.context.host 匹配 Word 时返回 word', () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { host: 'z' },
      HostType: { Excel: 'x', PowerPoint: 'y', Word: 'z' },
    };
    expect(detectHost()).toBe('word');
  });

  it('host 是完全不认识的宿主时返回 unknown，不是乱猜', () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { host: 'q' },
      HostType: { Excel: 'x', PowerPoint: 'y', Word: 'z' },
    };
    expect(detectHost()).toBe('unknown');
  });

  it('Office 对象形状不对（缺 context）时不抛异常，返回 unknown', () => {
    (globalThis as { Office?: unknown }).Office = {};
    expect(detectHost()).toBe('unknown');
  });
});

describe('setHost / getHost', () => {
  it('默认是 unknown', () => {
    expect(getHost()).toBe('unknown');
  });

  it('set 完 get 能读到同一个值', () => {
    setHost('powerpoint');
    expect(getHost()).toBe('powerpoint');
  });
});
