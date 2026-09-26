import { runWord } from './coordinator';

/**
 * 文档「蓝图」——只含结构摘要，不含全文，每轮对话自动注入。
 * 和 src/excel/blueprint.ts / src/powerpoint/blueprint.ts 是同一个理由：
 * 让模型不用先调工具就知道这份文档长什么样。
 */

const PREVIEW_COUNT = 8;
const PREVIEW_LEN = 60;

export interface ParagraphPreview {
  /** 1-based，和 read_paragraph 工具的 paragraphIndex 参数对齐 */
  index: number;
  textPreview: string;
}

export interface WordBlueprint {
  paragraphCount: number;
  /** 只取前 PREVIEW_COUNT 段的预览，避免长文档把上下文撑爆 */
  previews: ParagraphPreview[];
  /** 当前选区的文字预览，选区为空或读取失败时为空字符串 */
  selectionPreview: string;
}

function truncate(text: string): string {
  const t = text.trim();
  return t.length > PREVIEW_LEN ? t.slice(0, PREVIEW_LEN) + '…' : t;
}

export async function buildWordBlueprint(): Promise<WordBlueprint> {
  return runWord(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('items');
    await ctx.sync();

    const top = paragraphs.items.slice(0, PREVIEW_COUNT);
    for (const p of top) p.load('text');
    await ctx.sync();

    const previews: ParagraphPreview[] = top.map((p, i) => ({
      index: i + 1,
      textPreview: truncate(p.text ?? ''),
    }));

    let selectionPreview = '';
    try {
      const selection = ctx.document.getSelection();
      selection.load('text');
      await ctx.sync();
      selectionPreview = truncate(selection.text ?? '');
    } catch {
      selectionPreview = '';
    }

    return {
      paragraphCount: paragraphs.items.length,
      previews,
      selectionPreview,
    };
  });
}

/** 蓝图 → 紧凑文本，直接拼进每轮的上下文块 */
export function renderWordBlueprint(bp: WordBlueprint): string {
  const lines: string[] = ['## 当前文档结构'];

  lines.push(`共 ${bp.paragraphCount} 段`);

  if (bp.selectionPreview) {
    lines.push(`当前选区预览：${bp.selectionPreview}`);
  } else {
    lines.push('（当前无选区，或选区为空）');
  }

  for (const p of bp.previews) {
    lines.push(`- 第 ${p.index} 段：${p.textPreview || '（空段落）'}`);
  }
  if (bp.paragraphCount > bp.previews.length) {
    lines.push(`（仅预览前 ${bp.previews.length} 段，其余请用 read_paragraph 按需查看）`);
  }

  return lines.join('\n');
}
