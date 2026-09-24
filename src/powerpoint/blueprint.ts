import { runPowerPoint, isSupported } from './coordinator';

/**
 * 演示文稿「蓝图」——只含结构摘要，不含全文，每轮对话自动注入。
 * 和 src/excel/blueprint.ts 是同一个理由：让模型不用先调工具就知道
 * 这份演示文稿长什么样。
 */

export interface SlideSummary {
  /** 0-based，和 slides.getItemAt 的下标一致 */
  index: number;
  /** 取第一个有文字的形状的首行文字作为近似标题——不是正式的
   *  Title 占位符判定（那需要 PlaceholderFormat，PowerPointApi 1.8），
   *  只是给模型一个大致印象，工具描述里会写清楚这是"近似"。 */
  titleGuess: string;
  shapeCount: number;
}

export interface PptBlueprint {
  presentationTitle: string;
  slideCount: number;
  slides: SlideSummary[];
  /** 当前选中的幻灯片下标；探测不到（PowerPointApi < 1.5）时为 null */
  activeSlideIndex: number | null;
}

export async function buildPptBlueprint(): Promise<PptBlueprint> {
  return runPowerPoint(async (ctx) => {
    const pres = ctx.presentation;
    pres.load('title');

    const slides = pres.slides;
    slides.load('items');

    await ctx.sync();

    const shapesPerSlide = slides.items.map((s) => {
      const shapes = s.shapes;
      shapes.load('items');
      return shapes;
    });

    await ctx.sync();

    // 每页取前 3 个形状的文字，拼出一个"大概讲了什么"的猜测标题——
    // 只加载少量形状的 textFrame，避免一份几十页的演示文稿把上下文撑爆
    const textLoads = shapesPerSlide.map((shapes) => {
      const top = shapes.items.slice(0, 3);
      for (const shape of top) {
        try {
          shape.textFrame.load('hasText');
        } catch {
          // 某些形状类型不支持 textFrame（如图片），忽略
        }
      }
      return top;
    });

    await ctx.sync();

    const titleGuessTexts = textLoads.map((shapes) => {
      for (const shape of shapes) {
        try {
          if (shape.textFrame.hasText) shape.textFrame.textRange.load('text');
        } catch {
          // 忽略不支持 textFrame 的形状
        }
      }
      return shapes;
    });

    await ctx.sync();

    const slideSummaries: SlideSummary[] = slides.items.map((_s, i) => {
      const shapes = titleGuessTexts[i];
      let titleGuess = '';
      for (const shape of shapes) {
        try {
          if (shape.textFrame.hasText) {
            const text = shape.textFrame.textRange.text ?? '';
            const firstLine = text.split('\n')[0]?.trim();
            if (firstLine) { titleGuess = firstLine; break; }
          }
        } catch {
          // 忽略
        }
      }
      return { index: i, titleGuess, shapeCount: shapesPerSlide[i].items.length };
    });

    // 当前选中页需要 PowerPointApi 1.5，基线只声明了 1.1，探测不到就老实说 null，
    // 不要瞎猜——猜错了模型会对着错误的那一页操作。
    let activeSlideIndex: number | null = null;
    if (isSupported('1.5')) {
      try {
        const selected = pres.getSelectedSlides();
        selected.load('items');
        await ctx.sync();
        const first = selected.items[0];
        if (first) {
          first.load('id');
          await ctx.sync();
          activeSlideIndex = slides.items.findIndex((s) => s.id === first.id);
          if (activeSlideIndex < 0) activeSlideIndex = null;
        }
      } catch {
        activeSlideIndex = null;
      }
    }

    return {
      presentationTitle: pres.title ?? '',
      slideCount: slides.items.length,
      slides: slideSummaries,
      activeSlideIndex,
    };
  });
}

/** 蓝图 → 紧凑文本，直接拼进每轮的上下文块 */
export function renderPptBlueprint(bp: PptBlueprint): string {
  const lines: string[] = ['## 当前演示文稿结构'];

  if (bp.presentationTitle) lines.push(`标题：${bp.presentationTitle}`);
  lines.push(`共 ${bp.slideCount} 页`);

  if (bp.activeSlideIndex != null) {
    lines.push(`当前选中第 ${bp.activeSlideIndex + 1} 页`);
  } else {
    lines.push('（当前选中页无法探测，需要时请主动调用工具查看具体某一页）');
  }

  for (const s of bp.slides) {
    const mark = s.index === bp.activeSlideIndex ? '（当前）' : '';
    const guess = s.titleGuess ? `「${s.titleGuess}」` : '（空白或无法读取文字）';
    lines.push(`- 第 ${s.index + 1} 页${mark}：${guess}，共 ${s.shapeCount} 个形状`);
  }

  return lines.join('\n');
}
