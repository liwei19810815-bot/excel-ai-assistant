import { z } from 'zod';
import { register } from '../registry';
import { runPowerPoint, requireApi } from '../../powerpoint/coordinator';
import { buildPptBlueprint, renderPptBlueprint } from '../../powerpoint/blueprint';

register({
  name: 'get_presentation_overview',
  description:
    '获取演示文稿结构：总页数、每页大致内容（近似标题+形状个数）、当前选中页。' +
    '每轮对话已自动注入该信息，仅在你做完结构性改动后需要重新确认时调用。',
  policy: 'read',
  schema: z.object({}),
  summarize: () => '读取演示文稿结构',
  async run() {
    const bp = await buildPptBlueprint();
    return { text: renderPptBlueprint(bp) };
  },
});

register({
  name: 'read_slide',
  description:
    '读取指定页的完整文字内容（按形状顺序列出每个形状里的文字）。' +
    'slideIndex 从 1 开始计数，对应用户口中的"第几页"。',
  policy: 'read',
  schema: z.object({
    slideIndex: z.number().int().min(1).describe('第几页，从 1 开始'),
  }),
  summarize: (a) => `读取第 ${a.slideIndex} 页`,
  async run(args) {
    requireApi('1.4', '读取形状文字');
    const zeroBased = args.slideIndex - 1;

    const result = await runPowerPoint(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items');
      await ctx.sync();

      const slide = slides.items[zeroBased];
      if (!slide) return null;

      const shapes = slide.shapes;
      shapes.load('items');
      await ctx.sync();

      for (const shape of shapes.items) {
        try {
          shape.textFrame.load('hasText');
        } catch {
          // 部分形状类型（如图片）没有 textFrame，忽略
        }
      }
      await ctx.sync();

      for (const shape of shapes.items) {
        try {
          if (shape.textFrame.hasText) shape.textFrame.textRange.load('text');
        } catch {
          // 忽略
        }
      }
      await ctx.sync();

      const texts: string[] = [];
      for (const shape of shapes.items) {
        try {
          if (shape.textFrame.hasText) {
            const t = shape.textFrame.textRange.text;
            if (t) texts.push(t);
          }
        } catch {
          // 忽略
        }
      }
      return { shapeCount: shapes.items.length, texts };
    });

    if (result === null) {
      return { text: `第 ${args.slideIndex} 页不存在，请先用 get_presentation_overview 确认总页数。` };
    }
    if (result.texts.length === 0) {
      return { text: `第 ${args.slideIndex} 页共 ${result.shapeCount} 个形状，没有读到任何文字内容。` };
    }
    return {
      text: `第 ${args.slideIndex} 页（共 ${result.shapeCount} 个形状）：\n\n` +
        result.texts.map((t, i) => `[形状 ${i + 1}]\n${t}`).join('\n\n'),
    };
  },
});
