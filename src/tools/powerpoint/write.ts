import { z } from 'zod';
import { register } from '../registry';
import { runPowerPoint, requireApi } from '../../powerpoint/coordinator';

/**
 * 【这里为什么没有 checkpoint/撤销快照，以及为什么两个工具都是
 * mutate:structure 而不是更宽松的 mutate:content】
 *
 * 一开始以为 Office.js 的改动会进 PowerPoint 原生撤销栈，Ctrl+Z 直接管用，
 * 不需要在 JS 层再建一套平行快照——这个假设是错的，查证后发现两点：
 *
 * 1. Office.js 加载项改动普遍不进原生撤销栈是有据可查的已知问题
 *    （OfficeDev/office-js 仓库 issue #5966 等），不是 Excel 特例。
 * 2. 这个仓库自己的 UI 已经在替 Excel 侧说这件事了——
 *    ChatPane.tsx 底部固定提示"插件的改动无法用 Ctrl+Z 撤销，
 *    请用操作卡片上的「撤销」"，Excel 侧建
 *    src/excel/checkpoint.ts 那套自定义快照系统，根子上就是因为
 *    原生撤销靠不住，不只是 sidecar VBA 宏那一条路径的问题。
 *
 * 所以老实的结论是：这两个工具**很可能也无法用 Ctrl+Z 撤销**，而且
 * 本轮没有给 PowerPoint 建对应的自定义快照系统（PPT 没有 Excel
 * Range.Copy 那种现成的"整块拷贝备份"能力，值得单独设计，不是这轮
 * 的范围）。唯一诚实的安全边界就是**执行前强制确认**——所以两个工具
 * 都标成 mutate:structure（不是 mutate:content，那个策略隐含"会自动
 * 建快照、可撤销"的承诺，这里给不出这个承诺），和 VBA 侧 PPT 工具箱
 * 已经定下的"PPT 全部不可撤销，所以确认要更强"是同一个结论
 * （见 docs/规划-Word与PPT.md）。
 */

const positionSchema = {
  left: z.number().optional().describe('距幻灯片左边的距离（磅），不填用默认位置'),
  top: z.number().optional().describe('距幻灯片顶部的距离（磅），不填用默认位置'),
  width: z.number().optional().describe('宽度（磅），不填用默认宽度'),
  height: z.number().optional().describe('高度（磅），不填用默认高度'),
};

register({
  name: 'add_text_box',
  description:
    '在指定页新增一个文本框。slideIndex 从 1 开始计数。' +
    '位置/尺寸参数都可省略，省略时由 PowerPoint 给默认值。' +
    '这个改动大概率无法用 Ctrl+Z 撤销，所以每次调用都需要用户明确确认。',
  policy: 'mutate:structure',
  schema: z.object({
    slideIndex: z.number().int().min(1).describe('第几页，从 1 开始'),
    text: z.string().min(1).describe('文本框内容'),
    ...positionSchema,
  }),
  summarize: (a) => `第 ${a.slideIndex} 页新增文本框`,
  async run(args) {
    requireApi('1.4', '新增文本框');
    const zeroBased = args.slideIndex - 1;

    const result = await runPowerPoint(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items');
      await ctx.sync();

      const slide = slides.items[zeroBased];
      if (!slide) return null;

      const shape = slide.shapes.addTextBox(args.text, {
        left: args.left,
        top: args.top,
        width: args.width,
        height: args.height,
      });
      shape.load('id');
      await ctx.sync();
      return { shapeId: shape.id };
    });

    if (result === null) {
      return { text: `第 ${args.slideIndex} 页不存在，请先用 get_presentation_overview 确认总页数。` };
    }
    return { text: `已在第 ${args.slideIndex} 页新增文本框（内容：${args.text}）。` };
  },
});

register({
  name: 'add_slide',
  description:
    '新增一页幻灯片，追加在演示文稿末尾，并可选地放一个标题文本框和一个正文文本框。' +
    '这会改变总页数和页面顺序，而且大概率无法用 Ctrl+Z 撤销，' +
    '执行前会强制你和用户确认。',
  policy: 'mutate:structure',
  schema: z.object({
    title: z.string().optional().describe('标题文字，省略则不加标题文本框'),
    bullets: z
      .array(z.string())
      .optional()
      .describe('正文要点，每条一行，省略则不加正文文本框'),
  }),
  summarize: (a) => `新增幻灯片${a.title ? `「${a.title}」` : ''}`,
  async run(args) {
    requireApi('1.4', '新增幻灯片并放置文本框');

    const result = await runPowerPoint(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items');
      await ctx.sync();
      const countBefore = slides.items.length;

      slides.add();
      await ctx.sync();

      // add() 不返回新幻灯片对象，重新取集合按数量定位新增的那一页——
      // add() 只会追加到末尾，新页面必然是最后一个
      const slidesAfter = ctx.presentation.slides;
      slidesAfter.load('items');
      await ctx.sync();
      const newSlide = slidesAfter.items[countBefore];
      if (!newSlide) throw new Error('新增幻灯片后读不到新页面，插入可能失败了。');

      if (args.title) {
        newSlide.shapes.addTextBox(args.title, { left: 40, top: 30, width: 640, height: 60 });
      }
      if (args.bullets?.length) {
        newSlide.shapes.addTextBox(args.bullets.join('\n'), {
          left: 40, top: 110, width: 640, height: 340,
        });
      }
      await ctx.sync();

      return { newIndex: countBefore + 1 };
    });

    return { text: `已在第 ${result.newIndex} 页（末尾）新增幻灯片。` };
  },
});
