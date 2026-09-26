import { z } from 'zod';
import { register } from '../registry';
import { runWord } from '../../word/coordinator';
import { buildWordBlueprint, renderWordBlueprint } from '../../word/blueprint';

register({
  name: 'get_document_overview',
  description:
    '获取文档结构：总段数、前几段预览、当前选区预览。' +
    '每轮对话已自动注入该信息，仅在你做完结构性改动后需要重新确认时调用。',
  policy: 'read',
  schema: z.object({}),
  summarize: () => '读取文档结构',
  async run() {
    const bp = await buildWordBlueprint();
    return { text: renderWordBlueprint(bp) };
  },
});

register({
  name: 'read_paragraph',
  description:
    '读取指定段落的完整文字内容。paragraphIndex 从 1 开始计数，' +
    '对应用户口中的"第几段"。',
  policy: 'read',
  schema: z.object({
    paragraphIndex: z.number().int().min(1).describe('第几段，从 1 开始'),
  }),
  summarize: (a) => `读取第 ${a.paragraphIndex} 段`,
  async run(args) {
    const zeroBased = args.paragraphIndex - 1;

    const result = await runWord(async (ctx) => {
      const paragraphs = ctx.document.body.paragraphs;
      paragraphs.load('items');
      await ctx.sync();

      const p = paragraphs.items[zeroBased];
      if (!p) return null;

      p.load('text');
      await ctx.sync();
      return { text: p.text ?? '' };
    });

    if (result === null) {
      return { text: `第 ${args.paragraphIndex} 段不存在，请先用 get_document_overview 确认总段数。` };
    }
    return { text: result.text || '（空段落）' };
  },
});
