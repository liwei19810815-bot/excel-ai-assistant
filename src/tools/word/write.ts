import { z } from 'zod';
import { register } from '../registry';
import { runWord } from '../../word/coordinator';

/**
 * 【为什么没有 checkpoint/撤销快照，以及为什么两个工具都是 mutate:structure】
 *
 * 和 src/tools/powerpoint/write.ts 同一个结论：Office.js 加载项的改动
 * 普遍不进原生撤销栈是有据可查的已知问题（OfficeDev/office-js 仓库
 * issue #5966 等），不是某个宿主的特例。docs/规划-Word与PPT.md 里提到
 * 的"Word 原生 Application.UndoRecord 比 Excel 简单"说的是 VBA 宏
 * （COM 自动化）那条路径，和这里的 Word JS API（任务窗格加载项）是
 * 两回事，不能把 VBA 侧的结论直接套到这里——没有证据支持"Word AI 这几个
 * 工具能进原生撤销栈"，所以老实按不可撤销处理，标 mutate:structure
 * 强制确认，和 Excel/PPT 侧同类工具的分类原则一致。
 */

register({
  name: 'insert_paragraph',
  description:
    '在文档末尾插入一个新段落。这个改动大概率无法用 Ctrl+Z 撤销，' +
    '所以每次调用都需要用户明确确认。',
  policy: 'mutate:structure',
  schema: z.object({
    text: z.string().min(1).describe('要插入的段落文字'),
  }),
  summarize: (a) => `文档末尾新增段落：${a.text.slice(0, 20)}${a.text.length > 20 ? '…' : ''}`,
  async run(args) {
    await runWord(async (ctx) => {
      ctx.document.body.insertParagraph(args.text, Word.InsertLocation.end);
      await ctx.sync();
    });
    return { text: '已在文档末尾插入新段落。' };
  },
});

register({
  name: 'replace_text',
  description:
    '在全文范围内查找并替换文字（区分大小写，精确匹配，不支持正则）。' +
    '会改动文档内容，大概率无法用 Ctrl+Z 撤销，执行前会强制你和用户确认。',
  policy: 'mutate:structure',
  schema: z.object({
    find: z.string().min(1).max(255).describe('要查找的文字，最长 255 字符'),
    replace: z.string().describe('替换成的文字，可以是空字符串（等于删除）'),
  }),
  summarize: (a) => `全文替换「${a.find}」→「${a.replace}」`,
  async run(args) {
    const count = await runWord(async (ctx) => {
      const results = ctx.document.body.search(args.find, { matchCase: true });
      results.load('items');
      await ctx.sync();

      const items = results.items;
      for (const r of items) {
        r.insertText(args.replace, Word.InsertLocation.replace);
      }
      await ctx.sync();
      return items.length;
    });

    if (count === 0) {
      return { text: `全文没有找到「${args.find}」，未做任何改动。` };
    }
    return { text: `已把全文 ${count} 处「${args.find}」替换为「${args.replace}」。` };
  },
});
