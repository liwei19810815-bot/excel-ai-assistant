import { z } from 'zod';
import { register } from '../registry';
import { runExcel } from '../../excel/coordinator';
import { resolveSheet, stripSheet } from './read';

/**
 * 图表创建。ChartType 枚举在 ExcelApi 1.1 即可用，
 * 但部分样式属性属于更高版本，这里只用基线能力以保证 Mac/Web 也能跑。
 */
const CHART_TYPES = [
  'ColumnClustered',
  'ColumnStacked',
  'BarClustered',
  'Line',
  'LineMarkers',
  'Pie',
  'Doughnut',
  'XYScatter',
  'Area',
] as const;

register({
  name: 'create_chart',
  description:
    '基于指定数据区域创建图表。dataRange 应包含表头行/列，Excel 会自动识别为系列名与分类轴。' +
    '图表会插入到数据所在工作表。创建后返回图表名称，后续可用于引用。',
  policy: 'mutate:content',
  schema: z.object({
    dataRange: z
      .string()
      .describe('数据区域（含表头），如 "A1:B13" 或 "Sheet1!A1:B13"'),
    sheet: z.string().optional().describe('工作表名，dataRange 中已含表名时可省略'),
    chartType: z.enum(CHART_TYPES).describe('图表类型'),
    title: z.string().optional().describe('图表标题'),
    seriesBy: z
      .enum(['Auto', 'Columns', 'Rows'])
      .optional()
      .default('Auto')
      .describe('系列方向。数据按列组织时用 Columns，按行组织时用 Rows'),
    top: z.number().optional().describe('图表左上角距工作表顶部的距离（磅），默认自动'),
    left: z.number().optional().describe('图表左上角距工作表左侧的距离（磅），默认自动'),
    width: z.number().optional().default(400),
    height: z.number().optional().default(280),
  }),
  summarize: (a) => `创建${a.title ? `「${a.title}」` : ''}图表（${a.chartType}）`,

  async run(args) {
    const addr = stripSheet(args.dataRange);

    const result = await runExcel(async (ctx) => {
      const ws = resolveSheet(ctx, args.sheet, args.dataRange);
      const dataRange = ws.getRange(addr);

      const chart = ws.charts.add(
        args.chartType as Excel.ChartType,
        dataRange,
        args.seriesBy as Excel.ChartSeriesBy,
      );

      if (args.title) {
        chart.title.text = args.title;
        chart.title.visible = true;
      }

      chart.width = args.width;
      chart.height = args.height;

      // 未指定位置时放到数据区右侧，避免盖住数据
      if (args.top !== undefined) chart.top = args.top;
      if (args.left !== undefined) chart.left = args.left;

      chart.load('name');
      await ctx.sync();

      if (args.top === undefined && args.left === undefined) {
        chart.setPosition(dataRange.getOffsetRange(0, 1).getEntireColumn());
        await ctx.sync();
      }

      ws.load('name');
      await ctx.sync();
      return { chartName: chart.name, sheetName: ws.name };
    });

    return {
      text:
        `已在工作表「${result.sheetName}」创建图表「${result.chartName}」` +
        `（类型 ${args.chartType}，数据源 ${addr}）。`,
    };
  },
});
