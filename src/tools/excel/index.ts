/**
 * Excel 专属工具的名字清单，供按宿主过滤用（PPT AI 那批改动新增）。
 *
 * 【只列 src/tools/excel/* 里的四个文件】。sandbox/runScript 和
 * sidecar/{powerQuery,runMacro} 虽然也是 Excel 专属概念，但它们已经
 * 各自导出过名字数组（没有的话是 run_script，直接按名字加），
 * ChatPane 里按宿主过滤时把这几份名单都塞进 disabled 即可，
 * 不需要在这里重复列一遍——名单散在各自的源头文件里，改一个工具名
 * 只用改一处。
 */
export const EXCEL_TOOL_NAMES = [
  'get_workbook_overview',
  'get_selection',
  'read_range',
  'write_cells',
  'format_cells',
  'modify_structure',
  'create_chart',
] as const;
