import { QUICK_ACTIONS, QUICK_ACTION_GROUPS, PPT_QUICK_ACTIONS, PPT_QUICK_ACTION_GROUPS } from '../quickActions';
import { getHost } from '../../store/host';

/**
 * 预置提示词面板。
 *
 * 【只填入输入框，不直接发送】。和工具箱搜索框"只带你去看说明、
 * 不执行命令"是同一个取舍：一键就让 AI 动手，省下的那次点击
 * 远不值得冒一次意外修改的风险。用户看一眼措辞、按需要改改再回车。
 *
 * 【会改数据的条目要标出来】。业务人员点之前应该知道这一下是"看看"
 * 还是"要动表"。真正的把关仍然是 AI 的工具分级（改内容自动建快照、
 * 改结构强制确认），这里的标注只是让人心里有数。
 */
export function QuickActions({ onPick }: { onPick: (prompt: string) => void }) {
  // host=unknown 时按 Excel 版兜底，理由同 ChatPane 的标题文案——
  // 这份列表历史上只服务过 Excel。
  const isPpt = getHost() === 'powerpoint';
  const groups = isPpt ? PPT_QUICK_ACTION_GROUPS : QUICK_ACTION_GROUPS;
  const actions = isPpt ? PPT_QUICK_ACTIONS : QUICK_ACTIONS;

  return (
    <div className="space-y-3 text-left">
      {groups.map((group) => {
        const items = actions.filter((a) => a.group === group);
        if (!items.length) return null;
        return (
          <div key={group} className="space-y-1">
            <div className="px-0.5 text-[11px] font-medium text-neutral-400">{group}</div>
            {items.map((a) => (
              <button
                key={a.id}
                onClick={() => onPick(a.prompt)}
                title={a.prompt}
                className="flex w-full items-center gap-1.5 rounded border border-neutral-200 bg-white px-2.5 py-1.5 text-left text-xs text-neutral-700 hover:border-sky-300 hover:text-sky-700"
              >
                <span className="flex-1">{a.label}</span>
                {a.mutates && (
                  <span
                    className="shrink-0 rounded bg-amber-50 px-1 py-0.5 text-[10px] text-amber-700"
                    title={
                      isPpt
                        ? '这一条会改动演示文稿内容。改结构（如新增页）前 AI 会先问你。'
                        : '这一条会改动表格内容。改动可以撤销，改结构前 AI 会先问你。'
                    }
                  >
                    {isPpt ? '会改动' : '会改表'}
                  </span>
                )}
              </button>
            ))}
          </div>
        );
      })}

      <p className="px-0.5 pt-1 text-[11px] leading-relaxed text-neutral-400">
        点一下只是把描述填进输入框，你可以改完再发送。
      </p>
    </div>
  );
}
