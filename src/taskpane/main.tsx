import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { ChatPane } from './components/ChatPane';
import { provision, applyRibbonVisibility } from '../store/provisioning';
import '../styles.css';
// 导入即注册全部工具
import '../tools';

/**
 * Office.onReady 必须在访问任何 Office API 之前 resolve。
 * 在非 Office 宿主（纯浏览器调试）中 Office 对象不存在，此时直接挂载，
 * UI 可以正常预览，只是 Excel 相关工具会在调用时报错。
 */
function mount() {
  const root = document.getElementById('root');
  if (!root) throw new Error('找不到挂载点 #root');
  createRoot(root).render(
    <StrictMode>
      <ChatPane />
    </StrictMode>,
  );

  /**
   * 白名单分流 + 可见性：问一次网关。
   *
   * 【这里不 await，但界面不会提前放行】。
   * 挂载和问配置是并行的，可 `provision.status` 初始是 `pending`，
   * ChatPane 在 pending 期间只显示"正在获取配置"，不渲染聊天界面。
   *
   * 最初的写法是"先挂界面、默认按可用处理"，为的是不让人对着转圈——
   * 但那样会在配置到位之前放行，服务端已经关掉功能的用户
   * 仍能在那几百毫秒里正常发消息，治理开关被绕过。
   *
   * provision 内部把所有失败都收敛成 byok 并置 ready，
   * 所以这里不需要 catch，也不会卡在 pending。
   */
  void provision().then((r) => {
    // 按服务端下发的可见性把按钮置灰。尽力而为，失败无所谓——
    // 真正的强制在 ChatPane 里（停用时根本不渲染聊天界面）。
    void applyRibbonVisibility(r.visibility);
  });
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(() => mount());
} else {
  mount();
}
