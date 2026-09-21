import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { ChatPane } from './components/ChatPane';
import { runStartup } from './startup';
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
   * 白名单分流 + 可见性：问一次网关，再按结果把功能区按钮置灰。
   *
   * 【链路本身在 startup.ts 里，这里只负责调】。抽出去是为了能测——
   * main.tsx 要 DOM、要 import 样式，在测试环境里跑不起来，
   * 接线写在这儿就等于没有测试覆盖。
   *
   * 【不 await，但界面不会提前放行】：挂载和问配置并行，
   * 可 provision.status 初始是 pending，ChatPane 在 pending 期间
   * 只显示"正在获取配置"，不渲染聊天界面。
   *
   * 最初的写法是"先挂界面、默认按可用处理"，为的是不让人对着转圈——
   * 但那样会在配置到位之前放行，服务端已经关掉功能的用户
   * 仍能在那几百毫秒里正常发消息，治理开关被绕过。
   */
  void runStartup();
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(() => mount());
} else {
  mount();
}
